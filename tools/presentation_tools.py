"""
Presentation management tools for PowerPoint MCP Server.
Handles presentation creation, opening, saving, and core properties.
"""
from typing import Dict, List, Optional, Any
import io
import os
import uuid
from datetime import datetime, timezone
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
import utils as ppt_utils

# Bizzdesign fork: lazy boto3 import so local stdio/HTTP runs (without AWS
# credentials) don't fail at module load. The S3 upload tool only imports
# when invoked.
_S3_CLIENT = None


def _purge_orphan_slide_parts(pptx_bytes: bytes) -> bytes:
    """Drop slide parts that are no longer referenced by `presentation.xml.rels`.

    python-pptx's `drop_rel` removes a slide from the in-memory rel graph
    and `sldIdLst`, but the underlying `SlidePart` object stays registered
    on the Package. On save it's still serialised as `ppt/slides/slideN.xml`
    even though nothing reaches it. PowerPoint's strict validator flags
    these orphan parts on open ("this file has problems, do you want to
    repair it?") AND the orphan slides occasionally surface in the Outline
    pane / file recovery flow as zombie content like "Thank You" or
    "Questions?" left over from `delete_slide` calls during build.

    This post-save pass:
      1. Reads the legitimate slide target list from `presentation.xml.rels`.
      2. Walks every `ppt/slides/slideN.xml` + matching `_rels/slideN.xml.rels`
         in the zip and removes any whose filename isn't in the legitimate set.
      3. Strips the corresponding `<Override PartName="/ppt/slides/slideN.xml">`
         entries from `[Content_Types].xml` so the package validator stays
         consistent.

    Pure zip-level rewrite — no python-pptx state involved. Returns the
    cleaned bytes.
    """
    import zipfile
    import re
    from io import BytesIO

    src_buf = BytesIO(pptx_bytes)
    with zipfile.ZipFile(src_buf, "r") as src:
        names = src.namelist()
        try:
            pres_rels = src.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
        except KeyError:
            # Unexpected shape — return unchanged rather than corrupt the file.
            return pptx_bytes

        legitimate = set()
        for match in re.finditer(r'Target="slides/(slide\d+\.xml)"', pres_rels):
            legitimate.add(match.group(1))

        slide_path_re = re.compile(r"^ppt/slides/(slide\d+\.xml)$")
        slide_rels_re = re.compile(r"^ppt/slides/_rels/(slide\d+\.xml)\.rels$")

        orphans: set = set()
        for name in names:
            m = slide_path_re.match(name)
            if m and m.group(1) not in legitimate:
                orphans.add(name)
                continue
            m = slide_rels_re.match(name)
            if m and m.group(1) not in legitimate:
                orphans.add(name)

        if not orphans:
            return pptx_bytes

        # Build the override-prune set: every orphan slide's content-type
        # Override entry must come out of [Content_Types].xml too.
        override_targets = {f"/{name}" for name in orphans if slide_path_re.match(name)}

        out_buf = BytesIO()
        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as dst:
            for name in names:
                if name in orphans:
                    continue
                data = src.read(name)
                if name == "[Content_Types].xml" and override_targets:
                    text = data.decode("utf-8")
                    for target in override_targets:
                        # Match the full <Override .../> element with that PartName.
                        text = re.sub(
                            rf'<Override[^/]*PartName="{re.escape(target)}"[^/]*/>',
                            "",
                            text,
                        )
                    data = text.encode("utf-8")
                dst.writestr(name, data)
        return out_buf.getvalue()


def _strip_template_slides(pres) -> int:
    """Remove every slide from a Presentation while keeping the slide
    masters and layouts (which carry the brand styling) intact.

    Drops both the slide ID entry from sldIdLst and the corresponding
    relationship on the presentation part, so PowerPoint's strict
    validator doesn't surface "this file has problems, do you want to
    repair it?" on open. (Just removing from sldIdLst leaves orphan
    rels behind, which is what PowerPoint complains about.)

    Returns the number of slides removed.
    """
    xml_slides = pres.slides._sldIdLst
    slide_entries = list(xml_slides)
    removed = 0
    for slide_entry in slide_entries:
        rId = slide_entry.rId
        pres.part.drop_rel(rId)
        xml_slides.remove(slide_entry)
        removed += 1
    return removed


def _get_s3_client():
    global _S3_CLIENT
    if _S3_CLIENT is None:
        import boto3
        from botocore.config import Config

        # Force the regional S3 endpoint + SigV4 so presigned GETs against
        # buckets outside us-east-1 don't come back as 307
        # TemporaryRedirect XML errors. `region_name` alone is not enough
        # in recent boto3 — virtual-host URLs still hit the global host
        # unless we also pass an explicit `endpoint_url`. AWS_REGION is
        # set automatically in Lambda; AWS_DEFAULT_REGION is the
        # local-dev fallback.
        region = os.environ.get("AWS_REGION") or os.environ.get(
            "AWS_DEFAULT_REGION"
        )
        if region:
            _S3_CLIENT = boto3.client(
                "s3",
                region_name=region,
                endpoint_url=f"https://s3.{region}.amazonaws.com",
                config=Config(signature_version="s3v4"),
            )
        else:
            _S3_CLIENT = boto3.client("s3")
    return _S3_CLIENT


def register_presentation_tools(app: FastMCP, presentations: Dict, get_current_presentation_id, get_template_search_directories, library_template_paths: Optional[Dict[str, str]] = None, set_current_presentation_id=None):
    """Register presentation management tools with the FastMCP app"""
    
    @app.tool(
        annotations=ToolAnnotations(
            title="Create Presentation",
        ),
    )
    def create_presentation(
        id: Optional[str] = None,
        blank: bool = False,
        include_template_slides: bool = False,
    ) -> Dict:
        """Create a new PowerPoint presentation.

        Bizzdesign fork: when the `PPTX_DEFAULT_TEMPLATE` env var is set
        on the server, this tool seeds the new presentation from that
        bundled template by default — the slide masters and slide
        layouts (which carry the brand styling) are imported, but the
        template's example slides are stripped so the assistant starts
        with a clean 0-slide deck and adds new slides using the
        branded layouts via `add_slide(layout_index=...)`.

        Pass `blank=True` to opt out entirely (no template imported).
        Pass `include_template_slides=True` to keep the template's
        example slides as the starting point (useful if you want to
        modify existing content in place).

        If the env-var template can't be found in the configured search
        directories, the tool falls back to a blank presentation rather
        than failing — operators see the warning in the returned
        `template_warning` field.
        """
        default_template = os.environ.get("PPTX_DEFAULT_TEMPLATE", "").strip()
        template_warning: Optional[str] = None
        resolved_template_path: Optional[str] = None

        if not blank and default_template:
            # Same lookup as `create_presentation_from_template`: accept an
            # absolute path verbatim, otherwise scan the configured search
            # directories for a basename match.
            if os.path.exists(default_template):
                resolved_template_path = default_template
            else:
                template_name = os.path.basename(default_template)
                for directory in get_template_search_directories():
                    candidate = os.path.join(directory, template_name)
                    if os.path.exists(candidate):
                        resolved_template_path = candidate
                        break

            if resolved_template_path is None:
                template_warning = (
                    f"PPTX_DEFAULT_TEMPLATE='{default_template}' not found in "
                    f"search dirs {get_template_search_directories()}; using a "
                    "blank presentation instead."
                )

        stripped_slide_count = 0
        if resolved_template_path is not None:
            try:
                pres = ppt_utils.create_presentation_from_template(resolved_template_path)
                if not include_template_slides:
                    stripped_slide_count = _strip_template_slides(pres)
            except Exception as e:
                pres = ppt_utils.create_presentation()
                template_warning = (
                    f"Failed to load PPTX_DEFAULT_TEMPLATE '{resolved_template_path}': "
                    f"{e}. Using a blank presentation instead."
                )
                resolved_template_path = None
        else:
            pres = ppt_utils.create_presentation()

        if id is None:
            id = f"presentation_{len(presentations) + 1}"

        presentations[id] = pres
        # Make this newly-created presentation the current one. Without
        # this, the server's global `current_presentation_id` keeps
        # pointing at whatever presentation was most recently opened or
        # created by an earlier session — and Lambda concurrency=1 means
        # every MCP session shares that global. When the model omits
        # `presentation_id` in a downstream tool call, slides land in
        # the wrong presentation, producing multi-brief Frankenstein
        # decks. (Discovered via ralph iter 13: llm-legacy's deck came
        # back with 27 slides — 7 econ + 9 co-obj + 11 llm-legacy.)
        if set_current_presentation_id is not None:
            set_current_presentation_id(id)
        # Record the source template path so composition tools can re-open
        # the library as a fresh Presentation per call.
        if library_template_paths is not None and resolved_template_path is not None:
            library_template_paths[id] = resolved_template_path

        # Surface the slide layouts so the model knows which layout_index
        # to pass to `add_slide` for each slide type. Without this info
        # the model picks layout indices blindly and the resulting deck
        # doesn't reflect the template's brand styling.
        layouts_info = [
            {"index": i, "name": layout.name}
            for i, layout in enumerate(pres.slide_layouts)
        ]

        result: Dict[str, Any] = {
            "presentation_id": id,
            "slide_count": len(pres.slides),
            "layouts": layouts_info,
        }
        if resolved_template_path is not None:
            if stripped_slide_count:
                result["message"] = (
                    f"Created presentation '{id}' from bundled template "
                    f"'{resolved_template_path}'. Starts with 0 slides. The "
                    f"template's {stripped_slide_count} example slides were "
                    f"stripped from the working deck but are available as a "
                    f"COMPOSITION LIBRARY — prefer the high-level "
                    f"`add_cover_slide`, `add_section_divider`, "
                    f"`add_value_props_slide`, `add_solution_detail_slide` "
                    f"tools (see `list_compositions`) over the low-level "
                    f"`add_slide` + `add_bullet_points` combo. Compositions "
                    f"clone proven branded slides from the template and "
                    f"only need you to supply content."
                )
            else:
                result["message"] = (
                    f"Created presentation '{id}' from bundled template "
                    f"'{resolved_template_path}' with its {len(pres.slides)} "
                    f"example slides retained. Modify them in place to keep "
                    f"the brand look."
                )
            result["template_path"] = resolved_template_path
        else:
            result["message"] = f"Created new (blank) presentation '{id}'."
        if template_warning is not None:
            result["template_warning"] = template_warning
        return result

    @app.tool(
        annotations=ToolAnnotations(
            title="List Bundled Template Files",
            readOnlyHint=True,
        ),
    )
    def list_template_files() -> Dict:
        """List `.pptx` / `.potx` template files bundled with the server.

        Bizzdesign fork: the container image ships brand template decks
        baked into `/app/templates`. Operators add new templates by
        copying them into that directory at build time. The returned
        `filename` values can be passed directly to
        `create_presentation_from_template`.

        For purely layout templates (no .pptx file), see
        `list_slide_templates` which returns the JSON layouts.
        """
        search_dirs = get_template_search_directories()
        templates: List[Dict[str, str]] = []
        seen = set()
        for directory in search_dirs:
            if not os.path.isdir(directory):
                continue
            for filename in sorted(os.listdir(directory)):
                if not filename.lower().endswith(('.pptx', '.potx')):
                    continue
                if filename in seen:
                    continue
                seen.add(filename)
                full_path = os.path.join(directory, filename)
                try:
                    size_bytes = os.path.getsize(full_path)
                except OSError:
                    size_bytes = 0
                templates.append({
                    "filename": filename,
                    "directory": directory,
                    "size_bytes": size_bytes,
                })
        return {
            "templates": templates,
            "search_directories": search_dirs,
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Create Presentation from Template",
        ),
    )
    def create_presentation_from_template(
        template_path: str,
        id: Optional[str] = None,
        include_template_slides: bool = False,
    ) -> Dict:
        """Create a new PowerPoint presentation from a template file.

        By default the template's slide masters and layouts are imported
        (carrying the brand styling) but the example slides are stripped
        so the assistant starts with 0 slides and builds the deck using
        `add_slide(layout_index=...)`. Pass `include_template_slides=True`
        to keep the example slides as the starting point if you want to
        modify existing content in place.
        """
        # Check if template file exists
        if not os.path.exists(template_path):
            # Try to find the template by searching in configured directories
            search_dirs = get_template_search_directories()
            template_name = os.path.basename(template_path)
            
            for directory in search_dirs:
                potential_path = os.path.join(directory, template_name)
                if os.path.exists(potential_path):
                    template_path = potential_path
                    break
            else:
                env_path_info = f" (PPT_TEMPLATE_PATH: {os.environ.get('PPT_TEMPLATE_PATH', 'not set')})" if os.environ.get('PPT_TEMPLATE_PATH') else ""
                return {
                    "error": f"Template file not found: {template_path}. Searched in {', '.join(search_dirs)}{env_path_info}"
                }
        
        # Create presentation from template
        try:
            pres = ppt_utils.create_presentation_from_template(template_path)
        except Exception as e:
            return {
                "error": f"Failed to create presentation from template: {str(e)}"
            }

        stripped = 0
        if not include_template_slides:
            stripped = _strip_template_slides(pres)

        # Generate an ID if not provided
        if id is None:
            id = f"presentation_{len(presentations) + 1}"

        # Store the presentation
        presentations[id] = pres
        if set_current_presentation_id is not None:
            set_current_presentation_id(id)
        if library_template_paths is not None:
            library_template_paths[id] = template_path

        layouts_info = [
            {"index": i, "name": layout.name}
            for i, layout in enumerate(pres.slide_layouts)
        ]

        if stripped:
            message = (
                f"Created presentation '{id}' from template '{template_path}'. "
                f"Starts with 0 slides — the template's {stripped} example slides "
                f"were stripped, but the brand layouts (see `layouts`) and slide "
                f"masters are preserved. Build the deck by calling "
                f"`add_slide(layout_index=...)` with the layout whose name "
                f"matches each slide's role."
            )
        else:
            message = (
                f"Created presentation '{id}' from template '{template_path}' "
                f"with its {len(pres.slides)} example slides retained. Modify "
                f"them in place to keep the brand look."
            )

        return {
            "presentation_id": id,
            "message": message,
            "template_path": template_path,
            "slide_count": len(pres.slides),
            "layouts": layouts_info,
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Open Presentation",
            readOnlyHint=True,
        ),
    )
    def open_presentation(file_path: str, id: Optional[str] = None) -> Dict:
        """Open an existing PowerPoint presentation from a file."""
        # Check if file exists
        if not os.path.exists(file_path):
            return {
                "error": f"File not found: {file_path}"
            }
        
        # Open the presentation
        try:
            pres = ppt_utils.open_presentation(file_path)
        except Exception as e:
            return {
                "error": f"Failed to open presentation: {str(e)}"
            }
        
        # Generate an ID if not provided
        if id is None:
            id = f"presentation_{len(presentations) + 1}"
        
        # Store the presentation
        presentations[id] = pres
        if set_current_presentation_id is not None:
            set_current_presentation_id(id)

        return {
            "presentation_id": id,
            "message": f"Opened presentation from {file_path} with ID: {id}",
            "slide_count": len(pres.slides)
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Save Presentation To Download URL",
            destructiveHint=False,
        ),
    )
    def save_presentation_to_url(presentation_id: Optional[str] = None) -> Dict:
        """Save the presentation to S3 and return a download URL.

        Replaces the upstream `save_presentation(file_path, ...)` tool. The
        server serializes the in-memory deck to S3 and returns a plain
        `https://s3.<region>.amazonaws.com/<bucket>/<key>` URL. The bucket
        policy grants public-read on the `pptx/` prefix; the object key
        uses a UUIDv4 suffix (~2^122 entropy) so the URL is effectively
        unguessable, and a 7-day S3 lifecycle bounds exposure.

        Configuration via env vars:
          - PPTX_OUTPUT_BUCKET (required)
          - AWS_REGION (set automatically in Lambda; local-dev fallback
            is AWS_DEFAULT_REGION).
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()

        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        bucket = os.environ.get("PPTX_OUTPUT_BUCKET")
        if not bucket:
            return {
                "error": "PPTX_OUTPUT_BUCKET env var is not set on the MCP server."
            }

        region = os.environ.get("AWS_REGION") or os.environ.get(
            "AWS_DEFAULT_REGION"
        ) or "us-east-1"

        try:
            buffer = io.BytesIO()
            presentations[pres_id].save(buffer)
            cleaned = _purge_orphan_slide_parts(buffer.getvalue())

            timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
            key = f"pptx/{timestamp}-{uuid.uuid4()}.pptx"

            client = _get_s3_client()
            client.put_object(
                Bucket=bucket,
                Key=key,
                Body=cleaned,
                ContentType=(
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                ),
            )

            # Plain public-read URL. The bucket policy (defined in CDK)
            # allows `s3:GetObject` on `pptx/*` from any principal.
            url = f"https://s3.{region}.amazonaws.com/{bucket}/{key}"

            return {
                "message": "Presentation uploaded.",
                "download_url": url,
                "s3_key": key,
            }
        except Exception as e:
            return {
                "error": f"Failed to upload presentation: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Render Deck to Images",
            readOnlyHint=True,
        ),
    )
    def render_deck_to_images(presentation_id: Optional[str] = None) -> Dict:
        """Render the current presentation to one PNG per slide, upload
        the PNGs to S3, and return a map of slide_index → public PNG URL.

        Use this AFTER validate_deck passes and BEFORE
        save_presentation_to_url. You — the AI assistant — should fetch
        each PNG URL into your context (you have multimodal vision) and
        visually inspect each slide. Look for:
          - text overflow / orphan words on titles
          - empty slots / unfilled card columns
          - inappropriate template imagery bleeding through
          - font-size inconsistency across same-role shapes
          - low-contrast text or any layout artifact

        If you spot a visual issue, fix it with `delete_slide(<idx>)` +
        a fresh composition call, then re-call this tool to confirm.
        Only call save_presentation_to_url when every slide looks right.

        Returns:
          slides: list of {slide_index, png_url} ordered by slide.
          message: human-readable summary.
          error / setup_required: present if rendering isn't supported
            in the running environment (e.g. soffice not installed).
        """
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        bucket = os.environ.get("PPTX_OUTPUT_BUCKET")
        if not bucket:
            return {
                "error": "PPTX_OUTPUT_BUCKET env var is not set on the MCP server."
            }

        # Locate soffice. Lambda image installs it via apt (libreoffice-core
        # + libreoffice-impress); skip the rendering path with a clear
        # `setup_required` flag if missing.
        import shutil
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        pdftoppm = shutil.which("pdftoppm")
        if not soffice or not pdftoppm:
            return {
                "error": "Rendering tools not installed on this MCP server.",
                "setup_required": {
                    "missing": [n for n, p in [("soffice", soffice), ("pdftoppm", pdftoppm)] if not p],
                    "fix": (
                        "Rebuild the MCP server Docker image with `apt-get install "
                        "libreoffice-core libreoffice-impress poppler-utils`. "
                        "Local dev: `brew install --cask libreoffice && brew install poppler`."
                    ),
                },
            }

        region = os.environ.get("AWS_REGION") or os.environ.get("AWS_DEFAULT_REGION") or "us-east-1"
        client = _get_s3_client()

        import subprocess
        import tempfile
        import glob

        try:
            with tempfile.TemporaryDirectory() as tmp:
                # 1. Save the current pres to a temp .pptx (with orphan-slide
                # cleanup applied so the soffice render matches what users
                # will see when they open the saved deck).
                pptx_path = os.path.join(tmp, "deck.pptx")
                save_buf = io.BytesIO()
                presentations[pres_id].save(save_buf)
                cleaned = _purge_orphan_slide_parts(save_buf.getvalue())
                with open(pptx_path, "wb") as f:
                    f.write(cleaned)

                # 2. soffice → pdf.
                pdf_dir = os.path.join(tmp, "pdf")
                os.makedirs(pdf_dir, exist_ok=True)
                # HOME needs to be writable for soffice user-profile init.
                env = os.environ.copy()
                env.setdefault("HOME", tmp)
                subprocess.run(
                    [soffice, "--headless", "--convert-to", "pdf", "--outdir", pdf_dir, pptx_path],
                    check=True, env=env, capture_output=True, timeout=120,
                )
                pdf_path = os.path.join(pdf_dir, "deck.pdf")
                if not os.path.exists(pdf_path):
                    return {"error": f"soffice didn't produce a PDF at {pdf_path}"}

                # 3. pdftoppm → png per page.
                png_dir = os.path.join(tmp, "pngs")
                os.makedirs(png_dir, exist_ok=True)
                subprocess.run(
                    [pdftoppm, "-png", "-r", "100", pdf_path, os.path.join(png_dir, "slide")],
                    check=True, timeout=60,
                )

                # 4. Upload each PNG to S3.
                pngs = sorted(glob.glob(os.path.join(png_dir, "slide-*.png")))
                timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
                upload_id = uuid.uuid4()
                results = []
                for i, png_path in enumerate(pngs):
                    key = f"pptx-renders/{timestamp}-{upload_id}-slide-{i + 1:03d}.png"
                    with open(png_path, "rb") as fh:
                        client.put_object(
                            Bucket=bucket,
                            Key=key,
                            Body=fh.read(),
                            ContentType="image/png",
                        )
                    results.append({
                        "slide_index": i,
                        "png_url": f"https://s3.{region}.amazonaws.com/{bucket}/{key}",
                    })

                return {
                    "slides": results,
                    "message": (
                        f"Rendered {len(results)} slides to PNG and uploaded to S3. "
                        "FETCH each URL into your context and visually inspect — "
                        "look for text overflow, orphan words, empty card/column "
                        "slots, inappropriate template imagery, font inconsistency, "
                        "low contrast. Fix visual issues with delete_slide + a "
                        "fresh composition call, then re-render to confirm."
                    ),
                }
        except subprocess.CalledProcessError as e:
            return {"error": f"Render command failed: {e.stderr.decode('utf-8', 'replace') if e.stderr else e}"}
        except Exception as e:
            return {"error": f"Render failed: {e!r}"}

    @app.tool(
        annotations=ToolAnnotations(
            title="Get Presentation Info",
            readOnlyHint=True,
        ),
    )
    def get_presentation_info(presentation_id: Optional[str] = None) -> Dict:
        """Get information about a presentation."""
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        try:
            info = ppt_utils.get_presentation_info(pres)
            info["presentation_id"] = pres_id
            return info
        except Exception as e:
            return {
                "error": f"Failed to get presentation info: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Get Template File Info",
            readOnlyHint=True,
        ),
    )
    def get_template_file_info(template_path: str) -> Dict:
        """Get information about a template file including layouts and properties."""
        # Check if template file exists
        if not os.path.exists(template_path):
            # Try to find the template by searching in configured directories
            search_dirs = get_template_search_directories()
            template_name = os.path.basename(template_path)
            
            for directory in search_dirs:
                potential_path = os.path.join(directory, template_name)
                if os.path.exists(potential_path):
                    template_path = potential_path
                    break
            else:
                return {
                    "error": f"Template file not found: {template_path}. Searched in {', '.join(search_dirs)}"
                }
        
        try:
            return ppt_utils.get_template_info(template_path)
        except Exception as e:
            return {
                "error": f"Failed to get template info: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Set Core Properties",
        ),
    )
    def set_core_properties(
        title: Optional[str] = None,
        subject: Optional[str] = None,
        author: Optional[str] = None,
        keywords: Optional[str] = None,
        comments: Optional[str] = None,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """Set core document properties."""
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        try:
            ppt_utils.set_core_properties(
                pres,
                title=title,
                subject=subject,
                author=author,
                keywords=keywords,
                comments=comments
            )
            
            return {
                "message": "Core properties updated successfully"
            }
        except Exception as e:
            return {
                "error": f"Failed to set core properties: {str(e)}"
            }