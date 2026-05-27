"""Bizzdesign-tailored composition tools.

Each composition wraps a specific source slide from the bundled brand
template (`Bizzdesign_CorporateDeck_2026.pptx`). The model picks a
composition by name and hands over content as named fields. The server
clones the source slide's shape tree into a freshly-added slide on the
working presentation and text-replaces each tagged shape with the
supplied content.

This is the deliberate alternative to letting the model assemble shapes
from primitives. The model can't see what it builds, so we ship pre-built
compositions proven by the template's own example slides — the same
visual quality bar the brand designer set, with the model just supplying
text.
"""
from copy import deepcopy
from typing import Any, Dict, List, Optional
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pptx import Presentation


# ------------------------------------------------------------------
# Composition catalogue
# ------------------------------------------------------------------
# Each composition declares:
#   source_slide_index: 0-based index into the brand template (None = no clone, uses layout directly)
#   description: short blurb (returned by list_compositions)
#   use_when: when to pick this composition (returned to model)
#   fields: dict of field_name -> { "match": <prefix>, optional "required": bool }
#           For repeatable groups (e.g. details, pillars), the field value is a list of matchers.
#
# Matchers: we locate shapes by the leading text content in the template,
# because (a) it's stable across deepcopies, (b) shape names are reused
# across shapes in the template (e.g. "Text Placeholder 6" appears more
# than once on a single slide), and (c) it's the most semantically clear.

COMPOSITIONS: Dict[str, Dict[str, Any]] = {
    "cover": {
        "source_slide_index": 0,
        "description": "Opening cover slide: title + subtitle on the brand cover background.",
        "use_when": "First slide of every deck.",
        "fields": {
            "title": {"match": "Corporate Deck", "required": True},
            "subtitle": {"match": "2026", "required": True},
        },
    },
    "section_divider": {
        "source_slide_index": 4,
        "description": "Section break: small kicker text above a large title.",
        "use_when": "Between major sections of a deck (rare in 7-slide decks; useful in longer ones).",
        "fields": {
            "kicker": {"match": "What We Offer", "required": True},
            "title": {"match": "Transformation That Flows", "required": True},
        },
    },
    "value_props_4": {
        "source_slide_index": 8,
        "description": "Lead claim + 4 value-proposition pillars + subtitle. Right-side stack of 4 short pillars.",
        "use_when": "4 parallel differentiators or value propositions. Don't use for fewer than 4.",
        "fields": {
            "title": {"match": "Why", "required": True},
            "lead_claim": {"match": "Only true end-to-end", "required": True},
            "subtitle": {"match": "The combined resources", "required": True},
            "pillars": [
                {"match": "Proven Market Leadership"},
                {"match": "Trusted by 2,000+ Enterprises"},
                {"match": "Continuous AI-driven innovation"},
                {"match": "Global reach, local expertise"},
            ],
        },
    },
    "solution_detail": {
        "source_slide_index": 13,
        "description": "The workhorse: category kicker + offering title + benefits paragraph + 3 detail items.",
        "use_when": "Any substantial topic with 3 detail items and a benefits summary. Use for each priority phase or each platform pillar in a roadmap deck.",
        "fields": {
            "category": {"match": "Transformation Planning", "required": True},
            "title": {"match": "Strategic Portfolio Management", "required": True},
            "benefits": {"match": "Align investments with strategic goals", "required": True},
            "details": [
                {"match": "Business Strategy Alignment"},
                {"match": "IT Investment Optimization"},
                {"match": "Strategic"},  # matches "Strategic Roadmapping" (and others — see locator)
            ],
        },
    },
}


# ------------------------------------------------------------------
# Locator + shape walking
# ------------------------------------------------------------------

def _shape_text(shape) -> str:
    """Concatenate all text runs in a shape's text frame (single string)."""
    if not shape.has_text_frame:
        return ""
    parts = []
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.text:
                parts.append(run.text)
    return "".join(parts)


def _set_shape_text(shape, new_text: str) -> None:
    """Replace a shape's entire text content, preserving the styling of the
    FIRST run in the FIRST paragraph (font, size, colour, bold, etc.).

    python-pptx-friendly approach: keep the first paragraph + first run,
    overwrite its text, delete any subsequent runs / paragraphs.
    """
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    # Keep first paragraph; delete the rest.
    paragraphs = tf.paragraphs
    if not paragraphs:
        tf.text = new_text
        return
    first_para = paragraphs[0]
    # Delete all paragraphs after the first by manipulating XML.
    for para in list(paragraphs[1:]):
        para._p.getparent().remove(para._p)
    # Within the first paragraph, keep only the first run.
    runs = first_para.runs
    if not runs:
        first_para.text = new_text
        return
    first_run = runs[0]
    for run in list(runs[1:]):
        run._r.getparent().remove(run._r)
    first_run.text = new_text


def _set_shape_multiline(shape, lines: List[str]) -> None:
    """Replace a shape's text with multiple lines (one paragraph each),
    preserving the first paragraph's styling for every line.
    """
    if not shape.has_text_frame or not lines:
        _set_shape_text(shape, "")
        return
    # First line goes through _set_shape_text (preserves first run styling).
    _set_shape_text(shape, lines[0])
    if len(lines) == 1:
        return
    tf = shape.text_frame
    # Clone the first paragraph for each additional line.
    first_p = tf.paragraphs[0]._p
    for line in lines[1:]:
        new_p = deepcopy(first_p)
        # Clear all run text in the clone, then set the first run's text.
        from lxml import etree
        a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        runs = new_p.findall(f"{{{a_ns}}}r")
        if not runs:
            new_p.text = line
        else:
            # Drop extra runs in the clone, leaving just the first.
            for r in runs[1:]:
                new_p.remove(r)
            # Update the surviving run's <a:t>
            t = runs[0].find(f"{{{a_ns}}}t")
            if t is not None:
                t.text = line
        first_p.addnext(new_p) if False else tf._txBody.append(new_p)


def _find_shape_by_match(shapes, match_prefix: str):
    """Return the first shape whose text starts with the given prefix
    (case-sensitive, stripped). None if not found.
    """
    target = match_prefix.strip()
    for shape in shapes:
        text = _shape_text(shape).strip()
        if text.startswith(target):
            return shape
    return None


def _clear_shape_text(shape) -> None:
    """Empty a shape's text frame (used when a repeatable field has fewer
    entries than the template provides).
    """
    if shape.has_text_frame:
        _set_shape_text(shape, "")


# ------------------------------------------------------------------
# Slide cloning
# ------------------------------------------------------------------

def _clone_slide_into(working_pres, library_pres, source_index: int):
    """Append a new slide to `working_pres` that is a deep clone of
    `library_pres.slides[source_index]`. Returns the new Slide.

    The clone uses the same slide layout (matched by name from the
    library's layout). All shapes from the source slide are deep-copied
    into the new slide's shape tree.
    """
    src_slide = library_pres.slides[source_index]
    src_layout = src_slide.slide_layout
    src_layout_name = src_layout.name

    # Find the matching layout in the working presentation by name.
    target_layout = None
    for layout in working_pres.slide_layouts:
        if layout.name == src_layout_name:
            target_layout = layout
            break
    if target_layout is None:
        # Fallback: use the first layout. Shouldn't happen because both
        # presentations were loaded from the same template.
        target_layout = working_pres.slide_layouts[0]

    new_slide = working_pres.slides.add_slide(target_layout)

    # Strip the default placeholders that add_slide inserted — they would
    # clash with the placeholders coming from the cloned source.
    for ph in list(new_slide.placeholders):
        sp = ph._element
        sp.getparent().remove(sp)

    # Deep-copy each shape from source into the new slide's shape tree.
    # Skip:
    #   - {nv,}GrpSpPr (group props at the top of spTree)
    #   - <p:pic> (pictures — their <a:blip r:embed="rIdN"/> refs point to
    #     relationships on the SOURCE slide that we don't replicate, so on
    #     the clone they resolve to "missing image" empty boxes. Branded
    #     text shapes carry the styling; decorative imagery is a Phase 2
    #     concern that needs proper rel + media copying.)
    src_tree = src_slide.shapes._spTree
    new_tree = new_slide.shapes._spTree
    for child in list(src_tree):
        tag = child.tag
        if tag.endswith("}nvGrpSpPr") or tag.endswith("}grpSpPr"):
            continue
        if tag.endswith("}pic"):
            continue
        new_tree.append(deepcopy(child))

    return new_slide


# ------------------------------------------------------------------
# Field application
# ------------------------------------------------------------------

def _apply_fields(slide, fields_spec: Dict[str, Any], content: Dict[str, Any]) -> List[str]:
    """Apply text replacements to the slide's shapes per the composition's
    fields spec.

    `fields_spec` maps field_name → matcher dict OR list of matcher dicts (repeatable).
    `content` maps field_name → string OR list of strings.

    Returns a list of warnings (missing shapes, etc).
    """
    warnings: List[str] = []
    shapes = list(slide.shapes)

    for field_name, spec in fields_spec.items():
        value = content.get(field_name)
        if isinstance(spec, list):
            # Repeatable group.
            values = value or []
            if not isinstance(values, list):
                warnings.append(f"Field '{field_name}' expected a list, got {type(values).__name__}")
                values = []
            for i, matcher in enumerate(spec):
                shape = _find_shape_by_match(shapes, matcher["match"])
                if shape is None:
                    warnings.append(f"Could not locate shape for {field_name}[{i}] (match prefix: {matcher['match']!r})")
                    continue
                if i < len(values):
                    new_text = values[i]
                    if isinstance(new_text, str):
                        # Detect newlines → multiline
                        if "\n" in new_text:
                            _set_shape_multiline(shape, new_text.split("\n"))
                        else:
                            _set_shape_text(shape, new_text)
                    else:
                        _set_shape_text(shape, str(new_text))
                else:
                    # User supplied fewer values than the template has slots —
                    # blank the unused shape.
                    _clear_shape_text(shape)
        else:
            # Scalar field.
            if value is None:
                if spec.get("required", False):
                    warnings.append(f"Required field '{field_name}' missing from content")
                continue
            shape = _find_shape_by_match(shapes, spec["match"])
            if shape is None:
                warnings.append(f"Could not locate shape for '{field_name}' (match prefix: {spec['match']!r})")
                continue
            if isinstance(value, list):
                _set_shape_multiline(shape, [str(v) for v in value])
            else:
                if "\n" in str(value):
                    _set_shape_multiline(shape, str(value).split("\n"))
                else:
                    _set_shape_text(shape, str(value))

    return warnings


# ------------------------------------------------------------------
# Library cache
# ------------------------------------------------------------------
# We need a fresh Presentation instance per working deck (so the model's
# changes don't leak into the library). Reload from disk per session.

_LIBRARY_TEMPLATE_CACHE: Dict[str, str] = {}  # presentation_id → template path used


def _get_library(presentation_id: str, library_template_paths: Dict[str, str]):
    """Re-open the library template from disk to get a fresh copy."""
    path = library_template_paths.get(presentation_id)
    if path is None:
        return None
    return Presentation(path)


# ------------------------------------------------------------------
# Tool registration
# ------------------------------------------------------------------

def register_composition_tools(
    app: FastMCP,
    presentations: Dict,
    get_current_presentation_id,
    library_template_paths: Dict[str, str],
):
    """Register the composition tools with the FastMCP app.

    `library_template_paths` is a presentation_id → template path dict
    maintained by `create_presentation`. It lets us reload the library
    from disk per composition call.
    """

    def _build_composition(
        composition_name: str,
        content: Dict[str, Any],
        presentation_id: Optional[str],
    ) -> Dict[str, Any]:
        comp = COMPOSITIONS.get(composition_name)
        if comp is None:
            return {"error": f"Unknown composition: {composition_name!r}"}

        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        if pres_id is None or pres_id not in presentations:
            return {"error": "No presentation is currently loaded."}

        working = presentations[pres_id]
        library = _get_library(pres_id, library_template_paths)
        if library is None:
            return {
                "error": (
                    "No brand template library available for this presentation. "
                    "Compositions require a presentation created via the default-template "
                    "path (PPTX_DEFAULT_TEMPLATE) — `blank=True` decks can't use compositions."
                )
            }

        source_idx = comp["source_slide_index"]
        try:
            new_slide = _clone_slide_into(working, library, source_idx)
            warnings = _apply_fields(new_slide, comp["fields"], content)
        except Exception as e:
            return {"error": f"Failed to build {composition_name}: {e}"}

        slide_index = len(working.slides) - 1
        result: Dict[str, Any] = {
            "composition": composition_name,
            "slide_index": slide_index,
            "message": f"Added '{composition_name}' as slide {slide_index}.",
        }
        if warnings:
            result["warnings"] = warnings
        return result

    @app.tool(
        annotations=ToolAnnotations(
            title="List Compositions",
            readOnlyHint=True,
        ),
    )
    def list_compositions() -> Dict:
        """List the high-level slide compositions available in the brand template.

        Each composition is backed by a known source slide from the bundled
        corporate template. Pick a composition by matching content shape to
        `use_when`. Compositions handle layout, placeholders, fonts, and
        styling for you — you just supply the content.
        """
        return {
            "compositions": {
                name: {
                    "description": meta["description"],
                    "use_when": meta["use_when"],
                    "fields": {
                        fname: ("(repeatable list of strings)" if isinstance(fspec, list)
                                else "(string)")
                        for fname, fspec in meta["fields"].items()
                    },
                }
                for name, meta in COMPOSITIONS.items()
            },
            "guidance": (
                "Always prefer these compositions over `add_slide` for branded output. "
                "Each composition clones a proven branded slide from the template — the "
                "model only supplies content. For repeatable fields (e.g. `details`, "
                "`pillars`), supply a list of strings; if shorter than the template "
                "expects, unused slots are blanked. Multi-line content can include `\\n`."
            ),
        }

    @app.tool(
        annotations=ToolAnnotations(title="Add Cover Slide"),
    )
    def add_cover_slide(
        title: str,
        subtitle: str,
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add the deck's opening cover slide with brand-styled title and subtitle.

        Use this for slide 1 of every deck. The cover uses the dark
        Bizzdesign brand background with the large title centred and a
        subtitle line below.

        Args:
            title: The deck's headline (claim form, e.g. "From feature to platform").
            subtitle: Subtitle text (audience, year, or short scope statement).
        """
        return _build_composition("cover", {"title": title, "subtitle": subtitle}, presentation_id)

    @app.tool(
        annotations=ToolAnnotations(title="Add Section Divider"),
    )
    def add_section_divider(
        kicker: str,
        title: str,
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a section divider slide — a small kicker label above a large title.

        Use this between major sections of a deck. For 7-slide decks this is
        rare; for 12+ slide decks it gives the audience a visual breath.

        Args:
            kicker: Short label above the title (e.g. "What We Offer", "Where We Go Next").
            title: Section title (claim form).
        """
        return _build_composition("section_divider", {"kicker": kicker, "title": title}, presentation_id)

    @app.tool(
        annotations=ToolAnnotations(title="Add Value Props (4 Pillars)"),
    )
    def add_value_props_slide(
        title: str,
        lead_claim: str,
        subtitle: str,
        pillars: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add a 4-pillar value-props slide with a lead claim and subtitle.

        Right side: a stack of 4 short pillar statements (each ~5-8 words).
        Left side: the deck's lead claim with a smaller subtitle beneath.

        Use this when you have exactly 4 parallel differentiators or value
        propositions. Don't use for 3 or fewer — pick `solution_detail` or
        `bullets` instead.

        Args:
            title: Main title (e.g. "Why Bizzdesign?").
            lead_claim: One-line lead claim (e.g. "Only true end-to-end Enterprise Transformation offering").
            subtitle: Supporting subtitle.
            pillars: List of exactly 4 short pillar statements.
        """
        return _build_composition(
            "value_props_4",
            {"title": title, "lead_claim": lead_claim, "subtitle": subtitle, "pillars": pillars},
            presentation_id,
        )

    @app.tool(
        annotations=ToolAnnotations(title="Add Solution Detail"),
    )
    def add_solution_detail_slide(
        category: str,
        title: str,
        benefits: str,
        details: List[str],
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Add the workhorse "solution detail" composition.

        Use this for any substantial topic with 3 detail items. The slide
        has:
        - A small kicker (`category`) at the top
        - A large product/area title beneath
        - A multi-line benefits paragraph (the "why this matters")
        - 3 detail items in a row at the bottom (each: short heading + description)

        For a roadmap or platform deck, use this for EACH phase or EACH
        pillar — it's the brand's standard "topic detail" template.

        Args:
            category: Top-of-slide kicker (e.g. "Transformation Planning", "Now — Q2 2026").
            title: Main heading (the topic name, ~3-5 words).
            benefits: Multi-line benefits paragraph. Use `\\n` between lines.
            details: List of 3 detail items, each formatted as
                "Heading\\nDescription text spanning one or two lines". The first
                line becomes the bold heading; subsequent lines are the body.
        """
        return _build_composition(
            "solution_detail",
            {"category": category, "title": title, "benefits": benefits, "details": details},
            presentation_id,
        )
