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


def register_presentation_tools(app: FastMCP, presentations: Dict, get_current_presentation_id, get_template_search_directories):
    """Register presentation management tools with the FastMCP app"""
    
    @app.tool(
        annotations=ToolAnnotations(
            title="Create Presentation",
        ),
    )
    def create_presentation(id: Optional[str] = None) -> Dict:
        """Create a new PowerPoint presentation."""
        # Create a new presentation
        pres = ppt_utils.create_presentation()
        
        # Generate an ID if not provided
        if id is None:
            id = f"presentation_{len(presentations) + 1}"
        
        # Store the presentation
        presentations[id] = pres
        # Set as current presentation (this would need to be handled by caller)
        
        return {
            "presentation_id": id,
            "message": f"Created new presentation with ID: {id}",
            "slide_count": len(pres.slides)
        }

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
    def create_presentation_from_template(template_path: str, id: Optional[str] = None) -> Dict:
        """Create a new PowerPoint presentation from a template file."""
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
        
        # Generate an ID if not provided
        if id is None:
            id = f"presentation_{len(presentations) + 1}"
        
        # Store the presentation
        presentations[id] = pres
        
        return {
            "presentation_id": id,
            "message": f"Created new presentation from template '{template_path}' with ID: {id}",
            "template_path": template_path,
            "slide_count": len(pres.slides),
            "layout_count": len(pres.slide_layouts)
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
        """Save the presentation to S3 and return a presigned download URL.

        Replaces the upstream `save_presentation(file_path, ...)` tool. The
        model no longer chooses a filesystem path; the server serializes the
        in-memory deck to an S3 object and returns a short-lived download
        URL the caller can share with the end user.

        Configuration via env vars (Lambda task role provides AWS creds):
          - PPTX_OUTPUT_BUCKET (required)
          - PRESIGNED_URL_TTL_SECONDS (optional, default 1800)
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

        try:
            ttl = int(os.environ.get("PRESIGNED_URL_TTL_SECONDS", "1800"))
        except ValueError:
            ttl = 1800

        try:
            buffer = io.BytesIO()
            presentations[pres_id].save(buffer)
            buffer.seek(0)

            timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
            key = f"pptx/{timestamp}-{uuid.uuid4()}.pptx"

            client = _get_s3_client()
            client.put_object(
                Bucket=bucket,
                Key=key,
                Body=buffer.getvalue(),
                ContentType=(
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                ),
            )

            url = client.generate_presigned_url(
                ClientMethod="get_object",
                Params={"Bucket": bucket, "Key": key},
                ExpiresIn=ttl,
            )

            return {
                "message": "Presentation uploaded.",
                "download_url": url,
                "expires_in_seconds": ttl,
                "s3_key": key,
            }
        except Exception as e:
            return {
                "error": f"Failed to upload presentation: {str(e)}"
            }

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