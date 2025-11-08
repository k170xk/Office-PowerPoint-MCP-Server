"""
Presentation management tools for PowerPoint MCP Server.
Handles presentation creation, opening, saving, and core properties.
"""
from typing import Dict, List, Optional, Any
import os
from mcp.server.fastmcp import FastMCP
import utils as ppt_utils


def register_presentation_tools(
    app: FastMCP,
    presentations: Dict,
    get_current_presentation_id,
    get_template_search_directories,
    set_current_presentation_id=None
):
    """Register presentation management tools with the FastMCP app"""
    
    @app.tool()
    def create_presentation(
        id: Optional[str] = None,
        file_path: Optional[str] = None,
        title: Optional[str] = None,
        subtitle: Optional[str] = None,
        slide_layout_index: int = 0,
        set_as_current: bool = True,
        auto_save: bool = True
    ) -> Dict:
        """
        Create a new PowerPoint presentation.

        Args:
            id: Optional presentation identifier (auto-generated when omitted)
            file_path: Optional path to immediately save the presentation
            title: Optional title text for the first slide
            subtitle: Optional subtitle text for the first slide
            slide_layout_index: Slide layout to use when creating a title slide
            set_as_current: Whether to mark the created presentation as current
            auto_save: Save to file_path (when provided) before returning
        """
        # Create a new presentation
        pres = ppt_utils.create_presentation()
        
        # Generate an ID if not provided
        if id is None:
            id = f"presentation_{len(presentations) + 1}"
        
        # Store the presentation
        presentations[id] = pres

        # Optionally set as current presentation
        if set_as_current and callable(set_current_presentation_id):
            try:
                set_current_presentation_id(id)
            except Exception:
                pass

        # Optionally add a title slide with provided metadata
        title_applied = False
        subtitle_applied = False
        if title or subtitle:
            try:
                slide, _ = ppt_utils.add_slide(pres, slide_layout_index)
                if title:
                    try:
                        ppt_utils.set_title(slide, title)
                        title_applied = True
                    except Exception:
                        pass
                if subtitle:
                    try:
                        # Commonly the subtitle placeholder has index 1
                        ppt_utils.populate_placeholder(slide, 1, subtitle)
                        subtitle_applied = True
                    except Exception:
                        # Attempt to find a placeholder with matching name as fallback
                        for placeholder in getattr(slide, "placeholders", []):
                            try:
                                if "subtitle" in placeholder.name.lower():
                                    placeholder.text = subtitle
                                    subtitle_applied = True
                                    break
                            except Exception:
                                continue
            except Exception:
                pass

        saved_path = None
        # Auto-save the presentation if requested and a path is provided
        if auto_save and file_path:
            try:
                os.makedirs(os.path.dirname(file_path), exist_ok=True)
            except (FileNotFoundError, TypeError):
                # Directory portion may be empty (current directory) or invalid
                pass
            try:
                saved_path = ppt_utils.save_presentation(pres, file_path)
            except Exception as exc:
                saved_path = None
                save_error = str(exc)
            else:
                save_error = None
        else:
            save_error = None
        
        result: Dict[str, Any] = {
            "presentation_id": id,
            "message": f"Created new presentation with ID: {id}",
            "slide_count": len(pres.slides),
            "title_applied": title_applied,
            "subtitle_applied": subtitle_applied
        }

        if saved_path:
            result["file_path"] = saved_path
            result["saved"] = True
        else:
            result["saved"] = False

        if save_error:
            result["save_error"] = save_error

        # Include echo of provided metadata for transparency
        if title:
            result["title"] = title
        if subtitle:
            result["subtitle"] = subtitle

        return result

    @app.tool()
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

    @app.tool()
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

    @app.tool()
    def save_presentation(file_path: str, presentation_id: Optional[str] = None) -> Dict:
        """Save a presentation to a file."""
        # Use the specified presentation or the current one
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        # Save the presentation
        try:
            saved_path = ppt_utils.save_presentation(presentations[pres_id], file_path)
            return {
                "message": f"Presentation saved to {saved_path}",
                "file_path": saved_path
            }
        except Exception as e:
            return {
                "error": f"Failed to save presentation: {str(e)}"
            }

    @app.tool()
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

    @app.tool()
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

    @app.tool()
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