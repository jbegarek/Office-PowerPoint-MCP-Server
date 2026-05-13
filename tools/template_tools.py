"""
Enhanced template-based slide creation tools for PowerPoint MCP Server.
Handles template application, template management, automated slide generation,
and advanced features like dynamic sizing, auto-wrapping, and visual effects.
"""
from typing import Dict, List, Optional, Any
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
import utils.template_utils as template_utils


def register_template_tools(app: FastMCP, presentations: Dict, get_current_presentation_id):
    """Register template-based tools with the FastMCP app"""
    
    @app.tool(
        annotations=ToolAnnotations(
            title="List Slide Templates",
            readOnlyHint=True,
        ),
    )
    def list_slide_templates() -> Dict:
        """List all available slide layout templates."""
        try:
            available_templates = template_utils.get_available_templates()
            usage_examples = template_utils.get_template_usage_examples()
            
            return {
                "available_templates": available_templates,
                "total_templates": len(available_templates),
                "usage_examples": usage_examples,
                "message": "Use apply_slide_template to apply templates to slides"
            }
        except Exception as e:
            return {
                "error": f"Failed to list templates: {str(e)}"
            }
    
    @app.tool(
        annotations=ToolAnnotations(
            title="Apply Slide Template",
        ),
    )
    def apply_slide_template(
        slide_index: int,
        template_id: str,
        color_scheme: str = "modern_blue",
        content_mapping: Optional[Dict[str, str]] = None,
        image_paths: Optional[Dict[str, str]] = None,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """Apply a built-in layout template to an existing slide. Use when you have already created a slide and want to restyle it. For creating a new slide with a template, prefer create_slide_from_template. Call list_slide_templates to see valid template_id values. image_paths is not supported — file uploads are not available."""
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        if slide_index < 0 or slide_index >= len(pres.slides):
            return {
                "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
            }
        
        slide = pres.slides[slide_index]
        
        try:
            result = template_utils.apply_slide_template(
                slide, template_id, color_scheme, 
                content_mapping or {}, image_paths or {}
            )
            
            if result['success']:
                return {
                    "message": f"Applied template '{template_id}' to slide {slide_index}",
                    "slide_index": slide_index,
                    "template_applied": result
                }
            else:
                return {
                    "error": f"Failed to apply template: {result.get('error', 'Unknown error')}"
                }
                
        except Exception as e:
            return {
                "error": f"Failed to apply template: {str(e)}"
            }
    
    @app.tool(
        annotations=ToolAnnotations(
            title="Create Slide from Template",
        ),
    )
    def create_slide_from_template(
        template_id: str,
        color_scheme: str = "modern_blue",
        content_mapping: Optional[Dict[str, str]] = None,
        image_paths: Optional[Dict[str, str]] = None,
        layout_index: int = 1,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """PREFERRED: Create a new slide using a built-in layout template. Call list_slide_templates first to see valid template_id values. content_mapping maps element roles (e.g. "title", "content") to text. image_paths is not supported. When all slides are done, call save_presentation and return the download_url to the user."""
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        # Validate layout index
        if layout_index < 0 or layout_index >= len(pres.slide_layouts):
            return {
                "error": f"Invalid layout index: {layout_index}. Available layouts: 0-{len(pres.slide_layouts) - 1}"
            }
        
        try:
            # Add new slide
            layout = pres.slide_layouts[layout_index]
            slide = pres.slides.add_slide(layout)
            slide_index = len(pres.slides) - 1
            
            # Apply template
            result = template_utils.apply_slide_template(
                slide, template_id, color_scheme,
                content_mapping or {}, image_paths or {}
            )
            
            if result['success']:
                return {
                    "message": f"Created slide {slide_index} using template '{template_id}'",
                    "slide_index": slide_index,
                    "template_applied": result
                }
            else:
                return {
                    "error": f"Failed to apply template to new slide: {result.get('error', 'Unknown error')}"
                }
                
        except Exception as e:
            return {
                "error": f"Failed to create slide from template: {str(e)}"
            }
    
    @app.tool(
        annotations=ToolAnnotations(
            title="Create Presentation from Templates",
        ),
    )
    def create_presentation_from_templates(
        template_sequence: List[Dict[str, Any]],
        color_scheme: str = "modern_blue",
        presentation_title: Optional[str] = None,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """Create a full presentation from a sequence of template slides in one call. template_sequence is a list of objects, each with "template_id" and "content" keys (content maps role names like "title" / "content" to text). Call list_slide_templates first to see valid template IDs. image_paths inside templates are not supported. After creation, call save_presentation and return the download_url to the user."""
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        if not template_sequence:
            return {
                "error": "Template sequence cannot be empty"
            }
        
        try:
            # Set presentation title if provided
            if presentation_title:
                pres.core_properties.title = presentation_title
            
            # Create slides from template sequence
            result = template_utils.create_presentation_from_template_sequence(
                pres, template_sequence, color_scheme
            )
            
            if result['success']:
                return {
                    "message": f"Created presentation with {result['total_slides']} slides",
                    "presentation_id": pres_id,
                    "creation_result": result,
                    "total_slides": len(pres.slides)
                }
            else:
                return {
                    "warning": "Presentation created with some errors",
                    "presentation_id": pres_id,
                    "creation_result": result,
                    "total_slides": len(pres.slides)
                }
                
        except Exception as e:
            return {
                "error": f"Failed to create presentation from templates: {str(e)}"
            }
    
    @app.tool(
        annotations=ToolAnnotations(
            title="Get Template Info",
            readOnlyHint=True,
        ),
    )
    def get_template_info(template_id: str) -> Dict:
        """Get detailed information about a specific template: its elements, available color schemes, and usage tip. Use list_slide_templates first to find valid template_id values."""
        try:
            templates_data = template_utils.load_slide_templates()
            
            if template_id not in templates_data.get('templates', {}):
                available_templates = list(templates_data.get('templates', {}).keys())
                return {
                    "error": f"Template '{template_id}' not found",
                    "available_templates": available_templates
                }
            
            template = templates_data['templates'][template_id]
            
            # Extract element information
            elements_info = []
            for element in template.get('elements', []):
                element_info = {
                    "type": element.get('type'),
                    "role": element.get('role'),
                    "position": element.get('position'),
                    "placeholder_text": element.get('placeholder_text', ''),
                    "styling_options": list(element.get('styling', {}).keys())
                }
                elements_info.append(element_info)
            
            return {
                "template_id": template_id,
                "name": template.get('name'),
                "description": template.get('description'),
                "layout_type": template.get('layout_type'),
                "elements": elements_info,
                "element_count": len(elements_info),
                "has_background": 'background' in template,
                "background_type": template.get('background', {}).get('type'),
                "color_schemes": list(templates_data.get('color_schemes', {}).keys()),
                "usage_tip": f"Use create_slide_from_template with template_id='{template_id}' to create a slide with this layout"
            }
            
        except Exception as e:
            return {
                "error": f"Failed to get template info: {str(e)}"
            }
    
    # Text optimization tools
    
    
    @app.tool(
        annotations=ToolAnnotations(
            title="Optimize Slide Text",
        ),
    )
    def optimize_slide_text(
        slide_index: int,
        auto_resize: bool = True,
        auto_wrap: bool = True,
        optimize_spacing: bool = True,
        min_font_size: int = 8,
        max_font_size: int = 36,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """Automatically resize fonts, wrap text, and adjust spacing on a slide so text fits its containers. slide_index is 0-based. font sizes are clamped between min_font_size and max_font_size (defaults: 8–36pt)."""
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        
        pres = presentations[pres_id]
        
        if slide_index < 0 or slide_index >= len(pres.slides):
            return {
                "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
            }
        
        slide = pres.slides[slide_index]
        
        try:
            optimizations_applied = []
            manager = template_utils.get_enhanced_template_manager()
            
            # Analyze each text shape on the slide
            for i, shape in enumerate(slide.shapes):
                if hasattr(shape, 'text_frame') and shape.text_frame.text:
                    text = shape.text_frame.text
                    
                    # Calculate container dimensions
                    container_width = shape.width.inches
                    container_height = shape.height.inches
                    
                    shape_optimizations = []
                    
                    # Apply auto-resize if enabled
                    if auto_resize:
                        optimal_size = template_utils.calculate_dynamic_font_size(
                            text, container_width, container_height
                        )
                        optimal_size = max(min_font_size, min(max_font_size, optimal_size))
                        
                        # Apply the calculated font size
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = template_utils.Pt(optimal_size)
                        
                        shape_optimizations.append(f"Font resized to {optimal_size}pt")
                    
                    # Apply auto-wrap if enabled
                    if auto_wrap:
                        current_font_size = 14  # Default assumption
                        if shape.text_frame.paragraphs and shape.text_frame.paragraphs[0].runs:
                            if shape.text_frame.paragraphs[0].runs[0].font.size:
                                current_font_size = shape.text_frame.paragraphs[0].runs[0].font.size.pt
                        
                        wrapped_text = template_utils.wrap_text_automatically(
                            text, container_width, current_font_size
                        )
                        
                        if wrapped_text != text:
                            shape.text_frame.text = wrapped_text
                            shape_optimizations.append("Text wrapped automatically")
                    
                    # Optimize spacing if enabled
                    if optimize_spacing:
                        text_length = len(text)
                        if text_length > 300:
                            line_spacing = 1.4
                        elif text_length > 150:
                            line_spacing = 1.3
                        else:
                            line_spacing = 1.2
                        
                        for paragraph in shape.text_frame.paragraphs:
                            paragraph.line_spacing = line_spacing
                        
                        shape_optimizations.append(f"Line spacing set to {line_spacing}")
                    
                    if shape_optimizations:
                        optimizations_applied.append({
                            "shape_index": i,
                            "optimizations": shape_optimizations
                        })
            
            return {
                "message": f"Optimized {len(optimizations_applied)} text elements on slide {slide_index}",
                "slide_index": slide_index,
                "optimizations_applied": optimizations_applied,
                "settings": {
                    "auto_resize": auto_resize,
                    "auto_wrap": auto_wrap,
                    "optimize_spacing": optimize_spacing,
                    "font_size_range": f"{min_font_size}-{max_font_size}pt"
                }
            }
            
        except Exception as e:
            return {
                "error": f"Failed to optimize slide text: {str(e)}"
            }