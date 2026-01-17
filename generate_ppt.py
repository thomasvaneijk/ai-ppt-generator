#!/usr/bin/env python3
"""
PowerPoint Generator MVP

Generates professional PowerPoint presentations from JSON slide models.
Follows strict architecture: JSON → Python → PowerPoint.
"""

import json
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER


# Debug mode: Print placeholder details for troubleshooting
DEBUG = True


# Layout indices in template.pptx
LAYOUT_INDICES = {
    "title": 0,
    "key_insights": 1,
    "chart": 2,
    "two_column": 3,
    "recommendation": 4,
    "appendix": 5,
}


# ============================================================================
# Placeholder Detection Helpers
# ============================================================================

def debug_print_placeholders(slide, slide_type: str):
    """Print all placeholder information for debugging."""
    if not DEBUG:
        return

    print(f"\n[DEBUG] Placeholders for {slide_type} slide:")
    if not hasattr(slide, 'shapes') or not hasattr(slide.shapes, 'placeholders'):
        print("  No placeholders found")
        return

    for placeholder in slide.shapes.placeholders:
        ph_type = "UNKNOWN"
        try:
            if hasattr(placeholder, 'placeholder_format'):
                ph_type = placeholder.placeholder_format.type
        except:
            pass

        shape_type = "UNKNOWN"
        try:
            shape_type = placeholder.shape_type
        except:
            pass

        idx = "UNKNOWN"
        try:
            idx = placeholder.placeholder_format.idx
        except:
            pass

        print(f"  - idx={idx}, type={ph_type}, shape_type={shape_type}")


def get_title_placeholder(slide):
    """
    Find the title placeholder dynamically by type.

    Returns the title placeholder or None if not found.
    """
    for shape in slide.shapes:
        if not hasattr(shape, 'is_placeholder'):
            continue
        if shape.is_placeholder:
            try:
                if shape.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                    return shape
            except:
                pass
    return None


def get_body_placeholder(slide):
    """
    Find the body/content placeholder dynamically by type.

    Returns the body placeholder or None if not found.
    """
    for shape in slide.shapes:
        if not hasattr(shape, 'is_placeholder'):
            continue
        if shape.is_placeholder:
            try:
                ph_type = shape.placeholder_format.type
                # Body can be BODY, OBJECT, or generic CONTENT
                if ph_type in [PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT]:
                    return shape
            except:
                pass
    return None


def get_chart_placeholder(slide):
    """
    Find the chart placeholder dynamically by type.

    Returns the chart placeholder or None if not found.
    """
    # First try to find explicit CHART type
    for shape in slide.shapes:
        if not hasattr(shape, 'is_placeholder'):
            continue
        if shape.is_placeholder:
            try:
                if shape.placeholder_format.type == PP_PLACEHOLDER.CHART:
                    return shape
            except:
                pass

    # Fallback: look for OBJECT type placeholder
    for shape in slide.shapes:
        if not hasattr(shape, 'is_placeholder'):
            continue
        if shape.is_placeholder:
            try:
                if shape.placeholder_format.type == PP_PLACEHOLDER.OBJECT:
                    return shape
            except:
                pass

    return None


def list_all_placeholders(slide) -> str:
    """Return a formatted string listing all placeholders for error messages."""
    if not hasattr(slide, 'shapes') or not hasattr(slide.shapes, 'placeholders'):
        return "No placeholders found on slide"

    result = []
    for placeholder in slide.shapes.placeholders:
        try:
            idx = placeholder.placeholder_format.idx
            ph_type = placeholder.placeholder_format.type
            shape_type = placeholder.shape_type
            result.append(f"  idx={idx}, type={ph_type}, shape_type={shape_type}")
        except Exception as e:
            result.append(f"  Error reading placeholder: {e}")

    return "\n".join(result) if result else "No placeholders found"


# ============================================================================
# Slide Rendering Functions
# ============================================================================

def load_json_model(json_path: str) -> dict:
    """Load and parse the JSON slide model."""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def create_title_slide(slide, slide_data: dict):
    """
    Render a title slide.

    Expected layout:
    - Title placeholder
    - Subtitle placeholder (body/content type)
    """
    debug_print_placeholders(slide, "title")

    title = slide_data.get("title", "")
    subtitle = slide_data.get("content", {}).get("subtitle", "")

    # Set title
    title_placeholder = get_title_placeholder(slide)
    if title_placeholder:
        title_placeholder.text = title
    else:
        raise ValueError(
            f"Title slide has no title placeholder.\n"
            f"Available placeholders:\n{list_all_placeholders(slide)}"
        )

    # Set subtitle (use body placeholder for subtitle)
    subtitle_placeholder = get_body_placeholder(slide)
    if subtitle_placeholder:
        subtitle_placeholder.text = subtitle
    # Subtitle is optional, so no error if not found

    # Add speaker notes if present
    speaker_note = slide_data.get("speaker_note", "")
    if speaker_note:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = speaker_note


def create_key_insights_slide(slide, slide_data: dict):
    """
    Render a key insights slide.

    Expected layout:
    - Title placeholder
    - Body/content placeholder (for bullet points)
    """
    debug_print_placeholders(slide, "key_insights")

    title = slide_data.get("title", "")
    insights = slide_data.get("content", {}).get("insights", [])

    # Set title
    title_placeholder = get_title_placeholder(slide)
    if title_placeholder:
        title_placeholder.text = title
    else:
        raise ValueError(
            f"Key insights slide has no title placeholder.\n"
            f"Available placeholders:\n{list_all_placeholders(slide)}"
        )

    # Set bullet points
    body_placeholder = get_body_placeholder(slide)
    if not body_placeholder:
        raise ValueError(
            f"Key insights slide has no body/content placeholder for bullets.\n"
            f"Available placeholders:\n{list_all_placeholders(slide)}"
        )

    text_frame = body_placeholder.text_frame
    text_frame.clear()

    for idx, insight in enumerate(insights):
        if idx == 0:
            # First paragraph already exists
            p = text_frame.paragraphs[0]
        else:
            # Add new paragraphs for subsequent items
            p = text_frame.add_paragraph()

        p.text = insight
        p.level = 0

    # Add speaker notes if present
    speaker_note = slide_data.get("speaker_note", "")
    if speaker_note:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = speaker_note


def create_chart_slide(slide, slide_data: dict):
    """
    Render a chart slide.

    Expected layout:
    - Title placeholder
    - Chart placeholder

    Expected JSON content:
    {
      "chart_type": "column" | "bar" | "line" | "pie",
      "data": {
        "categories": [...],
        "series": [{"name": "...", "values": [...]}]
      }
    }
    """
    debug_print_placeholders(slide, "chart")

    title = slide_data.get("title", "")
    content = slide_data.get("content", {})
    chart_type = content.get("chart_type", "column")
    data = content.get("data", {})

    # Set title
    title_placeholder = get_title_placeholder(slide)
    if title_placeholder:
        title_placeholder.text = title
    else:
        raise ValueError(
            f"Chart slide has no title placeholder.\n"
            f"Available placeholders:\n{list_all_placeholders(slide)}"
        )

    # Validate chart data structure
    categories = data.get("categories", [])
    series_list = data.get("series", [])

    if not categories:
        raise ValueError(
            f"Chart slide {slide_data.get('slide_number')}: 'categories' is empty or missing in JSON data"
        )

    if not series_list:
        raise ValueError(
            f"Chart slide {slide_data.get('slide_number')}: 'series' is empty or missing in JSON data"
        )

    # Validate that all series have matching value counts
    expected_count = len(categories)
    for series in series_list:
        series_values = series.get("values", [])
        if len(series_values) != expected_count:
            raise ValueError(
                f"Chart slide {slide_data.get('slide_number')}: Series '{series.get('name')}' has "
                f"{len(series_values)} values but should have {expected_count} (matching categories)"
            )

    # Prepare chart data
    chart_data = ChartData()
    chart_data.categories = categories

    for series in series_list:
        series_name = series.get("name", "Series")
        series_values = series.get("values", [])
        chart_data.add_series(series_name, series_values)

    # Map chart type string to PowerPoint enum
    chart_type_map = {
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE,
        "pie": XL_CHART_TYPE.PIE,
    }

    if chart_type not in chart_type_map:
        raise ValueError(
            f"Chart slide {slide_data.get('slide_number')}: Unsupported chart type '{chart_type}'. "
            f"Supported types: {list(chart_type_map.keys())}"
        )

    xl_chart_type = chart_type_map[chart_type]

    # Find chart placeholder dynamically
    chart_placeholder = get_chart_placeholder(slide)

    if not chart_placeholder:
        raise ValueError(
            f"Chart slide has no chart placeholder (type CHART or OBJECT).\n"
            f"Available placeholders:\n{list_all_placeholders(slide)}\n"
            f"Please verify that layout index 2 in template.pptx has a Chart placeholder."
        )

    # Try to insert chart using placeholder
    chart = None
    try:
        if DEBUG:
            print(f"[DEBUG] Attempting insert_chart on placeholder...")
        chart = chart_placeholder.insert_chart(xl_chart_type, chart_data)
        if DEBUG:
            print(f"[DEBUG] ✓ Chart inserted successfully via placeholder")
    except AttributeError as e:
        # Placeholder doesn't support insert_chart, fallback to add_chart
        if DEBUG:
            print(f"[DEBUG] insert_chart not supported, falling back to add_chart: {e}")
        try:
            # Get placeholder position and size
            x = chart_placeholder.left
            y = chart_placeholder.top
            cx = chart_placeholder.width
            cy = chart_placeholder.height

            # Add chart at placeholder position
            chart = slide.shapes.add_chart(
                xl_chart_type, x, y, cx, cy, chart_data
            )
            if DEBUG:
                print(f"[DEBUG] ✓ Chart added successfully via add_chart fallback")
        except Exception as fallback_error:
            raise ValueError(
                f"Failed to insert chart using both methods:\n"
                f"  1. insert_chart: {str(e)}\n"
                f"  2. add_chart: {str(fallback_error)}\n"
                f"Available placeholders:\n{list_all_placeholders(slide)}"
            )
    except Exception as e:
        raise ValueError(
            f"Failed to insert chart into placeholder: {str(e)}\n"
            f"Available placeholders:\n{list_all_placeholders(slide)}"
        )

    if not chart:
        raise ValueError(
            f"Chart creation failed for unknown reason.\n"
            f"Available placeholders:\n{list_all_placeholders(slide)}"
        )

    # Add speaker notes if present
    speaker_note = slide_data.get("speaker_note", "")
    if speaker_note:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = speaker_note


def create_two_column_slide(slide, slide_data: dict):
    """Render a two-column slide (not yet implemented)."""
    raise NotImplementedError("Two-column slide type not yet implemented")


def create_recommendation_slide(slide, slide_data: dict):
    """Render a recommendation slide (not yet implemented)."""
    raise NotImplementedError("Recommendation slide type not yet implemented")


def create_appendix_slide(slide, slide_data: dict):
    """Render an appendix slide (not yet implemented)."""
    raise NotImplementedError("Appendix slide type not yet implemented")


# Map slide types to handler functions
SLIDE_HANDLERS = {
    "title": create_title_slide,
    "key_insights": create_key_insights_slide,
    "chart": create_chart_slide,
    "two_column": create_two_column_slide,
    "recommendation": create_recommendation_slide,
    "appendix": create_appendix_slide,
}


def validate_template(prs: Presentation):
    """
    Validate that the PowerPoint template has the required layouts.

    Raises detailed error messages if the template is misconfigured.
    """
    num_layouts = len(prs.slide_layouts)
    required_layouts = max(LAYOUT_INDICES.values()) + 1  # Indices are 0-based

    if num_layouts < required_layouts:
        raise ValueError(
            f"Template validation failed: Expected at least {required_layouts} layouts, "
            f"but template has only {num_layouts}.\n"
            f"Required layout indices: {LAYOUT_INDICES}\n"
            f"Please ensure template.pptx has layouts for all slide types."
        )

    print(f"✓ Template validated: {num_layouts} layouts found")

    # Warn about chart layout (most common misconfiguration)
    chart_layout_index = LAYOUT_INDICES.get("chart")
    if chart_layout_index is not None:
        chart_layout = prs.slide_layouts[chart_layout_index]
        print(f"  - Chart layout (index {chart_layout_index}): '{chart_layout.name}'")
        print(f"    → Ensure this layout has a chart placeholder, not just a content placeholder")


def generate_presentation(template_path: str, json_path: str, output_path: str):
    """
    Main function: Generate PowerPoint from JSON slide model.

    Args:
        template_path: Path to template.pptx
        json_path: Path to claude_output.json
        output_path: Path to save output.pptx
    """
    # Load template
    prs = Presentation(template_path)

    # Validate template structure
    validate_template(prs)

    # Load JSON model
    data = load_json_model(json_path)
    presentation_data = data.get("presentation", {})
    slides_data = presentation_data.get("slides", [])

    # Generate each slide
    for slide_data in slides_data:
        slide_type = slide_data.get("type")

        if slide_type not in LAYOUT_INDICES:
            print(f"Warning: Unknown slide type '{slide_type}', skipping slide {slide_data.get('slide_number')}")
            continue

        if slide_type not in SLIDE_HANDLERS:
            print(f"Warning: No handler for slide type '{slide_type}', skipping slide {slide_data.get('slide_number')}")
            continue

        # Create slide with appropriate layout
        layout_index = LAYOUT_INDICES[slide_type]
        slide_layout = prs.slide_layouts[layout_index]
        slide = prs.slides.add_slide(slide_layout)

        # Render slide content using appropriate handler
        try:
            handler = SLIDE_HANDLERS[slide_type]
            handler(slide, slide_data)
            print(f"✓ Generated slide {slide_data.get('slide_number')}: {slide_type}")
        except NotImplementedError as e:
            print(f"✗ Skipped slide {slide_data.get('slide_number')}: {e}")
        except Exception as e:
            print(f"✗ Error rendering slide {slide_data.get('slide_number')}: {e}")

    # Save presentation
    prs.save(output_path)
    print(f"\n✓ Presentation saved to: {output_path}")


if __name__ == "__main__":
    print("=== RUNNING LOCAL GENERATE_PPT.PY ===\n")

    # Default file paths
    TEMPLATE_PATH = "template.pptx"
    JSON_PATH = "claude_output.json"
    OUTPUT_PATH = "output.pptx"

    # Validate inputs
    if not Path(TEMPLATE_PATH).exists():
        print(f"Error: Template file not found: {TEMPLATE_PATH}")
        exit(1)

    if not Path(JSON_PATH).exists():
        print(f"Error: JSON file not found: {JSON_PATH}")
        exit(1)

    # Generate presentation
    print("Starting PowerPoint generation...\n")
    generate_presentation(TEMPLATE_PATH, JSON_PATH, OUTPUT_PATH)
