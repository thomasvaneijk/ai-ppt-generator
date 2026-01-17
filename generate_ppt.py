#!/usr/bin/env python3
"""
PowerPoint Generator MVP

Generates professional PowerPoint presentations from JSON slide models.
Follows strict architecture: JSON → Python → PowerPoint.
"""

import json
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt


# Layout indices in template.pptx
LAYOUT_INDICES = {
    "title": 0,
    "key_insights": 1,
    "chart": 2,
    "two_column": 3,
    "recommendation": 4,
    "appendix": 5,
}


def load_json_model(json_path: str) -> dict:
    """Load and parse the JSON slide model."""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def create_title_slide(slide, slide_data: dict):
    """
    Render a title slide.

    Expected layout:
    - Placeholder 0: Title
    - Placeholder 1: Subtitle
    """
    title = slide_data.get("title", "")
    subtitle = slide_data.get("content", {}).get("subtitle", "")

    # Set title
    if len(slide.shapes.placeholders) > 0:
        title_placeholder = slide.shapes.placeholders[0]
        title_placeholder.text = title

    # Set subtitle
    if len(slide.shapes.placeholders) > 1:
        subtitle_placeholder = slide.shapes.placeholders[1]
        subtitle_placeholder.text = subtitle

    # Add speaker notes if present
    speaker_note = slide_data.get("speaker_note", "")
    if speaker_note:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = speaker_note


def create_key_insights_slide(slide, slide_data: dict):
    """
    Render a key insights slide.

    Expected layout:
    - Placeholder 0: Title
    - Placeholder 1: Content (bullet points)
    """
    title = slide_data.get("title", "")
    insights = slide_data.get("content", {}).get("insights", [])

    # Set title
    if len(slide.shapes.placeholders) > 0:
        title_placeholder = slide.shapes.placeholders[0]
        title_placeholder.text = title

    # Set bullet points
    if len(slide.shapes.placeholders) > 1:
        content_placeholder = slide.shapes.placeholders[1]
        text_frame = content_placeholder.text_frame
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
    """Render a chart slide (not yet implemented)."""
    raise NotImplementedError("Chart slide type not yet implemented")


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
