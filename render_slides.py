"""
Slide Renderer - Converts PowerPoint slides to PNG images

Uses PowerPoint COM automation (win32com) to export slides as images.
Requires PowerPoint installed on Windows.
"""

import os
import sys
from pathlib import Path

try:
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False
    print("WARNING: win32com not available. Install with: pip install pywin32")


def render_slides_com(pptx_path: str, output_dir: str = "renders") -> list:
    """
    Render PowerPoint slides to PNG using COM automation.

    Args:
        pptx_path: Path to the PowerPoint file
        output_dir: Directory to save rendered PNG files

    Returns:
        List of rendered PNG file paths
    """
    if not COM_AVAILABLE:
        raise RuntimeError("win32com.client not available. Cannot render slides.")

    # Create output directory
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)

    # Get absolute paths
    pptx_abs = Path(pptx_path).absolute()
    output_abs = output_path.absolute()

    if not pptx_abs.exists():
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")

    print(f"Opening PowerPoint: {pptx_abs}")

    # Launch PowerPoint
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1  # Make visible for debugging

    try:
        # Open presentation
        presentation = powerpoint.Presentations.Open(str(pptx_abs), WithWindow=False)

        print(f"Rendering {presentation.Slides.Count} slides...")

        rendered_files = []

        # Export each slide
        for i, slide in enumerate(presentation.Slides, start=1):
            output_file = output_abs / f"slide_{i:02d}.png"

            # Export as PNG (resolution: 1920x1080)
            slide.Export(str(output_file), "PNG", 1920, 1080)

            rendered_files.append(str(output_file))
            print(f"  ✓ Rendered slide {i}: {output_file.name}")

        # Close presentation
        presentation.Close()

        print(f"\n✓ All slides rendered to: {output_abs}")
        return rendered_files

    finally:
        # Quit PowerPoint
        powerpoint.Quit()


def render_slides_fallback(pptx_path: str, output_dir: str = "renders") -> list:
    """
    Fallback renderer (placeholder for non-Windows environments).

    This is a placeholder. In production, you could use:
    - LibreOffice headless conversion
    - pptx2png libraries
    - Cloud-based rendering services
    """
    print("ERROR: COM automation not available and no fallback implemented.")
    print("This tool requires:")
    print("  1. Windows OS")
    print("  2. Microsoft PowerPoint installed")
    print("  3. pywin32: pip install pywin32")
    sys.exit(1)


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Render PowerPoint slides to PNG images")
    parser.add_argument("--input", default="output.pptx", help="Input PowerPoint file")
    parser.add_argument("--output", default="renders", help="Output directory for PNG files")

    args = parser.parse_args()

    if COM_AVAILABLE:
        rendered = render_slides_com(args.input, args.output)
        print(f"\n✓ Rendered {len(rendered)} slides")
    else:
        render_slides_fallback(args.input, args.output)


if __name__ == "__main__":
    main()
