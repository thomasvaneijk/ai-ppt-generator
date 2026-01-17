"""
Iterative Presentation Improvement Engine

Orchestrates the generate → render → critique → improve loop.
"""

import json
import subprocess
import sys
from pathlib import Path


def load_config():
    """Load configuration from config.json."""
    config_path = Path("config.json")

    if config_path.exists():
        with open(config_path, 'r') as f:
            return json.load(f)

    # Default configuration
    return {
        "max_iterations": 3,
        "target_score": 8.0
    }


def run_command(command: list, description: str):
    """Run a command and handle errors."""
    print(f"\n{'='*60}")
    print(f"  {description}")
    print(f"{'='*60}")

    result = subprocess.run(command, capture_output=True, text=True)

    if result.returncode != 0:
        print(f"\n✗ Error running: {' '.join(command)}")
        print(result.stderr)
        return False

    print(result.stdout)
    return True


def load_critique():
    """Load critique results from critique.json."""
    critique_path = Path("critique.json")

    if not critique_path.exists():
        return None

    with open(critique_path, 'r') as f:
        return json.load(f)


def apply_improvements(critique_data: dict) -> bool:
    """
    Apply improvements based on critique.

    This is a placeholder for the improvement logic.
    In production, this would:
    - Parse critique actions
    - Modify claude_output.json accordingly
    - Adjust titles, bullet counts, layout choices

    Returns True if improvements were applied.
    """
    # TODO: Implement actual improvement logic
    # For now, this is a placeholder that would need:
    # - Action parsing (e.g., "shorten title" → reduce title word count)
    # - JSON modification (e.g., edit claude_output.json)
    # - Layout adjustments (e.g., change slide type if needed)

    print("\n[INFO] Improvement application not yet implemented.")
    print("[INFO] This is where we would:")
    print("  - Parse critique actions")
    print("  - Modify claude_output.json")
    print("  - Adjust titles, bullets, or layouts")

    return False  # No improvements applied yet


def iterate_improvements():
    """
    Main iteration loop.

    1. Generate PowerPoint (generate_ppt.py)
    2. Render slides to PNG (render_slides.py)
    3. Critique slides (critique_slides.py)
    4. Check if target score reached
    5. If not, apply improvements and repeat
    """
    config = load_config()
    max_iterations = config["max_iterations"]
    target_score = config["target_score"]

    print("="*60)
    print("  ITERATIVE PRESENTATION IMPROVEMENT ENGINE")
    print("="*60)
    print(f"  Max iterations: {max_iterations}")
    print(f"  Target score: {target_score}/10")
    print("="*60)

    for iteration in range(1, max_iterations + 1):
        print(f"\n\n{'#'*60}")
        print(f"  ITERATION {iteration}/{max_iterations}")
        print(f"{'#'*60}")

        # Step 1: Generate PowerPoint
        if not run_command(
            [sys.executable, "generate_ppt.py"],
            f"Step 1/{iteration}: Generating PowerPoint"
        ):
            print("✗ Failed to generate PowerPoint")
            return False

        # Check if output.pptx exists
        if not Path("output.pptx").exists():
            print("✗ output.pptx not found after generation")
            return False

        # Step 2: Render slides (skip on non-Windows or if COM unavailable)
        print("\n[INFO] Slide rendering requires Windows + PowerPoint.")
        print("[INFO] Run manually: python render_slides.py")
        print("[INFO] Skipping render step in this iteration.")

        # For now, skip rendering and critique
        # In production on Windows, these would run:
        # run_command([sys.executable, "render_slides.py"], ...)
        # run_command([sys.executable, "critique_slides.py"], ...)

        print("\n[INFO] Rendering and critique not executed (Windows-only).")
        print("[INFO] To complete the loop on Windows:")
        print("  1. Run: python render_slides.py")
        print("  2. Run: python critique_slides.py")
        print("  3. Check critique.json for feedback")

        break  # Exit after first iteration since we can't render

    print("\n✓ Iteration complete.")
    print("\nTo use the full feedback loop on Windows:")
    print("  1. Ensure PowerPoint is installed")
    print("  2. Install: pip install pywin32")
    print("  3. Run: python iterate.py")

    return True


def main():
    """Entry point for iteration engine."""
    try:
        success = iterate_improvements()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n✗ Interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
