#!/usr/bin/env python3
"""
Iterative Presentation Improvement Engine

Orchestrates the full creative feedback loop:
1. Generate presentation (creative director + rendering)
2. Render slides to PNG (Windows only)
3. AI vision critique
4. Apply improvements automatically
5. Repeat until quality threshold reached
"""

import json
import subprocess
import sys
import os
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
        "target_score": 8.5
    }


def run_command(command: list, description: str, allow_failure: bool = False) -> bool:
    """Run a command and handle errors."""
    print(f"\n{'='*70}")
    print(f"  {description}")
    print(f"{'='*70}")

    result = subprocess.run(command, capture_output=False, text=True)

    if result.returncode != 0:
        if not allow_failure:
            print(f"\n✗ Error running: {' '.join(command)}")
            return False
        else:
            print(f"\n⚠ Warning: Command failed (non-critical)")
            return True

    return True


def load_critique():
    """Load critique results from critique.json."""
    critique_path = Path("critique.json")

    if not critique_path.exists():
        return None

    with open(critique_path, 'r') as f:
        return json.load(f)


def iterate_improvements():
    """
    Main iteration loop.

    Flow:
    1. Generate presentation (generate.py)
    2. Render slides to PNG (render_slides.py)
    3. AI vision critique (vision_feedback.py)
    4. Apply improvements (improvement_engine.py)
    5. Check if target score reached, repeat if not
    """
    config = load_config()
    max_iterations = config.get("max_iterations", 3)
    target_score = config.get("target_score", 8.5)

    print("="*70)
    print("  ITERATIVE CREATIVE IMPROVEMENT ENGINE")
    print("="*70)
    print(f"  Max iterations: {max_iterations}")
    print(f"  Target score: {target_score}/10")
    print("="*70)

    # Check prerequisites
    has_api_key = bool(os.getenv("ANTHROPIC_API_KEY"))
    if not has_api_key:
        print("\n⚠ WARNING: ANTHROPIC_API_KEY not set")
        print("  Vision feedback will not work without API key")
        print("  Set it with: export ANTHROPIC_API_KEY=your-key")
        print("\nContinuing without vision feedback (heuristic fallback)...\n")

    # Determine which JSON to use as starting point
    input_json = "creative_output.json" if Path("creative_output.json").exists() else "claude_output.json"

    for iteration in range(1, max_iterations + 1):
        print(f"\n\n{'#'*70}")
        print(f"  ITERATION {iteration}/{max_iterations}")
        print(f"{'#'*70}")

        # Step 1: Generate PowerPoint
        print(f"\n[STEP 1] Generating presentation from {input_json}")
        if not run_command(
            [sys.executable, "generate.py"],
            f"Iteration {iteration}: Full Pipeline (Creative Director + Rendering)"
        ):
            print("✗ Failed to generate PowerPoint")
            return False

        # Check if output.pptx exists
        if not Path("output.pptx").exists():
            print("✗ output.pptx not found after generation")
            return False

        # Step 2: Render slides (Windows + PowerPoint required)
        print(f"\n[STEP 2] Rendering slides to PNG")
        if sys.platform != "win32":
            print("⚠ Skipping render step (requires Windows + PowerPoint)")
            print("  To use full feedback loop, run on Windows")
            break

        if not run_command(
            [sys.executable, "render_slides.py", "--input", "output.pptx", "--output", "renders"],
            f"Iteration {iteration}: Rendering Slides",
            allow_failure=True
        ):
            print("⚠ Rendering failed - continuing without visual feedback")
            break

        # Check if renders exist
        if not Path("renders").exists() or not list(Path("renders").glob("slide_*.png")):
            print("⚠ No rendered slides found - skipping critique")
            break

        # Step 3: AI Vision Critique
        print(f"\n[STEP 3] AI vision critique")

        if has_api_key:
            # Use AI vision feedback
            if not run_command(
                [sys.executable, "vision_feedback.py", "--input", "renders", "--output", "critique.json"],
                f"Iteration {iteration}: AI Vision Critique",
                allow_failure=True
            ):
                print("⚠ Vision critique failed - using heuristic fallback")
                run_command(
                    [sys.executable, "critique_slides.py", "--input", "renders", "--output", "critique.json"],
                    "Fallback: Heuristic Critique",
                    allow_failure=True
                )
        else:
            # Fallback to heuristic critique
            if not run_command(
                [sys.executable, "critique_slides.py", "--input", "renders", "--output", "critique.json"],
                f"Iteration {iteration}: Heuristic Critique (no API key)",
                allow_failure=True
            ):
                print("⚠ Critique failed - cannot continue iteration")
                break

        # Load critique results
        critique_data = load_critique()
        if not critique_data:
            print("✗ No critique.json found - cannot apply improvements")
            break

        avg_score = critique_data.get("average_score", 0)
        print(f"\n{'='*70}")
        print(f"  SCORE: {avg_score:.1f}/10 (target: {target_score}/10)")
        print(f"{'='*70}")

        # Check if target reached
        if avg_score >= target_score:
            print(f"\n✓ TARGET REACHED: {avg_score:.1f} ≥ {target_score}")
            print(f"✓ Presentation quality meets threshold after {iteration} iteration(s)")
            return True

        # Check if this is last iteration
        if iteration == max_iterations:
            print(f"\n⚠ Max iterations reached ({max_iterations})")
            print(f"  Final score: {avg_score:.1f}/{target_score}")
            print(f"  Presentation may benefit from manual review")
            break

        # Step 4: Apply improvements
        print(f"\n[STEP 4] Applying improvements for next iteration")

        # Determine input/output for improvement engine
        improvement_input = "creative_output.json" if Path("creative_output.json").exists() else input_json
        improvement_output = "improved_output.json"

        if not run_command(
            [sys.executable, "improvement_engine.py",
             "--input", improvement_input,
             "--critique", "critique.json",
             "--output", improvement_output],
            f"Iteration {iteration}: Applying Improvements",
            allow_failure=True
        ):
            print("⚠ Improvement application failed - manual review recommended")
            break

        # Use improved output for next iteration
        if Path(improvement_output).exists():
            # Copy improved_output.json to creative_output.json for next render
            import shutil
            shutil.copy(improvement_output, "creative_output.json")
            input_json = "creative_output.json"
            print(f"✓ Improvements applied - will use {improvement_output} for next iteration")
        else:
            print("⚠ No improved output generated - stopping iteration")
            break

    # Final summary
    final_critique = load_critique()
    if final_critique:
        final_score = final_critique.get("average_score", 0)
        print(f"\n{'='*70}")
        print(f"  FINAL RESULT")
        print(f"{'='*70}")
        print(f"  Final Score: {final_score:.1f}/10")
        print(f"  Target: {target_score}/10")
        print(f"  Status: {'✓ MEETS THRESHOLD' if final_score >= target_score else '✗ BELOW THRESHOLD'}")
        print(f"{'='*70}\n")

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
