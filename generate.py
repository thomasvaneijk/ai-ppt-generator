#!/usr/bin/env python3
"""
Autonomous PowerPoint Generation Pipeline

Orchestrates the full creative workflow:
1. Creative Director (pre-render intelligence)
2. Slide Renderer (PowerPoint generation)

This is the main entry point for presentation generation.
"""

import sys
import subprocess
from pathlib import Path


def run_step(script: str, description: str) -> bool:
    """Run a pipeline step and report success/failure."""
    print(f"\n{'='*70}")
    print(f"  {description}")
    print(f"{'='*70}")

    try:
        result = subprocess.run(
            [sys.executable, script],
            check=True,
            capture_output=False,
            text=True
        )
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n✗ FAILED: {description}")
        print(f"  Error: {e}")
        return False


def main():
    """Execute the full presentation generation pipeline."""
    print("\n" + "="*70)
    print("  AUTONOMOUS PRESENTATION GENERATION PIPELINE")
    print("="*70)

    # Check prerequisites
    if not Path("template.pptx").exists():
        print("✗ ERROR: template.pptx not found")
        sys.exit(1)

    if not Path("claude_output.json").exists():
        print("✗ ERROR: claude_output.json not found")
        sys.exit(1)

    # Step 1: Creative Director (Pre-render Intelligence)
    if not run_step("creative_director.py", "STEP 1: Creative Director Review"):
        print("\n✗ Pipeline failed at Creative Director stage")
        sys.exit(1)

    # Step 2: Generate PowerPoint (Rendering)
    # Modify generate_ppt.py to use creative_output.json instead of claude_output.json
    if not run_step("generate_ppt.py", "STEP 2: PowerPoint Generation"):
        print("\n✗ Pipeline failed at PowerPoint Generation stage")
        sys.exit(1)

    # Success
    print("\n" + "="*70)
    print("  ✓ PIPELINE COMPLETE")
    print("="*70)
    print("\nOutput:")
    print("  - creative_output.json (creative review)")
    print("  - output.pptx (final presentation)")
    print()


if __name__ == "__main__":
    main()
