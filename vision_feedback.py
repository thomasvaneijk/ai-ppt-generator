#!/usr/bin/env python3
"""
AI Vision-Based Slide Feedback System

Uses vision-capable AI models to analyze rendered slide images and provide
concrete, actionable feedback on visual quality, hierarchy, and design.

This replaces heuristic-based critique with actual visual intelligence.
"""

import json
import base64
import os
from pathlib import Path
from typing import List, Dict, Optional
import anthropic


class VisionCritic:
    """
    AI-powered visual critic using Claude with vision capabilities.

    Analyzes slide images for:
    - Visual hierarchy and balance
    - Text density and readability
    - Professional polish vs template-likeness
    - Alignment and spacing
    - Chart clarity and storytelling
    """

    # Quality thresholds
    EXCELLENT_THRESHOLD = 8.5
    GOOD_THRESHOLD = 7.0
    POOR_THRESHOLD = 5.0

    CRITIQUE_PROMPT = """You are a senior presentation design consultant reviewing a PowerPoint slide.

Analyze this slide image and provide:

1. **VISUAL SCORE (0-10)**: Rate overall visual quality
   - 9-10: Exceptional - consultant-grade, high visual impact
   - 7-8: Good - professional, minor improvements possible
   - 5-6: Acceptable - functional but template-like
   - 3-4: Poor - needs significant rework
   - 0-2: Fails - empty, broken, or unusable

2. **DESIGN ISSUES**: List specific visual problems:
   - Text density (too many bullets, walls of text)
   - Visual hierarchy (title/body balance, emphasis)
   - Alignment and spacing problems
   - Template-likeness (boring, generic layout)
   - Chart clarity (if applicable: labels, legends, data story)
   - Empty or underutilized space
   - Color/contrast issues

3. **ACTIONABLE FIXES**: Concrete improvements (not vague suggestions):
   - "Reduce bullets from 8 to 5, keep only key insights"
   - "Enlarge chart to 60% of slide height"
   - "Add visual emphasis box around recommendation"
   - "Split into 2 slides: current is too dense"
   - "Remove legend, add data labels directly to bars"
   - "Shorten title from 12 words to 4-5 words"

4. **SLIDE ROLE**: What is this slide trying to do?
   - Title slide (introduction)
   - Data/chart slide (visual story)
   - Insight slide (key takeaway)
   - Recommendation slide (call to action)
   - Appendix/detail slide (reference)
   - UNCLEAR (fails to communicate purpose)

Return ONLY valid JSON (no markdown, no explanation):
{
  "score": <0-10>,
  "role": "<title|chart|insight|recommendation|appendix|unclear>",
  "issues": ["issue 1", "issue 2", ...],
  "actions": ["action 1", "action 2", ...],
  "summary": "One-sentence assessment"
}"""

    def __init__(self, api_key: Optional[str] = None):
        """
        Initialize vision critic.

        Args:
            api_key: Anthropic API key (defaults to ANTHROPIC_API_KEY env var)
        """
        self.api_key = api_key or os.getenv("ANTHROPIC_API_KEY")
        if not self.api_key:
            raise ValueError(
                "Anthropic API key required. Set ANTHROPIC_API_KEY environment variable "
                "or pass api_key parameter."
            )

        self.client = anthropic.Anthropic(api_key=self.api_key)

    def analyze_slide(self, image_path: str, slide_number: int) -> Dict:
        """
        Analyze a single slide image using AI vision.

        Args:
            image_path: Path to slide PNG image
            slide_number: Slide number (for tracking)

        Returns:
            Critique dictionary with score, issues, actions
        """
        # Read and encode image
        with open(image_path, "rb") as f:
            image_data = base64.standard_b64encode(f.read()).decode("utf-8")

        # Detect image format
        image_ext = Path(image_path).suffix.lower()
        media_type = {
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".webp": "image/webp",
            ".gif": "image/gif"
        }.get(image_ext, "image/png")

        # Call Claude with vision
        try:
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=1024,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": media_type,
                                    "data": image_data,
                                },
                            },
                            {
                                "type": "text",
                                "text": self.CRITIQUE_PROMPT
                            }
                        ],
                    }
                ],
            )

            # Parse response
            response_text = response.content[0].text.strip()

            # Remove markdown code blocks if present
            if response_text.startswith("```"):
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]
                response_text = response_text.strip()

            critique = json.loads(response_text)

            # Add metadata
            critique["slide"] = slide_number
            critique["image_path"] = image_path

            return critique

        except json.JSONDecodeError as e:
            # Fallback if JSON parsing fails
            return {
                "slide": slide_number,
                "score": 5.0,
                "role": "unclear",
                "issues": [f"Vision analysis failed to parse: {str(e)}"],
                "actions": ["Retry analysis or use heuristic fallback"],
                "summary": "Analysis error",
                "image_path": image_path
            }
        except Exception as e:
            # Fallback for other errors
            return {
                "slide": slide_number,
                "score": 5.0,
                "role": "unclear",
                "issues": [f"Vision analysis error: {str(e)}"],
                "actions": ["Check API key and network connection"],
                "summary": "Analysis error",
                "image_path": image_path
            }

    def analyze_directory(self, renders_dir: str = "renders") -> List[Dict]:
        """
        Analyze all slides in a directory.

        Args:
            renders_dir: Directory containing slide PNG files

        Returns:
            List of critique dictionaries
        """
        renders_path = Path(renders_dir)

        if not renders_path.exists():
            raise FileNotFoundError(f"Renders directory not found: {renders_dir}")

        # Find all slide PNG files
        slide_files = sorted(renders_path.glob("slide_*.png"))

        if not slide_files:
            raise FileNotFoundError(f"No slide PNG files found in {renders_dir}")

        print(f"\n{'='*70}")
        print(f"  AI VISION CRITIQUE - Analyzing {len(slide_files)} slides")
        print(f"{'='*70}\n")

        critiques = []

        for slide_file in slide_files:
            # Extract slide number from filename (e.g., slide_01.png → 1)
            slide_number = int(slide_file.stem.split("_")[1])

            print(f"Analyzing slide {slide_number}...", end=" ", flush=True)

            critique = self.analyze_slide(str(slide_file), slide_number)
            critiques.append(critique)

            # Print result with symbol
            score = critique.get("score", 0)
            if score >= self.EXCELLENT_THRESHOLD:
                symbol = "✓"
                color = "excellent"
            elif score >= self.GOOD_THRESHOLD:
                symbol = "○"
                color = "good"
            else:
                symbol = "✗"
                color = "poor"

            print(f"{symbol} {score:.1f}/10 - {critique.get('summary', 'No summary')}")

        return critiques

    def save_critique(self, critiques: List[Dict], output_path: str = "critique.json") -> Dict:
        """
        Save critique results to JSON file.

        Args:
            critiques: List of critique dictionaries
            output_path: Output JSON file path

        Returns:
            Summary dictionary
        """
        avg_score = sum(c.get("score", 0) for c in critiques) / len(critiques) if critiques else 0

        # Count issues by category
        all_issues = [issue for c in critiques for issue in c.get("issues", [])]
        all_actions = [action for c in critiques for action in c.get("actions", [])]

        critique_data = {
            "average_score": round(avg_score, 2),
            "total_slides": len(critiques),
            "quality_threshold": self.EXCELLENT_THRESHOLD,
            "meets_threshold": avg_score >= self.EXCELLENT_THRESHOLD,
            "slides": critiques,
            "summary": {
                "total_issues": len(all_issues),
                "total_actions": len(all_actions),
                "slides_excellent": sum(1 for c in critiques if c.get("score", 0) >= self.EXCELLENT_THRESHOLD),
                "slides_good": sum(1 for c in critiques if self.GOOD_THRESHOLD <= c.get("score", 0) < self.EXCELLENT_THRESHOLD),
                "slides_poor": sum(1 for c in critiques if c.get("score", 0) < self.GOOD_THRESHOLD)
            }
        }

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(critique_data, f, indent=2, ensure_ascii=False)

        # Print summary
        print(f"\n{'='*70}")
        print(f"  CRITIQUE SUMMARY")
        print(f"{'='*70}")
        print(f"  Average Score: {avg_score:.1f}/10")
        print(f"  Target: {self.EXCELLENT_THRESHOLD}/10")
        print(f"  Status: {'✓ MEETS THRESHOLD' if critique_data['meets_threshold'] else '✗ NEEDS IMPROVEMENT'}")
        print(f"\n  Quality Distribution:")
        print(f"    Excellent (≥{self.EXCELLENT_THRESHOLD}): {critique_data['summary']['slides_excellent']} slides")
        print(f"    Good ({self.GOOD_THRESHOLD}-{self.EXCELLENT_THRESHOLD}): {critique_data['summary']['slides_good']} slides")
        print(f"    Poor (<{self.GOOD_THRESHOLD}): {critique_data['summary']['slides_poor']} slides")
        print(f"\n✓ Critique saved to: {output_path}")
        print(f"{'='*70}\n")

        return critique_data


def main():
    """CLI entry point for vision-based slide critique."""
    import argparse

    parser = argparse.ArgumentParser(
        description="AI Vision-Based Slide Critique - Analyze rendered slides with Claude"
    )
    parser.add_argument("--input", default="renders", help="Directory containing slide PNG files")
    parser.add_argument("--output", default="critique.json", help="Output JSON file for critique")
    parser.add_argument("--api-key", help="Anthropic API key (or set ANTHROPIC_API_KEY env var)")

    args = parser.parse_args()

    # Check for API key
    api_key = args.api_key or os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("✗ ERROR: Anthropic API key required")
        print("\nOptions:")
        print("  1. Set environment variable: export ANTHROPIC_API_KEY=your-key")
        print("  2. Pass as argument: --api-key your-key")
        print("\nGet your API key from: https://console.anthropic.com/")
        return 1

    try:
        critic = VisionCritic(api_key=api_key)
        critiques = critic.analyze_directory(args.input)
        result = critic.save_critique(critiques, args.output)

        # Print detailed issues for poor-performing slides
        print("Detailed Feedback for Slides Needing Improvement:\n")
        for critique in critiques:
            if critique.get("score", 0) < VisionCritic.EXCELLENT_THRESHOLD:
                print(f"  Slide {critique['slide']}: {critique.get('score', 0):.1f}/10")
                print(f"    Role: {critique.get('role', 'unknown')}")
                if critique.get("issues"):
                    print(f"    Issues:")
                    for issue in critique["issues"]:
                        print(f"      - {issue}")
                if critique.get("actions"):
                    print(f"    Actions:")
                    for action in critique["actions"]:
                        print(f"      → {action}")
                print()

        return 0

    except Exception as e:
        print(f"✗ ERROR: {e}")
        return 1


if __name__ == "__main__":
    exit(main())
