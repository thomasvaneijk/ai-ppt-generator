#!/usr/bin/env python3
"""
Improvement Engine - Applies AI vision feedback to presentation JSON

Reads critique.json and automatically improves presentation structure
by modifying the JSON based on concrete, actionable feedback.
"""

import json
import re
from pathlib import Path
from typing import Dict, List, Any


class ImprovementEngine:
    """
    Applies visual critique feedback to presentation JSON.

    Transformations:
    - Reduce bullet counts based on density feedback
    - Shorten titles based on hierarchy feedback
    - Split dense slides into multiple slides
    - Merge weak slides
    - Adjust chart configurations for clarity
    - Add visual emphasis markers
    """

    def __init__(self, verbose: bool = True):
        self.verbose = verbose
        self.decisions_log = []

    def log(self, message: str, decision_type: str = "INFO"):
        """Log improvement decisions."""
        self.decisions_log.append({"type": decision_type, "message": message})
        if self.verbose:
            symbol = {
                "REMOVE": "✗",
                "SPLIT": "✂",
                "MERGE": "⊕",
                "TRIM": "✎",
                "ENHANCE": "✨",
                "INFO": "ℹ"
            }.get(decision_type, "·")
            print(f"  {symbol} {message}")

    def apply_improvements(
        self,
        presentation_path: str,
        critique_path: str,
        output_path: str
    ) -> Dict:
        """
        Apply improvements based on critique feedback.

        Args:
            presentation_path: Input JSON (creative_output.json)
            critique_path: Critique JSON from vision_feedback.py
            output_path: Output JSON with improvements applied

        Returns:
            Improved presentation data
        """
        print("=" * 70)
        print("  IMPROVEMENT ENGINE - Applying Vision Feedback")
        print("=" * 70)

        # Load presentation
        with open(presentation_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Load critique
        with open(critique_path, 'r', encoding='utf-8') as f:
            critique_data = json.load(f)

        presentation = data.get("presentation", {})
        original_slides = presentation.get("slides", [])
        slide_critiques = critique_data.get("slides", [])

        self.log(f"Original: {len(original_slides)} slides", "INFO")
        self.log(f"Critique: avg score {critique_data.get('average_score', 0):.1f}/10", "INFO")

        # Create critique lookup by slide number
        critique_map = {c["slide"]: c for c in slide_critiques}

        improved_slides = []

        for slide in original_slides:
            slide_num = slide.get("slide_number")
            critique = critique_map.get(slide_num, {})

            # Apply improvements based on critique
            improved_slide = self._improve_slide(slide, critique)

            # Check if slide should be split
            if self._should_split_slide(slide, critique):
                split_slides = self._split_slide(improved_slide, critique)
                improved_slides.extend(split_slides)
                self.log(f"Slide {slide_num}: Split into {len(split_slides)} slides", "SPLIT")
            elif improved_slide is not None:
                improved_slides.append(improved_slide)

        # Renumber slides
        for i, slide in enumerate(improved_slides, start=1):
            slide["slide_number"] = i

        # Update presentation
        presentation["slides"] = improved_slides
        data["presentation"] = presentation

        # Save improved presentation
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        self.log(f"Final: {len(improved_slides)} slides", "INFO")
        print("=" * 70)
        print(f"\n✓ Improvements applied: {output_path}")
        print(f"  {len(original_slides)} slides → {len(improved_slides)} slides")

        return data

    def _improve_slide(self, slide: Dict, critique: Dict) -> Dict:
        """
        Apply improvements to a single slide based on critique.

        Returns:
            Improved slide or None if slide should be removed
        """
        improved = slide.copy()
        slide_num = slide.get("slide_number")
        actions = critique.get("actions", [])
        score = critique.get("score", 10)

        # If slide scores too low, remove it
        if score < 3.0:
            self.log(f"Slide {slide_num}: Score {score:.1f} too low, removing", "REMOVE")
            return None

        # Apply action-based improvements
        for action in actions:
            improved = self._apply_action(improved, action, slide_num)

        return improved

    def _apply_action(self, slide: Dict, action: str, slide_num: int) -> Dict:
        """Apply a specific action to a slide."""
        action_lower = action.lower()

        # Reduce bullets
        if "reduce bullets" in action_lower or "trim bullets" in action_lower:
            match = re.search(r"(\d+)\s+(?:to|→)\s+(\d+)", action)
            if match:
                target = int(match.group(2))
                slide = self._trim_bullets(slide, target, slide_num)

        # Shorten title
        if "shorten title" in action_lower or "reduce title" in action_lower:
            match = re.search(r"(\d+)\s+words", action)
            if match:
                max_words = int(match.group(1))
                slide = self._shorten_title(slide, max_words, slide_num)

        # Enlarge chart
        if "enlarge chart" in action_lower or "increase chart" in action_lower:
            slide = self._add_chart_emphasis(slide, slide_num)

        # Add visual emphasis
        if "add visual emphasis" in action_lower or "add emphasis box" in action_lower:
            slide = self._add_visual_emphasis(slide, slide_num)

        # Remove legend
        if "remove legend" in action_lower:
            slide = self._remove_chart_legend(slide, slide_num)

        # Add data labels
        if "add data labels" in action_lower or "add labels" in action_lower:
            slide = self._add_data_labels(slide, slide_num)

        return slide

    def _trim_bullets(self, slide: Dict, target: int, slide_num: int) -> Dict:
        """Trim bullets to target count."""
        content = slide.get("content", {})
        bullets = content.get("bullets", [])

        if len(bullets) > target:
            original_count = len(bullets)
            content["bullets"] = bullets[:target]
            slide["content"] = content
            self.log(f"Slide {slide_num}: Trimmed bullets {original_count} → {target}", "TRIM")

        return slide

    def _shorten_title(self, slide: Dict, max_words: int, slide_num: int) -> Dict:
        """Shorten title to max words."""
        title = slide.get("title", "")
        words = title.split()

        if len(words) > max_words:
            original = title
            new_title = " ".join(words[:max_words])
            slide["title"] = new_title
            self.log(f"Slide {slide_num}: Shortened title '{original}' → '{new_title}'", "TRIM")

        return slide

    def _add_chart_emphasis(self, slide: Dict, slide_num: int) -> Dict:
        """Mark chart for visual emphasis (renderer will enlarge)."""
        if slide.get("type") == "chart":
            content = slide.get("content", {})
            content["visual_emphasis"] = "large_chart"
            slide["content"] = content
            self.log(f"Slide {slide_num}: Marked chart for enlargement", "ENHANCE")

        return slide

    def _add_visual_emphasis(self, slide: Dict, slide_num: int) -> Dict:
        """Add visual emphasis marker."""
        visuals = slide.get("visuals", {})
        visuals["emphasis"] = "add_background_shape"
        slide["visuals"] = visuals
        self.log(f"Slide {slide_num}: Added visual emphasis marker", "ENHANCE")

        return slide

    def _remove_chart_legend(self, slide: Dict, slide_num: int) -> Dict:
        """Mark chart to remove legend."""
        if slide.get("type") == "chart":
            content = slide.get("content", {})
            data = content.get("data", {})
            data["show_legend"] = False
            content["data"] = data
            slide["content"] = content
            self.log(f"Slide {slide_num}: Marked chart to hide legend", "ENHANCE")

        return slide

    def _add_data_labels(self, slide: Dict, slide_num: int) -> Dict:
        """Mark chart to show data labels."""
        if slide.get("type") == "chart":
            content = slide.get("content", {})
            data = content.get("data", {})
            data["show_data_labels"] = True
            content["data"] = data
            slide["content"] = content
            self.log(f"Slide {slide_num}: Marked chart to show data labels", "ENHANCE")

        return slide

    def _should_split_slide(self, slide: Dict, critique: Dict) -> bool:
        """Check if slide should be split based on critique."""
        actions = critique.get("actions", [])

        for action in actions:
            if "split" in action.lower() and ("2 slides" in action.lower() or "two slides" in action.lower()):
                return True

        return False

    def _split_slide(self, slide: Dict, critique: Dict) -> List[Dict]:
        """Split a dense slide into multiple slides."""
        slide_type = slide.get("type")

        if slide_type == "key_insights":
            return self._split_key_insights_slide(slide)

        # Default: return slide as-is
        return [slide]

    def _split_key_insights_slide(self, slide: Dict) -> List[Dict]:
        """Split key insights slide into multiple slides."""
        content = slide.get("content", {})
        bullets = content.get("bullets", [])

        if len(bullets) <= 5:
            return [slide]

        # Split into chunks of 4-5 bullets
        mid = len(bullets) // 2

        slide1 = slide.copy()
        slide1["content"] = {"bullets": bullets[:mid]}
        slide1["title"] = slide.get("title", "") + " (Part 1)"

        slide2 = slide.copy()
        slide2["content"] = {"bullets": bullets[mid:]}
        slide2["title"] = slide.get("title", "") + " (Part 2)"

        return [slide1, slide2]


def main():
    """CLI entry point for improvement engine."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Improvement Engine - Apply vision feedback to presentation JSON"
    )
    parser.add_argument("--input", default="creative_output.json", help="Input presentation JSON")
    parser.add_argument("--critique", default="critique.json", help="Critique JSON from vision feedback")
    parser.add_argument("--output", default="improved_output.json", help="Output JSON with improvements")
    parser.add_argument("--quiet", action="store_true", help="Suppress verbose output")

    args = parser.parse_args()

    # Check files exist
    if not Path(args.input).exists():
        print(f"✗ ERROR: Input file not found: {args.input}")
        return 1

    if not Path(args.critique).exists():
        print(f"✗ ERROR: Critique file not found: {args.critique}")
        return 1

    # Run improvement engine
    engine = ImprovementEngine(verbose=not args.quiet)
    engine.apply_improvements(args.input, args.critique, args.output)

    print("\n" + "=" * 70)
    print("  NEXT STEP: Re-render presentation")
    print("=" * 70)
    print(f"  python generate_ppt.py")
    print(f"  (Will use {args.output} if it exists)")

    return 0


if __name__ == "__main__":
    exit(main())
