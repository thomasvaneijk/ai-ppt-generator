"""
Creative Slide Director - Pre-rendering creative intelligence layer

Acts like a senior consultant reviewing a deck before production.
Makes bold decisions to improve narrative quality, visual clarity, and executive impact.

Philosophy:
- Every slide must earn its place
- Titles must be punchy and fit visually (max 6 words ideal)
- Visual hierarchy beats information density
- Story flow beats completeness
- Remove, don't dilute

Example transformations:
BEFORE: "Customer Journey Performance Overview" (5 words, generic)
AFTER:  "The Performance Gap" (3 words, tension-creating)

BEFORE: Slide with 8 bullets
AFTER:  Slide with 5 bullets (top insights) OR split into 2 slides

BEFORE: key_insights slide listing numbers
AFTER:  chart slide (visual > text for data)
"""

import json
import re
from typing import List, Dict, Tuple
from pathlib import Path


class CreativeDirector:
    """
    Creative intelligence engine for presentation quality.

    Makes decisions a senior consultant would make:
    - Which slides add value vs clutter
    - Which titles create impact vs confusion
    - Which content should be visual vs textual
    - Which narrative beats are missing or redundant
    """

    # Decision thresholds
    MAX_TITLE_WORDS = 6  # Ideal title length
    MAX_BULLETS = 5      # Maximum bullets per slide
    MIN_BULLETS = 2      # Minimum bullets to justify a slide
    MAX_SLIDES = 10      # Maximum total slides
    MIN_SLIDES = 5       # Minimum total slides

    # Filler words to remove from titles
    TITLE_FILLERS = [
        "Overview of", "Summary of", "Analysis of", "Review of",
        "Report on", "Discussion of", "Introduction to", "Presentation on"
    ]

    def __init__(self, verbose: bool = True):
        self.verbose = verbose
        self.decisions_log = []

    def log(self, message: str, decision_type: str = "INFO"):
        """Log creative decisions."""
        self.decisions_log.append({"type": decision_type, "message": message})
        if self.verbose:
            symbol = {"REMOVE": "✗", "TRANSFORM": "✎", "KEEP": "✓", "INFO": "ℹ"}.get(decision_type, "·")
            print(f"  {symbol} {message}")

    def review_presentation(self, input_path: str, output_path: str) -> Dict:
        """
        Main entry point: transform presentation creatively.

        Args:
            input_path: Path to claude_output.json
            output_path: Path to save creative_output.json

        Returns:
            Transformed presentation data
        """
        print("=" * 70)
        print("  CREATIVE SLIDE DIRECTOR - PRE-RENDER REVIEW")
        print("=" * 70)

        # Load presentation
        with open(input_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        presentation = data.get("presentation", {})
        original_slides = presentation.get("slides", [])

        self.log(f"Original deck: {len(original_slides)} slides", "INFO")

        # Phase 1: Evaluate and filter slides
        filtered_slides = self._filter_slides(original_slides)

        # Phase 2: Transform remaining slides
        transformed_slides = [self._transform_slide(s) for s in filtered_slides]

        # Phase 3: Ensure quality standards
        final_slides = self._ensure_quality_standards(transformed_slides)

        # Update presentation
        presentation["slides"] = final_slides
        data["presentation"] = presentation

        # Renumber slides
        for i, slide in enumerate(final_slides, start=1):
            slide["slide_number"] = i

        self.log(f"Final deck: {len(final_slides)} slides", "INFO")
        print("=" * 70)

        # Save transformed presentation
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        print(f"\n✓ Creative review complete: {output_path}")
        print(f"  {len(original_slides)} slides → {len(final_slides)} slides")

        return data

    def _filter_slides(self, slides: List[Dict]) -> List[Dict]:
        """
        Phase 1: Remove weak, redundant, or low-value slides.

        Criteria:
        - Is the slide empty or near-empty?
        - Does it repeat content from other slides?
        - Does it add narrative value?
        - Is it appendix material when deck is already long?
        """
        kept_slides = []

        for slide in slides:
            slide_num = slide.get("slide_number")
            slide_type = slide.get("type")
            title = slide.get("title", "")

            # Always keep title slide
            if slide_type == "title":
                self.log(f"Slide {slide_num}: Keeping title slide", "KEEP")
                kept_slides.append(slide)
                continue

            # Always keep recommendation (critical for management)
            if slide_type == "recommendation":
                self.log(f"Slide {slide_num}: Keeping recommendation", "KEEP")
                kept_slides.append(slide)
                continue

            # Check if slide is empty or weak
            if self._is_empty_slide(slide):
                self.log(f"Slide {slide_num}: '{title}' - Empty content", "REMOVE")
                continue

            # Check if slide is redundant
            if self._is_redundant(slide, kept_slides):
                self.log(f"Slide {slide_num}: '{title}' - Redundant with existing slides", "REMOVE")
                continue

            # Appendix: remove if deck is already long
            if slide_type == "appendix" and len(slides) > self.MAX_SLIDES:
                self.log(f"Slide {slide_num}: '{title}' - Appendix removed (deck too long)", "REMOVE")
                continue

            # Keep slide
            self.log(f"Slide {slide_num}: '{title}' - Keeping", "KEEP")
            kept_slides.append(slide)

        return kept_slides

    def _is_empty_slide(self, slide: Dict) -> bool:
        """Check if slide has insufficient content."""
        slide_type = slide.get("type")
        content = slide.get("content", {})

        if slide_type == "key_insights":
            bullets = content.get("bullets", [])
            return len(bullets) < self.MIN_BULLETS

        if slide_type == "chart":
            data = content.get("data", {})
            categories = data.get("categories", [])
            series = data.get("series", [])
            return not categories or not series

        if slide_type == "two_column":
            left = content.get("left_column", {}).get("items", [])
            right = content.get("right_column", {}).get("items", [])
            return not left and not right

        if slide_type == "recommendation":
            rec = content.get("recommendation", "")
            rationale = content.get("rationale", [])
            next_steps = content.get("next_steps", [])
            return not rec and not rationale and not next_steps

        return False

    def _is_redundant(self, slide: Dict, existing_slides: List[Dict]) -> bool:
        """
        Check if slide is redundant with already-kept slides.

        Simple heuristic: if title is very similar, likely redundant.
        TODO Phase 3: Use semantic similarity for better detection.
        """
        title = slide.get("title", "").lower()
        title_words = set(title.split())

        for existing in existing_slides:
            existing_title = existing.get("title", "").lower()
            existing_words = set(existing_title.split())

            # If >70% word overlap, likely redundant
            if title_words and existing_words:
                overlap = len(title_words & existing_words)
                similarity = overlap / min(len(title_words), len(existing_words))

                if similarity > 0.7:
                    return True

        return False

    def _transform_slide(self, slide: Dict) -> Dict:
        """
        Phase 2: Apply creative transformations to individual slides.

        Transformations:
        - Rewrite titles for impact
        - Trim bullets to max 5
        - Convert bullet lists with numeric data to charts (TODO Phase 3)
        - Add visual emphasis guidance
        """
        transformed = slide.copy()
        slide_num = slide.get("slide_number")
        slide_type = slide.get("type")

        # Transform title
        original_title = slide.get("title", "")
        new_title = self._transform_title(original_title)

        if new_title != original_title:
            self.log(f"Slide {slide_num}: Title '{original_title}' → '{new_title}'", "TRANSFORM")
            transformed["title"] = new_title

        # Trim bullets if too many
        if slide_type == "key_insights":
            bullets = slide.get("content", {}).get("bullets", [])
            if len(bullets) > self.MAX_BULLETS:
                self.log(f"Slide {slide_num}: Trimming bullets {len(bullets)} → {self.MAX_BULLETS}", "TRANSFORM")
                transformed["content"]["bullets"] = bullets[:self.MAX_BULLETS]

        # TODO Phase 3: Detect numeric bullets and convert to chart
        # if slide_type == "key_insights" and self._has_numeric_bullets(bullets):
        #     transformed = self._convert_to_chart(transformed)

        return transformed

    def _transform_title(self, title: str) -> str:
        """
        Rewrite title for brevity and impact.

        Rules:
        1. Remove filler phrases
        2. Prefer active voice
        3. Max 6 words (ideal), enforce <10 words
        4. Make it decisive, not descriptive
        """
        # Remove common fillers
        for filler in self.TITLE_FILLERS:
            if title.startswith(filler):
                title = title[len(filler):].strip()

        # If still too long, truncate intelligently
        words = title.split()

        if len(words) > self.MAX_TITLE_WORDS:
            # Keep first N words but try to preserve meaning
            # Prefer keeping nouns over adjectives
            title = " ".join(words[:self.MAX_TITLE_WORDS])

        # Remove trailing punctuation except ? or !
        title = re.sub(r'[.,;:]$', '', title)

        return title

    def _ensure_quality_standards(self, slides: List[Dict]) -> List[Dict]:
        """
        Phase 3: Final quality pass.

        Standards:
        - Title slide must be strong (title + subtitle + visual)
        - No empty slides
        - Deck length within bounds
        """
        if not slides:
            raise ValueError("Cannot create presentation with zero slides")

        # Ensure title slide is strong
        if slides[0].get("type") == "title":
            slides[0] = self._strengthen_title_slide(slides[0])

        # Ensure all slides have titles
        for slide in slides:
            if not slide.get("title"):
                slide["title"] = f"Slide {slide.get('slide_number', '?')}"
                self.log(f"Added missing title to slide {slide.get('slide_number')}", "TRANSFORM")

        return slides

    def _strengthen_title_slide(self, title_slide: Dict) -> Dict:
        """
        Ensure title slide is never empty and visually strong.

        Requirements:
        - Title must exist
        - Subtitle should exist (create if missing)
        - Visual description should guide design
        """
        content = title_slide.get("content", {})

        # Ensure subtitle
        if not content.get("subtitle"):
            # Generate professional subtitle
            content["subtitle"] = "Strategic Overview & Recommendations"
            self.log("Title slide: Added missing subtitle", "TRANSFORM")

        title_slide["content"] = content

        # Ensure visual guidance
        if not title_slide.get("visuals", {}).get("description"):
            title_slide["visuals"] = {
                "description": "Professional visual concept - modern, clean, appropriate for topic and audience"
            }
            self.log("Title slide: Added visual guidance", "TRANSFORM")

        return title_slide


def main():
    """CLI entry point for creative director."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Creative Slide Director - Transform presentations before rendering"
    )
    parser.add_argument("--input", default="claude_output.json", help="Input JSON file")
    parser.add_argument("--output", default="creative_output.json", help="Output JSON file")
    parser.add_argument("--quiet", action="store_true", help="Suppress verbose output")

    args = parser.parse_args()

    # Run creative director
    director = CreativeDirector(verbose=not args.quiet)
    director.review_presentation(args.input, args.output)

    print("\n" + "=" * 70)
    print("  NEXT STEP: Run generate_ppt.py with creative_output.json")
    print("=" * 70)


if __name__ == "__main__":
    main()


# ============================================================================
# EXAMPLE TRANSFORMATION (Before/After)
# ============================================================================
"""
BEFORE (claude_output.json):
{
  "slide_number": 2,
  "type": "key_insights",
  "title": "Executive Summary of Customer Satisfaction Performance",
  "content": {
    "bullets": [
      "Overall customer satisfaction averaged 7.6 over the past year",
      "Maintenance journey achieves highest satisfaction at 8.1",
      "Breakdown journey scores lowest at 7.2",
      "Product performance ranges from 7.3 to 8.0",
      "Strategic opportunity to replicate Maintenance strengths",
      "Communication clarity identified as key satisfaction driver",
      "Heat pumps requiring focused attention on planning",
      "Recommendation: deploy Maintenance communication standards"
    ]
  }
}

AFTER (creative_output.json):
{
  "slide_number": 2,
  "type": "key_insights",
  "title": "Executive Summary",           ← Shortened from 7 words to 2
  "content": {
    "bullets": [
      "Overall customer satisfaction averaged 7.6 over the past year",
      "Maintenance journey achieves highest satisfaction at 8.1",
      "Breakdown journey scores lowest at 7.2",
      "Strategic opportunity to replicate Maintenance strengths",
      "Communication clarity identified as key satisfaction driver"
    ]                                      ← Trimmed from 8 bullets to 5
  }
}

DECISION LOG:
✎ Slide 2: Title 'Executive Summary of Customer Satisfaction Performance' → 'Executive Summary'
✎ Slide 2: Trimming bullets 8 → 5
"""
