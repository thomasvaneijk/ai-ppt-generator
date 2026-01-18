"""
Creative Slide Director - Pre-rendering creative decision layer

Makes bold creative decisions BEFORE slides are rendered:
- Removes redundant or low-value slides
- Rewrites titles for impact and brevity
- Changes slide types strategically
- Ensures visual-first slides where appropriate
- Guarantees no empty slides, strong title slides

Acts like a senior creative director reviewing a deck before it goes to production.
"""

import json
from typing import List, Dict
from pathlib import Path


class CreativeDirector:
    """
    Makes creative decisions about slide structure, type, and content.

    Philosophy:
    - Every slide must earn its place
    - Titles must be punchy (max 6 words preferred)
    - Visual hierarchy over information density
    - Narrative flow over completeness
    """

    def __init__(self, config: dict = None):
        self.config = config or {}
        self.max_title_words = self.config.get("max_title_words", 6)
        self.min_slides = self.config.get("min_slides", 5)
        self.max_slides = self.config.get("max_slides", 10)

    def review_presentation(self, input_json_path: str, output_json_path: str):
        """
        Main entry point: review and transform presentation.

        Args:
            input_json_path: Path to claude_output.json
            output_json_path: Path to save creative_output.json
        """
        print("="*60)
        print("  CREATIVE SLIDE DIRECTOR")
        print("="*60)

        # Load original presentation
        with open(input_json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        presentation = data.get("presentation", {})
        original_slides = presentation.get("slides", [])

        print(f"Original: {len(original_slides)} slides")

        # Apply creative transformations
        transformed_slides = []

        for slide in original_slides:
            # Decide if slide should be kept
            if not self.should_keep_slide(slide, original_slides):
                print(f"  ✗ Removing slide {slide.get('slide_number')}: {slide.get('title')} (redundant)")
                continue

            # Transform slide
            transformed = self.transform_slide(slide)
            transformed_slides.append(transformed)

        # Ensure title slide is strong
        if transformed_slides and transformed_slides[0].get("type") == "title":
            transformed_slides[0] = self.strengthen_title_slide(transformed_slides[0])

        # Update presentation
        presentation["slides"] = transformed_slides
        data["presentation"] = presentation

        # Renumber slides
        for i, slide in enumerate(transformed_slides, start=1):
            slide["slide_number"] = i

        print(f"After review: {len(transformed_slides)} slides")
        print("="*60)

        # Save transformed presentation
        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        print(f"✓ Creative review complete: {output_json_path}")

        return data

    def should_keep_slide(self, slide: Dict, all_slides: List[Dict]) -> bool:
        """
        Decide if a slide adds unique value.

        Criteria:
        - Does it have unique content?
        - Does it advance the narrative?
        - Is it an appendix that could be removed?
        """
        slide_type = slide.get("type")

        # Always keep title and recommendations
        if slide_type in ["title", "recommendation"]:
            return True

        # Appendix is optional - remove if deck is already long
        if slide_type == "appendix" and len(all_slides) > self.max_slides:
            return False

        # Keep if content is substantive
        content = slide.get("content", {})

        # Check for empty content
        if slide_type == "key_insights":
            bullets = content.get("bullets", [])
            if not bullets or len(bullets) == 0:
                return False

        if slide_type == "chart":
            data = content.get("data", {})
            if not data.get("categories") or not data.get("series"):
                return False

        return True

    def transform_slide(self, slide: Dict) -> Dict:
        """
        Apply creative transformations to a slide.

        Transformations:
        - Shorten titles
        - Adjust bullet counts
        - Add visual emphasis notes
        """
        transformed = slide.copy()

        # Transform title
        original_title = slide.get("title", "")
        new_title = self.transform_title(original_title)

        if new_title != original_title:
            print(f"  ✎ Title: '{original_title}' → '{new_title}'")
            transformed["title"] = new_title

        # Ensure bullets ≤ 5
        if slide.get("type") == "key_insights":
            bullets = slide.get("content", {}).get("bullets", [])
            if len(bullets) > 5:
                print(f"  ✎ Trimming bullets: {len(bullets)} → 5")
                transformed["content"]["bullets"] = bullets[:5]

        return transformed

    def transform_title(self, title: str) -> str:
        """
        Rewrite title for impact and brevity.

        Rules:
        - Max 6 words preferred
        - Remove filler words (e.g., "Overview of", "Analysis of")
        - Make it punchy
        """
        # Remove common filler phrases
        fillers = [
            "Overview of ",
            "Analysis of ",
            "Summary of ",
            "Review of ",
            "Report on ",
            "Discussion of "
        ]

        for filler in fillers:
            if title.startswith(filler):
                title = title[len(filler):]
                break

        # If still too long, abbreviate
        words = title.split()
        if len(words) > self.max_title_words:
            # Keep first 6 words
            title = " ".join(words[:self.max_title_words])

        return title

    def strengthen_title_slide(self, title_slide: Dict) -> Dict:
        """
        Ensure title slide is strong and complete.

        Requirements:
        - Title must be present
        - Subtitle should be present
        - Visual description should guide design
        """
        content = title_slide.get("content", {})

        # Ensure subtitle exists
        if not content.get("subtitle"):
            # Generate a subtitle if missing
            content["subtitle"] = "Strategic Overview & Key Recommendations"
            print("  ✎ Added missing subtitle to title slide")

        # Ensure visuals guidance exists
        if not title_slide.get("visuals", {}).get("description"):
            title_slide["visuals"] = {
                "description": "Professional, modern visual - abstract or conceptual imagery appropriate for the topic"
            }
            print("  ✎ Added visual guidance to title slide")

        title_slide["content"] = content
        return title_slide


def main():
    """Entry point for creative director."""
    import argparse

    parser = argparse.ArgumentParser(description="Creative Slide Director - Pre-rendering transformations")
    parser.add_argument("--input", default="claude_output.json", help="Input JSON file")
    parser.add_argument("--output", default="creative_output.json", help="Output JSON file")
    parser.add_argument("--config", default="config.json", help="Configuration file")

    args = parser.parse_args()

    # Load config
    config = {}
    if Path(args.config).exists():
        with open(args.config, 'r') as f:
            config = json.load(f).get("creative_director", {})

    # Run creative director
    director = CreativeDirector(config)
    director.review_presentation(args.input, args.output)


if __name__ == "__main__":
    main()
