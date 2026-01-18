# AI PowerPoint Generator with Visual Feedback Loop

An autonomous system that generates professional, story-driven PowerPoint presentations from structured input with iterative visual quality improvement.

## 🎯 What This Does

**Input:** Topic, audience, content, and data
**Output:** Professional PowerPoint presentation with consultant-grade quality

The system:
1. Designs a logical slide narrative
2. Generates structured JSON (single source of truth)
3. Renders PowerPoint using `python-pptx`
4. (Optional) Iteratively improves visual quality based on rendered slide critique

---

## 🚀 Quick Start (Simple Workflow)

### For Colleagues: Generate a Presentation

**Option 1: Using the auto-sync helper (recommended)**

```powershell
.\run.ps1
```

This automatically syncs the latest version and runs the full creative pipeline.

**Option 2: Manual execution**

```powershell
# Sync latest version
git pull origin claude/powerpoint-generation-mvp-HdREb

# Run full pipeline (creative director + rendering)
python generate.py
```

**Option 3: Direct rendering (skip creative review)**

```powershell
# Generate directly from claude_output.json
python generate_ppt.py
```

**Result:** `output.pptx` will be created in the current directory.

---

## 📋 How to Request a New Presentation

Edit `claude_output.json` with your content, or provide input to the AI architect in this format:

```
Topic: [Your topic]
Audience: [senior management / clients / mixed]
Length: ~[N] slides
Content:
- [Your data points]
- [Key messages]
- [Numbers and facts]
```

**Example:**

```
Topic: Q4 Sales Performance
Audience: Senior management
Length: ~8 slides
Content:
- Total revenue: €2.3M (up 12% from Q3)
- Top product: Solar panels (45% of revenue)
- Key challenge: Installation delays in December
- Recommendation: Hire 2 additional installers
```

The AI will:
- Design the slide flow
- Choose appropriate slide types (title, bullets, charts, two-column, recommendations)
- Generate complete JSON
- Render professional PowerPoint

---

## 🏗️ System Architecture

### Core Flow with Creative Director (Phase 2)

```
claude_output.json          # Raw content from AI
       ↓
creative_director.py        # ✨ PRE-RENDER CREATIVE LAYER ✨
       ├─ Remove weak slides
       ├─ Rewrite titles for impact
       ├─ Trim bullets to max 5
       ├─ Ensure no empty slides
       └─ Strengthen title slide
       ↓
creative_output.json        # Refined, production-ready content
       ↓
generate_ppt.py             # Renderer (JSON → PowerPoint)
       ↓
output.pptx                 # Final presentation
```

### What the Creative Director Does

The **Creative Director** acts like a senior consultant reviewing a deck before production.

**It makes bold decisions:**
- ✗ **Removes** redundant or weak slides
- ✎ **Rewrites** long titles (max 6 words ideal)
- ✎ **Trims** bullet lists to max 5 points
- ✎ **Ensures** title slide is never empty (adds subtitle + visual guidance)
- ✓ **Preserves** critical slides (title, recommendations)

**Example transformations:**
```
BEFORE: "Executive Summary of Customer Satisfaction Performance" (7 words)
AFTER:  "Executive Summary" (2 words)

BEFORE: 8 bullets on one slide
AFTER:  5 bullets (top insights only)

BEFORE: Empty title slide with no subtitle
AFTER:  Strong title + professional subtitle + visual guidance
```

**Philosophy:**
- Every slide must earn its place
- Visual hierarchy beats information density
- Story flow beats completeness
- Remove, don't dilute

### Visual Feedback Loop (Optional, Windows-only)

```
output.pptx
       ↓
render_slides.py       # PowerPoint → PNG images (requires Windows + PowerPoint)
       ↓
renders/slide_*.png
       ↓
critique_slides.py     # Visual analysis → critique.json
       ↓
critique.json          # Feedback on visual quality
       ↓
[iterate.py]           # Apply improvements and repeat
```

---

## 🛠️ Requirements

### Basic Usage (All Platforms)

- Python 3.8+
- `pip install -r requirements.txt`

### Visual Feedback Loop (Windows Only)

- Microsoft PowerPoint installed
- `pip install pywin32`

---

## 📁 File Structure

```
ai-ppt-generator/
├── generate.py              # 🎯 MAIN ENTRY POINT (orchestrates full pipeline)
├── creative_director.py     # ✨ Pre-render creative intelligence layer
├── generate_ppt.py          # Renderer (JSON → PPT)
│
├── claude_output.json       # Input: raw presentation content
├── creative_output.json     # Intermediate: refined content after creative review
├── output.pptx              # Output: final presentation
│
├── template.pptx            # Corporate template (styling, layouts)
├── config.json              # Configuration (iterations, scoring, creative settings)
├── run.ps1                  # Helper script (auto-sync + generate)
│
├── render_slides.py         # Slide renderer (PPT → PNG)
├── critique_slides.py       # Visual critic (PNG → critique)
├── iterate.py               # Iteration orchestrator
│
├── renders/                 # Rendered slide images (Windows only)
└── critique.json            # Visual quality feedback (Windows only)
```

---

## 🎨 Supported Slide Types

The system fully supports 6 slide types:

1. **title** - Title slide with optional subtitle ✓
2. **key_insights** - Bullet points (max 5 per slide) ✓
3. **chart** - Data visualization (column, bar, line, pie) ✓
4. **two_column** - Left/right content split (TODO)
5. **recommendation** - Strategic recommendation with rationale and next steps ✓
6. **appendix** - Supplementary information ✓

**Note:** Two-column slides are not yet implemented. Use key_insights or split into multiple slides as a workaround.

---

## 🔧 Configuration

Edit `config.json` to adjust:

```json
{
  "max_iterations": 3,        // How many improvement loops
  "target_score": 8.0,        // Stop when average score ≥ 8/10
  "rendering": {
    "resolution_width": 1920,
    "resolution_height": 1080
  }
}
```

---

## 🚨 Known Limitations

### Current State

1. **Visual feedback loop requires Windows**
   - Slide rendering uses PowerPoint COM automation
   - Not available on macOS/Linux

2. **Critique is placeholder-based**
   - Actual visual analysis not yet implemented
   - Heuristic scoring returns default values
   - Production version would use PIL/OCR/vision APIs

3. **Improvement application is manual**
   - `iterate.py` shows where improvements would be applied
   - Automatic JSON modification not yet implemented
   - Requires manual editing based on critique feedback

4. **Two-column slide type not yet implemented**
   - Additional types (e.g., timeline, process flow) not yet supported
   - Workaround: adapt content to key_insights or split into multiple slides

5. **No image generation**
   - Visual descriptions in JSON are documentation only
   - Actual image insertion not implemented
   - Workaround: manually add images to template after generation

### Recommended Next Steps

- Implement actual image analysis in `critique_slides.py` (PIL, OCR)
- Build automatic improvement engine in `iterate.py`
- Add support for custom slide types
- Implement image placeholder support

---

## 📖 Usage Examples

### Example 1: Full Pipeline (Recommended)

```powershell
python generate.py
```

This runs:
1. Creative Director review (claude_output.json → creative_output.json)
2. PowerPoint generation (creative_output.json → output.pptx)

Result: `output.pptx` created with full creative intelligence applied

### Example 2: Direct Generation (Skip Creative Review)

```powershell
python generate_ppt.py
```

Result: `output.pptx` created directly from `claude_output.json` (or `creative_output.json` if it exists)

### Example 3: Run Full Feedback Loop (Windows Only)

```powershell
python iterate.py
```

This runs:
1. `generate.py` → creates `creative_output.json` and `output.pptx`
2. `render_slides.py` → creates `renders/slide_*.png`
3. `critique_slides.py` → creates `critique.json`
4. Checks if score ≥ 8/10, repeats if needed (max 3 iterations)

### Example 4: Manual Rendering

```powershell
# Generate presentation with creative review
python generate.py

# Render slides (Windows + PowerPoint required)
python render_slides.py --input output.pptx --output renders

# Critique slides
python critique_slides.py --input renders --output critique.json
```

---

## 🤝 For Colleagues: How to Use This Tool

**Step 1:** Provide your presentation brief (topic, audience, content, data)

**Step 2:** Run `.\run.ps1` (or `python generate_ppt.py`)

**Step 3:** Open `output.pptx` and review

**Step 4:** (Optional) Request adjustments by editing `claude_output.json` or providing feedback

**That's it.** The system handles slide design, layout, and formatting.

---

## 🐛 Troubleshooting

### "Template file not found"
- Ensure `template.pptx` exists in the project directory
- Download or create a template with the required 6 layouts (indices 0-5)

### "Module 'pptx' not found"
- Run: `pip install -r requirements.txt`

### "win32com.client not available"
- Windows only: `pip install pywin32`
- Or skip rendering and use PowerPoint manually

### Git sync issues
- Use `.\run.ps1` to auto-sync before running
- Or manually: `git pull origin claude/powerpoint-generation-mvp-HdREb`

---

## 📝 JSON Schema Reference

See `claude_output.json` for the complete schema.

**Minimal example:**

```json
{
  "presentation": {
    "title": "Presentation Title",
    "subtitle": "Subtitle",
    "language": "en",
    "audience": "senior_management",
    "slides": [
      {
        "slide_number": 1,
        "type": "title",
        "title": "My Title",
        "content": { "subtitle": "My Subtitle" },
        "visuals": {},
        "speaker_note": "Optional speaker note"
      }
    ]
  }
}
```

---

## 🔐 Corporate Template

The system respects `template.pptx` for:
- Fonts (e.g., Calibri)
- Colors (e.g., #003366, #00A3E0)
- Logo placement
- Layout structure

**Do NOT override styling in code.** All visual styling comes from the template.

---

## 📞 Support

For issues or questions:
- Check this README
- Review `claude_output.json` for schema examples
- Inspect `critique.json` for visual feedback (if using Windows feedback loop)

---

## 🎯 Design Principles

1. **JSON is the single source of truth** - All content and structure defined in JSON
2. **Template defines styling** - Code never overrides colors/fonts
3. **Autonomous creativity** - AI makes design decisions (slide types, flow, layout)
4. **No invented data** - Only structure and visualize provided content
5. **Professional quality** - Consultant-grade output, not template filling

---

**Built for:** Internal colleagues generating presentations for management and external stakeholders
**Platform:** Windows (with optional Linux/macOS basic support)
**Status:** MVP with visual feedback loop framework in place
