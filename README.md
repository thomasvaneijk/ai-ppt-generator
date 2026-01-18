# AI PowerPoint Generator with Visual Feedback Loop

An autonomous system that generates professional, story-driven PowerPoint presentations from structured input with iterative visual quality improvement.

## 🎯 What This Does

**Input:** Topic, audience, content, and data
**Output:** Professional PowerPoint presentation with consultant-grade quality

The system:
1. Designs a logical slide narrative
2. Generates structured JSON (single source of truth)
3. **Creative Director** - Pre-render intelligence (removes weak slides, rewrites titles, trims bullets)
4. Renders PowerPoint using `python-pptx` with **chart storytelling** (data labels, emphasis)
5. **AI Vision Critique** - Analyzes rendered slides using Claude Vision API (alignment, hierarchy, density)
6. **Improvement Engine** - Automatically applies fixes based on critique
7. Iterates until visual quality threshold reached (≥8.5/10)

---

## 🚀 Quick Start (Simple Workflow)

### For Colleagues: Generate a Presentation

**SYNC FIRST (After I Make Changes)**

```powershell
# Windows
.\sync.ps1

# macOS/Linux
./sync.sh
```

This pulls the latest code without running anything. Run this whenever you see files are out of sync.

---

**THEN GENERATE**

**Option 1: Auto-sync + Generate (recommended)**

```powershell
.\run.ps1
```

This automatically syncs the latest version AND runs the full creative pipeline.

**Option 2: Manual execution (when already synced)**

```powershell
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

### Phase 3+ AI Vision Feedback Loop (IMPLEMENTED)

```
output.pptx
       ↓
render_slides.py           # PowerPoint → PNG images (requires Windows + PowerPoint)
       ↓
renders/slide_*.png
       ↓
vision_feedback.py         # 🤖 AI vision critique using Claude Vision API
       ├─ Analyzes: alignment, hierarchy, text density, balance
       ├─ Scores: 0-10 per slide (target ≥8.5)
       └─ Outputs: concrete actionable fixes
       ↓
critique.json              # AI-powered feedback
       ↓
improvement_engine.py      # 🔧 Automatically applies fixes
       ├─ Trim bullets
       ├─ Shorten titles
       ├─ Enlarge charts
       ├─ Add visual emphasis
       └─ Split/merge slides
       ↓
improved_output.json       # Enhanced presentation JSON
       ↓
[Re-render and iterate until score ≥8.5]
```

**Key Features:**
- ✅ **Real AI Vision** - Claude analyzes actual slide images, not heuristics
- ✅ **Automatic Improvements** - Fixes applied without manual intervention
- ✅ **Chart Storytelling** - Data labels, legend control, gridline simplification
- ✅ **Quality Threshold** - Iterates until presentation reaches consultant-grade (≥8.5/10)
- ✅ **Fallback Support** - Works without API key (heuristic mode)

---

## 🛠️ Requirements

### Basic Usage (All Platforms)

- Python 3.8+
- `pip install -r requirements.txt`
- Includes: `python-pptx`, `anthropic`

### AI Vision Feedback (Phase 3+)

- **Anthropic API key** (for Claude Vision)
- Set environment variable: `export ANTHROPIC_API_KEY=your-key`
- Get API key from: https://console.anthropic.com/
- **Note:** System works without API key (fallback to heuristic critique)

### Visual Rendering (Windows Only)

- Microsoft PowerPoint installed
- `pip install pywin32`
- Required for: `render_slides.py`, full iteration loop

---

## 📁 File Structure

```
ai-ppt-generator/
├── generate.py              # 🎯 MAIN ENTRY POINT (orchestrates full pipeline)
├── creative_director.py     # ✨ Pre-render creative intelligence layer
├── generate_ppt.py          # Renderer (JSON → PPT) with chart storytelling
│
├── claude_output.json       # Input: raw presentation content
├── creative_output.json     # Intermediate: refined content after creative review
├── improved_output.json     # Intermediate: enhanced after AI critique (auto-generated)
├── output.pptx              # Output: final presentation
│
├── template.pptx            # Corporate template (styling, layouts)
├── config.json              # Configuration (iterations, scoring, creative settings)
├── run.ps1                  # Helper script (auto-sync + generate)
├── sync.ps1                 # Quick sync without generation
│
├── render_slides.py         # Slide renderer (PPT → PNG, Windows only)
├── vision_feedback.py       # 🤖 AI vision critique (Claude Vision API)
├── critique_slides.py       # Heuristic critic (fallback, no API key needed)
├── improvement_engine.py    # 🔧 Automatic improvement application
├── iterate.py               # Full iteration orchestrator
│
├── renders/                 # Rendered slide images (Windows only)
└── critique.json            # Visual quality feedback
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

### Current State (Phase 3+ IMPLEMENTED)

1. **Visual feedback loop requires Windows** ⚠️
   - Slide rendering uses PowerPoint COM automation
   - Not available on macOS/Linux
   - **Status:** Platform limitation, cannot be resolved without PowerPoint

2. **✅ AI Vision Critique IMPLEMENTED (Phase 3+)**
   - Uses Claude Vision API to analyze actual slide images
   - Judges: alignment, hierarchy, text density, balance, template-likeness
   - Outputs: concrete actionable fixes (not vague suggestions)
   - **Fallback:** Heuristic critique works without API key

3. **✅ Automatic Improvement Engine IMPLEMENTED (Phase 3+)**
   - Reads critique.json and applies fixes automatically
   - Actions: trim bullets, shorten titles, enlarge charts, split slides
   - Modifies JSON and triggers re-rendering
   - **Status:** Fully autonomous improvement loop

4. **✅ Chart Storytelling IMPLEMENTED (Phase 3+)**
   - Data labels (show/hide based on JSON config)
   - Legend control (remove when not needed)
   - Gridline simplification (less clutter)
   - Chart enlargement (visual emphasis)
   - **Status:** Charts now tell stories, not just display data

5. **Two-column slide type not yet implemented** ⚠️
   - Additional types (e.g., timeline, process flow) not yet supported
   - Workaround: adapt content to key_insights or split into multiple slides
   - **Priority:** Low (5 of 6 core types implemented)

6. **No image generation** ⚠️
   - Visual descriptions in JSON are documentation only
   - Actual image insertion not implemented
   - Workaround: manually add images to template after generation
   - **Future:** DALL-E/Stable Diffusion integration planned

### Recommended Next Steps

- ~~Implement actual image analysis~~ ✅ DONE (vision_feedback.py using Claude Vision)
- ~~Build automatic improvement engine~~ ✅ DONE (improvement_engine.py)
- ~~Add chart storytelling~~ ✅ DONE (data labels, legend control, gridlines)
- Implement two-column slide type handler
- Add image generation (DALL-E/Stable Diffusion integration)
- Implement visual emphasis shapes (background boxes, section dividers)
- Add support for custom slide types (timeline, process flow, comparison matrix)

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

### Example 3: Run Full AI Vision Feedback Loop (Phase 3+, Windows Only)

```powershell
# Set API key first
$env:ANTHROPIC_API_KEY="your-api-key-here"

# Run full iteration loop
python iterate.py
```

This runs:
1. `generate.py` → creates `creative_output.json` and `output.pptx`
2. `render_slides.py` → creates `renders/slide_*.png`
3. `vision_feedback.py` → AI vision critique → `critique.json`
4. `improvement_engine.py` → applies fixes → `improved_output.json`
5. Checks if score ≥ 8.5/10, repeats if needed (max 3 iterations)

**Without API key:** Falls back to heuristic critique (`critique_slides.py`)

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
