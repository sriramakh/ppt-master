# PPT Master — Internal Architecture

This document explains how the application works internally: the module structure, data flow, key abstractions, and how the various engines interact.

---

## High-Level Overview

PPT Master is built around two main pipelines:

```
┌─────────────────────────────────────────────────────────────────┐
│ AI BUILD PIPELINE (ai_builder.py)                               │
│                                                                  │
│  topic + config                                                  │
│      │                                                           │
│      ▼                                                           │
│  ┌─────────────────────┐                                         │
│  │  Content Generation │  builder_content_gen.py                 │
│  │  (LLM → JSON)       │  → selected_slides + sections + content │
│  └─────────────────────┘                                         │
│      │                                                           │
│      ▼                                                           │
│  ┌─────────────────────┐                                         │
│  │  PPTX Assembly      │  ai_builder.py                          │
│  │  (JSON → PPTX)      │  → for each slide: call slide module    │
│  └─────────────────────┘                                         │
│      │                                                           │
│      ▼                                                           │
│  output.pptx                                                     │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│ STATIC TEMPLATE PIPELINE (template_builder.py)                  │
│                                                                  │
│  theme_key                                                       │
│      │                                                           │
│      ▼                                                           │
│  TemplateTheme (colors, content, UXStyle)                        │
│      │                                                           │
│      ▼                                                           │
│  Build all 40 slides in sequence → output.pptx                  │
└─────────────────────────────────────────────────────────────────┘
```

---

## Module Map

```
src/pptmaster/
│
├── config.py                   # Pydantic settings (env vars)
├── models.py                   # Shared Pydantic data models
├── mcp_server.py               # FastMCP server (11 tools)
│
├── builder/                    # ── Core PPTX generation ──
│   ├── design_system.py        # Constants: dimensions, colors, fonts, grid
│   ├── helpers.py              # 27 drawing primitives
│   ├── ux_styles.py            # 14 UXStyle dataclasses
│   ├── themes.py               # 14 TemplateTheme factories + THEME_MAP
│   ├── template_builder.py     # 40-slide static build orchestrator
│   ├── ai_builder.py           # AI-driven selective build + _BUILDER_MAP
│   ├── icon_template_builder.py # Icon showcase template builder
│   └── slides/                 # One module per slide type
│       ├── s01_cover.py
│       ├── s02_toc.py
│       ├── s03_section_divider.py
│       ├── s04_company_overview.py
│       ├── ...
│       ├── s30_thank_you.py
│       ├── s31_funnel_diagram.py   # New visual types
│       ├── s32_pyramid_hierarchy.py
│       ├── s33_venn_diagram.py
│       ├── s34_hub_spoke.py
│       ├── s35_milestone_roadmap.py
│       ├── s36_kanban_board.py
│       ├── s37_matrix_quadrant.py
│       ├── s38_gauge_dashboard.py
│       ├── s39_icon_grid.py
│       ├── s40_risk_matrix.py
│       └── _chart_helpers.py   # DesignProfile bridge for chart renderer
│
├── content/                    # ── LLM content generation ──
│   ├── llm_client.py           # Provider-agnostic OpenAI SDK wrapper
│   ├── builder_content_gen.py  # SLIDE_CATALOG + prompt + JSON validation
│   ├── content_engine.py       # Template-based content engine
│   ├── openai_client.py        # Legacy direct OpenAI client
│   ├── input_processor.py      # PDF/DOCX/CSV/URL → text
│   ├── prompt_builder.py       # Prompt assembly helpers
│   └── slide_planner.py        # Slide sequence planner
│
├── analyzer/                   # ── PPTX template analysis ──
│   ├── template_analyzer.py    # Main analyzer orchestrator
│   ├── theme_extractor.py      # Extracts colors/fonts from PPTX theme XML
│   ├── layout_extractor.py     # 16-rule layout classifier
│   ├── shape_extractor.py      # Shape metadata extractor
│   ├── icon_extractor.py       # Icon/image asset extractor
│   └── potx_handler.py         # .potx → .pptx patch (Content_Types.xml)
│
├── composer/                   # ── Legacy template-based composer ──
│   ├── slide_composer.py       # Main composer entry point
│   ├── chart_renderer.py       # python-pptx chart builder
│   ├── table_renderer.py       # Table layout renderer
│   ├── shape_renderer.py       # Shape drawing
│   ├── text_renderer.py        # Text box renderer
│   ├── image_renderer.py       # Image placement
│   ├── layout_matcher.py       # Matches content to layouts
│   └── theme_applier.py        # Applies extracted theme to slides
│
├── assets/                     # ── Asset management ──
│   ├── color_utils.py          # hex_to_rgb, tint(), shade()
│   ├── icon_manager.py         # SVG icon toolkit interface
│   ├── raster_icon_manager.py  # PNG icon library interface
│   ├── icon_generator.py       # Icon generation utilities
│   └── image_handler.py        # Image resize/crop/embed
│
├── cli/
│   └── app.py                  # Typer CLI commands
│
└── web/
    └── streamlit_app.py        # Streamlit web UI
```

---

## The Design System (`design_system.py`)

Everything is measured in **EMU** (English Metric Units): 914400 EMU = 1 inch. All slide coordinates, widths, heights, and spacing use EMU. This makes layouts device-independent and consistent across PowerPoint, Keynote, and Google Slides.

**Slide dimensions:** 12192000 × 6858000 EMU (standard 16:9, 13.33" × 7.5")

**Layout structure (every content slide):**
```
┌──────────────────────────────────────────────────────────┐
│  0.3" — Title zone                                        │
│  [TITLE]  [optional accent line]                          │
├──────────────────────────────────────────────────────────┤
│  1.5" — Content area begins                               │
│                                                           │
│  [slide-specific content]                                 │
│                                                           │
├──────────────────────────────────────────────────────────┤
│  Footer zone (0.5" at bottom)                             │
│  [Company Name]           [Slide Number]                  │
└──────────────────────────────────────────────────────────┘
```

**`col_span(n_cols, col_idx)` helper:** Divides the content area into equal columns with standard gaps. Used by every multi-column slide — e.g., `col_span(3, 0)` returns the `(left, width)` of the first of three columns.

**`card_positions(n_cards)` helper:** Returns `(left, top, width, height)` tuples for a standard card grid layout (2-column or 3-column depending on count).

---

## Themes and UX Styles

### Two-layer design system

Themes and UXStyles are separate concerns composed together:

```
TemplateTheme (WHAT to show, what colors)
    ├── key, industry, company_name, tagline
    ├── primary, accent, secondary, light_bg  (60-30-10 color rule)
    ├── palette[6]  (chart/card accent colors)
    ├── font  (Inter for all themes)
    ├── content{}  (default text content for static builds)
    └── ux_style: UXStyle  (HOW to render it)

UXStyle (HOW to render each slide)
    ├── card_radius, card_shadow, card_accent  (card appearance)
    ├── margin_factor, gap_factor             (spacing)
    ├── title_align, title_caps, title_scale  (typography)
    ├── dark_mode  (dark background on all content slides)
    └── variant keys: cover, divider, process, timeline, quote,
                      cta, thankyou, kpi, funnel, pyramid, venn,
                      hub_spoke, milestone, kanban, matrix,
                      gauge, icon_grid, risk
```

**Variant dispatch:** Every high-impact slide module uses a `dispatch` dictionary that maps variant key → render function. When `build(slide, theme=t)` is called, the slide reads `t.ux_style.venn` (for example), looks it up in its dispatch dict, and calls the matching render function. This is how 14 themes can produce visually distinct slides from the same slide type module.

```python
# Example from s33_venn_diagram.py
dispatch = {
    "overlapping":    _overlapping,
    "labeled-rings":  _labeled_rings,
    "soft-overlap":   _soft_overlap,
    "scholarly-venn": _scholarly_venn,
    # ... 10 more variants
}
fn = dispatch.get(s.venn, _overlapping)
fn(slide, t, c, p)
```

### The 14 UXStyles

Each UXStyle name maps to a visual personality:

| Style | Character | Key traits |
|-------|-----------|------------|
| CLASSIC | Corporate conservative | Rounded cards, shadow, top accent bar, left-aligned titles |
| MINIMAL | Clean, spacious | Sharp cards, no shadow, bottom accent, centered titles |
| DARK | High-contrast dark mode | All slides dark background, glow effects |
| ELEVATED | Premium, financial | Thin borders, no shadows, wide margins |
| BOLD | High-energy, academic | Full-color card backgrounds, uppercase titles |
| EDITORIAL | Asymmetric, publication | Left color band, oversized title, minimal borders |
| GRADIENT | Luxury, dramatic | Gradient washes, pill-shaped cards |
| SPLIT | High-energy startup | Split-color layouts, bold CTAs |
| GEO | Structured, authoritative | Geometric accents, flag-stripe patterns |
| RETRO | Warm, approachable | Earth tones, film-strip decorations |
| MAGAZINE | Creative, dynamic | Full-bleed imagery, editorial typography |
| SCHOLARLY | Academic, rigorous | Dashed rules, serif-like spacing, footnote zones |
| LABORATORY | Scientific dark mode | Dark bg, chart-dominant layouts |
| DASHBOARD | Data-dense report | Compact margins, dense KPI grids |

---

## Slide Modules

Each slide module (`slides/s01_cover.py`, etc.) exports a single `build(slide, *, theme)` function.

```python
def build(slide, *, theme=None) -> None:
    from pptmaster.builder.themes import DEFAULT_THEME
    t = theme or DEFAULT_THEME  # TemplateTheme
    s = t.ux_style               # UXStyle
    c = t.content                # default content dict
    p = t.palette                # 6-color palette

    # Dispatch to variant renderer
    fn = DISPATCH.get(s.cover, _geometric)
    fn(slide, t, c, p)
```

Slide modules call drawing primitives from `helpers.py` and read all values from the theme. They never hardcode colors, fonts, or text — everything comes from the theme.

### Drawing Primitives (`helpers.py`)

27 helper functions that wrap python-pptx for common operations:

| Primitive | Purpose |
|-----------|---------|
| `add_textbox` | Styled text box with font/color/alignment |
| `add_rect` | Filled rectangle with optional border |
| `add_circle` | Filled ellipse/circle |
| `add_line` | Colored line with thickness |
| `add_hexagon` | Hexagonal shape |
| `add_gold_accent_line` | Theme accent underline for titles |
| `add_slide_title` | Standardized slide title placement |
| `add_styled_card` | Card with theme-specific radius/shadow/accent |
| `add_dark_bg` | Full-slide dark background fill |
| `add_background` | Light background fill |
| `add_pentagon` | Pentagon arrow shape |
| `add_funnel_tier` | Trapezoid tier for funnel diagrams |
| `add_notched_arrow` | Chevron process arrow |
| `add_block_arc` | Arc segment for circular diagrams |
| `add_snip_rect` | Rectangle with cut corner |
| `add_round_rect_2` | Rectangle with 2-corner rounding |
| `add_connector_arrow` | Connecting arrow between points |
| `add_color_cell` | Solid colored cell for tables/matrices |
| `add_semi_transparent_circle` | Translucent circle overlay |
| `add_themed_footer` | Footer bar with company name |
| `add_themed_slide_number` | Slide number badge |

All dimensions are in EMU. All colors accept hex strings (`"#1B2A4A"`).

---

## Content Generation (`content/`)

### SLIDE_CATALOG

`builder_content_gen.py` defines `SLIDE_CATALOG` — a dict of 32 slide types, each with:
- `description`: what the slide shows (used in LLM prompts)
- `content_keys`: dict of key → description explaining the expected data shape

This catalog drives both the LLM prompt (the AI knows exactly what data each slide needs) and validation (missing keys get defaults, unexpected keys are stripped).

### LLM Prompt Structure

`generate_builder_content()` builds a two-part prompt:

**System prompt:** Explains the slide catalog, content key schemas, and formatting rules. Tells the model to return a specific JSON schema.

**User prompt:** Topic, company, industry, audience, additional context, plus the full SLIDE_CATALOG descriptions.

**Expected JSON response:**
```json
{
  "selected_slides": ["executive_summary", "kpi_dashboard", "bar_chart", ...],
  "sections": [
    {"title": "Company Overview", "slides": ["company_overview", "team_leadership"]},
    {"title": "Financial Results", "slides": ["kpi_dashboard", "bar_chart"]}
  ],
  "content": {
    "exec_title": "Q4 2025 Earnings Highlights",
    "exec_bullets": ["Revenue grew 23% YoY...", ...],
    "kpi_title": "Key Performance Indicators",
    "kpis": [["Revenue", "$2.4B", "+23%", 0.88, "↑"], ...],
    ...
  }
}
```

### Response Repair

LLMs sometimes return imperfect JSON (especially reasoning models). `LLMClient` applies:
1. `_clean_response()` — strips `<think>...</think>` blocks and markdown code fences
2. `_repair_json()` — fixes common issues: unquoted numeric strings (`$6.2B` → `"$6.2B"`), trailing commas, control characters
3. `json.loads()` — final parse; exceptions propagate as errors

### Provider-Agnostic LLM Client (`llm_client.py`)

Uses the OpenAI Python SDK with the `base_url` parameter to support multiple providers:

```python
# OpenAI
client = OpenAI(api_key=openai_key, base_url=None)

# MiniMax (OpenAI-compatible API)
client = OpenAI(api_key=minimax_key, base_url="https://api.minimax.io/v1/")
```

Both resolve through `_PROVIDER_PRESETS`. Adding a new provider requires only a new entry in the presets dict.

---

## AI Build Pipeline (`ai_builder.py`)

### `ai_build_presentation()`

Entry point for the full pipeline:

```
ai_build_presentation(topic, company_name, theme_key, ...)
    │
    ├── get_theme(theme_key)           → TemplateTheme
    │
    ├── generate_builder_content(...)  → gen_result dict
    │   └── LLMClient.generate_json()
    │
    └── _build_selective_pptx(gen_result, theme, company_name, output_path)
        ├── Create blank Presentation() with 16:9 dimensions
        ├── build_s01(slide, theme=t)       # Cover
        ├── build_s02(slide, theme=t, ...)  # TOC (auto from sections)
        │   For each section in gen_result["sections"]:
        ├──   build_s03(slide, theme=t)     # Section divider
        │     For each slide_type in section["slides"]:
        ├──     _get_builder(slide_type)(slide, theme=t)  # Content slide
        ├── build_s30(slide, theme=t)       # Thank You
        └── prs.save(output_path)
```

### `_BUILDER_MAP`

A dict mapping slide type keys to module paths:
```python
_BUILDER_MAP = {
    "company_overview": "pptmaster.builder.slides.s04_company_overview",
    "venn_diagram":     "pptmaster.builder.slides.s33_venn_diagram",
    ...
}
```

`_get_builder(slide_type)` imports the module lazily and returns its `build` function. This avoids importing all 35 slide modules at startup.

### Content Injection

Each slide builder reads content from `theme.content` (default static content) but the AI build pipeline writes generated content into the theme before building. In practice, the `build()` functions accept content via the theme's content dict — the AI-generated content is merged into this dict before slides are built.

---

## Static Template Builder (`template_builder.py`)

`build_template(theme_key, company_name, output_path)` builds all 40 slides in a fixed sequence without any LLM calls:

```
Slide  1 — Cover
Slide  2 — Table of Contents
Slide  3 — Section divider: "About Us"
Slide  4 — Company Overview
Slide  5 — Our Values
Slide  6 — Team Leadership
Slide  7 — Key Facts & Figures
Slide  8 — Section divider: "Strategy"
Slide  9 — Executive Summary
Slide 10 — KPI Dashboard
Slide 11 — SWOT Matrix
Slide 12 — Matrix Quadrant
Slide 13 — Venn Diagram
Slide 14 — Section divider: "Process"
Slide 15 — Process Linear
Slide 16 — Process Circular
Slide 17 — Roadmap Timeline
Slide 18 — Funnel Diagram
Slide 19 — Pyramid Hierarchy
Slide 20 — Hub & Spoke
Slide 21 — Section divider: "Data"
Slide 22 — Bar Chart
Slide 23 — Line Chart
Slide 24 — Pie Chart
Slide 25 — Data Table
Slide 26 — Infographic Dashboard
Slide 27 — Gauge Dashboard
Slide 28 — Section divider: "Planning"
Slide 29 — Milestone Roadmap
Slide 30 — Kanban Board
Slide 31 — Risk Matrix
Slide 32 — Section divider: "Deliverables"
Slide 33 — Two Column
Slide 34 — Three Column
Slide 35 — Highlight Quote
Slide 36 — Text + Image
Slide 37 — Icon Grid
Slide 38 — Section divider: "Next Steps"
Slide 39 — Next Steps
Slide 40 — Thank You
```

Each call uses the theme's built-in `content` dict (populated by the `_*_content()` factory function in `themes.py`) so the template has realistic placeholder data out of the box.

---

## MCP Server (`mcp_server.py`)

Built with **FastMCP** from the `mcp[cli]` package. The server wraps the builder and content modules as callable tools following the Model Context Protocol.

### Tool Signatures

```python
@mcp.tool()
def list_themes() -> str: ...
    # Returns a formatted table of 14 themes

@mcp.tool()
def build_presentation(
    topic: str,
    company_name: str = "Acme Corp",
    theme_key: str = "corporate",
    audience: str = "",
    industry: str = "",
    additional_context: str = "",
    output_path: str = "output.pptx",
) -> str: ...
    # Runs the full AI pipeline, returns output path

@mcp.tool()
def build_single_slide(
    slide_type: str,
    theme_key: str = "corporate",
    content_json: str = "{}",
    output_path: str = "preview.pptx",
) -> str: ...
    # Builds 1-slide PPTX for rapid visual prototyping
```

### Path Resolution

All `output_path` arguments support relative paths. The `_resolve(path)` helper converts relative paths relative to the project root, and ensures parent directories exist:

```python
def _resolve(path: str) -> Path:
    p = Path(path)
    if not p.is_absolute():
        p = _PPT_ROOT / p
    p.parent.mkdir(parents=True, exist_ok=True)
    return p
```

### Starting the server

```bash
python -m pptmaster.mcp_server
```

The `__main__` block calls `mcp.run()` which starts an stdio-based MCP server process. The Agent (and other MCP clients) communicate with it over stdin/stdout.

---

## Analyzer Engine (`analyzer/`)

Used when analyzing an existing PPTX or POTX template (the `pptmaster analyze` command):

```
template_analyzer.py
    ├── potx_handler.py     → patches .potx Content_Types.xml to open as .pptx
    ├── theme_extractor.py  → walks slideMaster1 rels → correct theme XML
    │                          → extracts colors, fonts, effects
    ├── layout_extractor.py → 16-rule classifier:
    │                          inspects placeholder count, position, aspect ratio
    │                          → assigns layout type (title, content, two-col, etc.)
    ├── shape_extractor.py  → extracts shape metadata from each slide
    └── icon_extractor.py   → finds and extracts icon/image assets
```

**POTX handling:** POTX files have a `[Content_Types].xml` that references slides as `sld` instead of `pptx` types. `potx_handler.py` patches this in memory (as a ZIP) so python-pptx can open the file without errors.

**Theme extraction:** Must follow the relationship chain `slideMaster1.xml.rels → themeN.xml` — the first `theme*.xml` in the ZIP is not always the active theme for the master.

---

## Composer Engine (`composer/`)

The legacy engine for template-based generation (used by `pptmaster generate`, `from-doc`, `from-url` commands). It takes an analyzed template profile and populates a clone of that template with LLM-generated content:

```
SlideComposer
    ├── layout_matcher.py   → matches content type to template layout
    ├── chart_renderer.py   → adds charts to matched layout slides
    ├── table_renderer.py   → fills table placeholders
    ├── shape_renderer.py   → adds decorative shapes
    ├── text_renderer.py    → fills text placeholders
    ├── image_renderer.py   → places images
    └── theme_applier.py    → applies color/font overrides
```

The builder engine (new in v2) produces better results because it constructs slides from scratch with full design control, rather than trying to fill an existing template's placeholders.

---

## Data Flow: AI Build (end-to-end)

```
User: pptmaster ai-build --topic "Q4 Earnings" --theme finance

cli/app.py
  └─► ai_build_presentation("Q4 Earnings", theme_key="finance", ...)

builder/ai_builder.py
  ├─► get_theme("finance")
  │     └─► themes.py: _finance_theme()
  │           → TemplateTheme(primary="#14532D", ux_style=ELEVATED, ...)
  │
  ├─► generate_builder_content("Q4 Earnings", ...)
  │     └─► content/builder_content_gen.py
  │           ├─ Builds system + user prompt from SLIDE_CATALOG
  │           ├─► LLMClient.generate_json(prompt)
  │           │     └─► OpenAI API call (gpt-4o)
  │           │     └─► _clean_response() + _repair_json()
  │           │     └─► json.loads()
  │           ├─ Validates selected_slides against SLIDE_CATALOG
  │           └─► Returns gen_result dict
  │
  └─► _build_selective_pptx(gen_result, finance_theme, ...)
        ├─ prs = Presentation()  [blank 16:9]
        ├─ build_s01(slide, theme=finance_theme)    # Cover slide
        ├─ build_s02(slide, theme=finance_theme, sections=[...])  # TOC
        │
        │  Section: "Financial Results"
        ├─ build_s03(slide, theme=finance_theme)   # Section divider
        ├─ _get_builder("kpi_dashboard")(slide, theme=finance_theme)
        │     └─ s10_kpi_dashboard.py: build()
        │           ├─ s.kpi = "cards-row"  (from ELEVATED style)
        │           └─ _cards_row(slide, t, c, p)
        ├─ _get_builder("bar_chart")(slide, ...)
        │
        │  ... more slides ...
        │
        ├─ build_s30(slide, theme=finance_theme)   # Thank You
        └─ prs.save("output.pptx")

→ "output.pptx" (20 slides, finance theme, AI-generated content)
```

---

## Content Key System

Slide content is communicated via a flat string-keyed dict. The same dict is used for:
- **Static templates**: populated by `_*_content()` factory functions in `themes.py`
- **AI builds**: populated by the LLM (merged into `theme.content` before building)
- **MCP `build_from_content`**: provided directly by the caller

Each slide module reads its keys with `.get()` and defaults:
```python
title = c.get("kpi_title", "KPI Dashboard")
kpis  = c.get("kpis", _DEFAULT_KPIS)
```

**Key aliasing:** `_KEY_ALIASES` in `builder_content_gen.py` maps alternative key names (e.g., `"team_members"` → `"team"`) so the LLM's natural variations still resolve correctly.

---

## Configuration System (`config.py`)

Uses **pydantic-settings** with `PPTMASTER_` prefix:

```python
class Settings(BaseSettings):
    model_config = {"env_prefix": "PPTMASTER_", "extra": "ignore"}

    openai_api_key: str = ""
    model: str = "gpt-4o"
    builder_max_tokens: int = 12000
    llm_provider: str = "minimax"
    minimax_api_key: str = ""
    minimax_model: str = "MiniMax-M2.5"
    minimax_base_url: str = "https://api.minimax.io/v1/"
    ...
```

Loaded via `get_settings()` which reads `.env` automatically. All modules call `get_settings()` at the point of use (not at import time), so settings are always fresh.

---

## Key Design Decisions

**EMU everywhere.** Using python-pptx's native unit avoids conversion bugs. Helpers accept and return EMU. The design system constants are all in EMU.

**Variant dispatch over inheritance.** Rather than subclassing for each visual variant, each slide module has a dict mapping variant keys to render functions. Adding a new visual variant is one function + one dict entry, with zero impact on other variants.

**Flat content dict.** All slide content flows through a single `dict[str, Any]`. This makes it trivial to serialize (JSON), inspect, edit, and pass between LLM and builder without any object mapping layer.

**Lazy slide module imports.** `_get_builder(slide_type)` imports slide modules on demand via `importlib.import_module`. At startup, only the orchestrator loads — the 35 individual slide modules are imported only when their slide type is actually used.

**Theme = content + style.** Themes bundle both the visual style (UXStyle) and the default content (`content` dict). This means `build_template()` needs only a theme key — it gets both "what to show" and "how to render it" from a single object.

**Provider-agnostic via OpenAI SDK.** The `LLMClient` uses the OpenAI Python SDK with configurable `base_url`. Any OpenAI-compatible API (MiniMax, OpenRouter, local Ollama, etc.) works without code changes.

---

## Adding a New Slide Type

1. Create `src/pptmaster/builder/slides/sNN_my_type.py` with a `build(slide, *, theme)` function
2. Add to `_BUILDER_MAP` in `ai_builder.py`
3. Add to `SLIDE_CATALOG` in `builder_content_gen.py` with `description` and `content_keys`
4. Optionally add a variant key to `UXStyle` in `ux_styles.py` and a dispatch dict in your module
5. Optionally add it to the 40-slide sequence in `template_builder.py`

## Adding a New Theme

1. Create `_mytheme_theme()` factory in `themes.py`
2. Return a `TemplateTheme` with colors, a UXStyle, and a `_mytheme_content()` dict
3. Register in `THEME_MAP` with the theme key
4. Build and save to `templates/mytheme.pptx` via `build_template("mytheme")`
