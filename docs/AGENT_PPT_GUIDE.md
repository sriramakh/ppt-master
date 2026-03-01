# PPT Master — Agent-Driven Presentation Guide

AI-powered executive presentation generator with intelligent slide selection and 14 industry-specific themes.

## How It Works

The AI analyzes your topic, audience, and industry to:
1. **Select** which slides are relevant (from 23 available types)
2. **Organize** them into logical sections
3. **Generate** all content — text, metrics, charts, tables
4. **Build** a professionally styled PPTX

A pitch deck might get 12 content slides; a deep financial review might get 20+. The AI decides.

## Available Themes

| Key | Industry | UX Style | Primary | Accent |
|-----|----------|----------|---------|--------|
| `corporate` | General Corporate | classic | #1B2A4A | #C8A951 |
| `healthcare` | Healthcare | minimal | #0F4C5C | #E36414 |
| `technology` | Technology | dark | #0F172A | #06B6D4 |
| `finance` | Finance | elevated | #14532D | #D4AF37 |
| `education` | Education | bold | #881337 | #D4A574 |
| `sustainability` | Sustainability | editorial | #064E3B | #92400E |
| `luxury` | Luxury & Fashion | gradient | #1A1A2E | #B76E79 |
| `startup` | Startup / VC | split | #3B0764 | #EA580C |
| `government` | Government | geo | #1E3A5F | #B91C1C |
| `realestate` | Real Estate | retro | #374151 | #D97706 |
| `creative` | Creative & Media | magazine | #18181B | #BE185D |
| `academic` | Academic | editorial | #1E3A5F | #9B2335 |
| `research` | Research / Scientific | minimal | #0D4F4F | #D97706 |
| `report` | Reports / Analysis | elevated | #1F2937 | #0891B2 |

## Available Slide Types

The AI picks from these 23 content slide types:

| Slide Type | Description |
|---|---|
| `company_overview` | Company mission + 4 quick facts |
| `our_values` | 4 core values with descriptions |
| `team_leadership` | 4 executives with bios |
| `key_facts` | 6 large headline statistics |
| `sources` | Bibliography/references (4-8 cited sources) |
| `executive_summary` | 5 bullet points + 3 metrics |
| `kpi_dashboard` | 4-KPI dashboard with trends |
| `process_linear` | 5-step process flow |
| `process_circular` | 4-phase cycle diagram |
| `roadmap_timeline` | 5-milestone timeline |
| `swot_matrix` | SWOT analysis (4x3) |
| `bar_chart` | Grouped bar chart (3-8 categories, 1-4 series) |
| `line_chart` | Line chart (3-8 periods, 1-4 series) |
| `pie_chart` | Pie/donut chart (3-8 segments) |
| `comparison` | Side-by-side comparison (6 metrics) |
| `data_table` | 5-column data table (6 rows) |
| `two_column` | Two-column content layout |
| `three_column` | Three pillars/offerings |
| `highlight_quote` | Full-slide quote |
| `infographic_dashboard` | Mixed KPIs + chart + progress bars |
| `next_steps` | 4 action items with owners |
| `call_to_action` | Bold CTA + contacts |

Plus these are always included: Cover, Table of Contents, Section Dividers, Thank You.

## CLI Usage

### Generate an AI presentation

```bash
pptmaster ai-build \
  --topic "Q4 2025 Earnings Report" \
  --company "Meridian Capital" \
  --theme finance \
  --audience "Investors" \
  -o earnings_report.pptx
```

**Options:**

| Flag | Default | Description |
|------|---------|-------------|
| `--topic` | (required) | Presentation topic |
| `--company` | "Acme Corp" | Company name |
| `--theme` | "corporate" | Theme key (see table above) |
| `--audience` | "" | Target audience |
| `--industry` | "" | Industry context |
| `--context` | "" | Additional instructions for the AI |
| `-o` / `--output` | "output.pptx" | Output file path |
| `--provider` | "minimax" | LLM provider (minimax, openai) |

### List available themes

```bash
pptmaster list-themes
```

## MCP Server Setup (Agent)

### 1. Install the MCP dependency

```bash
cd "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master"
source .venv/bin/activate
pip install "mcp[cli]>=1.0.0"
```

### 2. Configure Agent

Add to `.agent/settings.json`:

```json
{
  "mcpServers": {
    "pptmaster": {
      "command": "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master/.venv/bin/python",
      "args": ["-m", "pptmaster.mcp_server"],
      "cwd": "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master"
    }
  }
}
```

### 3. Available MCP Tools

#### `generate_presentation`

Generate a complete presentation with AI-selected slides.

```
Parameters:
  topic (str, required): Presentation topic
  company_name (str): Company name (default: "Acme Corp")
  industry (str): Industry context
  audience (str): Target audience
  theme_key (str): Visual theme (default: "corporate")
  output_path (str): Output file path (default: "output.pptx")
  additional_context (str): Extra instructions for the AI
```

#### `list_themes`

List all 14 available themes with their visual styles. No parameters.

#### `generate_content_only`

Generate slide selection + content as JSON without building PPTX.

```
Parameters:
  topic (str, required): Presentation topic
  company_name (str): Company name
  industry (str): Industry context
  audience (str): Target audience
  additional_context (str): Extra instructions
```

Returns JSON with `selected_slides`, `sections`, and `content`.

#### `build_from_content`

Build a PPTX from a pre-generated content JSON.

```
Parameters:
  content_json (str, required): JSON with selected_slides, sections, content
  theme_key (str): Visual theme (default: "corporate")
  company_name (str): Company name
  output_path (str): Output file path
```

## Python API

```python
from pptmaster.builder.ai_builder import ai_build_presentation, list_available_themes

# Generate a presentation — AI picks slides automatically
path = ai_build_presentation(
    topic="AI Platform Product Launch",
    company_name="NexTech Systems",
    theme_key="technology",
    audience="CTO and Engineering Leadership",
    output_path="nextech_launch.pptx",
)

# List themes
themes = list_available_themes()
for t in themes:
    print(f"{t['key']:15s} {t['industry']}")
```

### Two-step workflow: generate content, then build

```python
from pptmaster.content.builder_content_gen import generate_builder_content
from pptmaster.builder.ai_builder import _build_selective_pptx
from pptmaster.builder.themes import get_theme

# Step 1: Generate (inspect/edit the result)
result = generate_builder_content(
    topic="Series B Funding Pitch",
    company_name="Velocity Ventures",
    industry="Startup / VC",
    audience="Venture Capitalists",
)
print(f"Selected: {result['selected_slides']}")
print(f"Sections: {[s['title'] for s in result['sections']]}")

# Step 2: Build from the result
theme = get_theme("startup")
_build_selective_pptx(result, theme, "Velocity Ventures", Path("pitch.pptx"))
```

## Example Use Cases

### 1. Finance — Earnings Report
```bash
pptmaster ai-build --topic "Q4 2025 Earnings Report" --company "Meridian Capital" --theme finance --audience "Investors" -o earnings.pptx
```

### 2. Technology — Product Launch
```bash
pptmaster ai-build --topic "AI Platform Product Launch" --company "NexTech Systems" --theme technology --audience "CTO/Engineers" -o product_launch.pptx
```

### 3. Healthcare — Patient Outcomes
```bash
pptmaster ai-build --topic "Patient Outcomes Improvement Program" --company "MedTech Solutions" --theme healthcare --audience "Hospital Board" -o patient_outcomes.pptx
```

### 4. Startup — Funding Pitch
```bash
pptmaster ai-build --topic "Series B Funding Pitch" --company "Velocity Ventures" --theme startup --audience "VCs" -o series_b.pptx
```

### 5. Sustainability — Carbon Roadmap
```bash
pptmaster ai-build --topic "Carbon Neutrality Roadmap 2030" --company "EcoForward Group" --theme sustainability --audience "Stakeholders" -o carbon_roadmap.pptx
```

### 6. Education — Digital Curriculum
```bash
pptmaster ai-build --topic "Digital Curriculum Transformation" --company "Horizon University" --theme education --audience "Faculty" -o curriculum.pptx
```

### 7. Luxury — Collection Launch
```bash
pptmaster ai-build --topic "2026 Luxury Collection Launch" --company "Maison Luxe" --theme luxury --audience "Retail Partners" -o luxury_launch.pptx
```

### 8. Government — Infrastructure
```bash
pptmaster ai-build --topic "National Infrastructure Modernization" --company "National Infrastructure Agency" --theme government --audience "Congressional Committee" -o infrastructure.pptx
```

### 9. Real Estate — Portfolio Review
```bash
pptmaster ai-build --topic "Commercial Portfolio Q1 Review" --company "Pinnacle Properties" --theme realestate --audience "Investors" -o portfolio.pptx
```

### 10. Creative — Brand Strategy
```bash
pptmaster ai-build --topic "Brand Refresh & Campaign Strategy" --company "Prism Media Group" --theme creative --audience "Brand Team" -o brand_strategy.pptx
```

### 11. Academic — Research Symposium
```bash
pptmaster ai-build --topic "Current Trends in AI and Future Directions" --company "Stanford AI Lab" --theme academic --audience "Faculty & Researchers" -o ai_trends.pptx
```

### 12. Research — Scientific Study
```bash
pptmaster ai-build --topic "Climate Change Impact on Global Agriculture" --company "Earth Sciences Institute" --theme research --audience "Scientific Community" -o climate_study.pptx
```

### 13. Report — Market Analysis
```bash
pptmaster ai-build --topic "Q1 2026 Global Market Analysis Report" --company "McKinsey Analytics" --theme report --audience "Executive Leadership" -o market_report.pptx
```

## Troubleshooting

### "LLM generation failed" error
- Check your API key is set in `.env` (`PPTMASTER_MINIMAX_API_KEY`)
- Verify network connectivity to `api.minimax.io`
- Try with `--provider openai` and ensure `PPTMASTER_OPENAI_API_KEY` is set

### Missing or placeholder content in slides
- The validator fills defaults for any missing keys — if the LLM was truncated, some slides may have generic content
- Try increasing `builder_max_tokens` in config or re-running

### Theme not found
- Run `pptmaster list-themes` to see valid theme keys
- Theme keys are lowercase: `realestate` not `real-estate`

### MCP server won't start
- Ensure `mcp[cli]` is installed: `pip install "mcp[cli]>=1.0.0"`
- Check the Python path in settings.json points to your venv
