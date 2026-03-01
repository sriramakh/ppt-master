Build a complete 40-slide template PPTX for a given theme (no AI/LLM required).

User input: $ARGUMENTS
Expected: [theme_key] [company_name] [output_path]

Steps:
1. Parse $ARGUMENTS for theme, company name, and output path.
   - theme: one of 14 keys (default: corporate). Run /ppt-list-themes to see all.
   - company: company/org name to appear on slides (default: "Acme Corp")
   - output: file path (default: templates/{theme}.pptx)
2. Run:

```bash
cd "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master" && \
source .venv/bin/activate && \
python -c "
from pptmaster.builder.template_builder import build_template
path = build_template(
    theme_key='THEME_HERE',
    company_name='COMPANY_HERE',
    output_path='OUTPUT_HERE'
)
print(f'Template built: {path}')
print('40 slides: cover, TOC, 7 section dividers, 29 content slides, thank you.')
"
```

3. Report the output file path.

Template builds are deterministic and fast (no API calls). The 40-slide sequence covers:
- About Us: company overview, values, team, key facts
- Strategy: executive summary, KPI dashboard, SWOT, matrix quadrant, Venn diagram
- Process: linear/circular process, roadmap, funnel, pyramid, hub & spoke
- Data: bar/line/pie charts, comparison, data table, gauge dashboard
- Planning: milestone roadmap, kanban board, risk matrix
- Deliverables: two-column, three-column, highlight quote, infographic, icon grid
- Closing: next steps, call to action, thank you

Pre-built templates are in: /Volumes/Seagate Expansion Drive/Seagate AI/PPT Master/templates/
