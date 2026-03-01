Build a single slide as a standalone PPTX for rapid prototyping and visual preview.

User input: $ARGUMENTS
Expected format: [slide_type] [theme] or just describe what you want.

Steps:
1. Parse $ARGUMENTS for slide_type and theme_key. If not clear, ask the user.
   - slide_type: any of the 32 types from /ppt-list-slides (e.g. venn_diagram, kanban_board, funnel_diagram)
   - theme_key: any of the 14 themes (default: corporate)
2. Choose a sensible output path like: /tmp/preview_{slide_type}_{theme_key}.pptx
3. Run:

```bash
cd "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master" && \
source .venv/bin/activate && \
python -c "
import sys, json
sys.path.insert(0, 'src')
from pptmaster.builder.ai_builder import _BUILDER_MAP
from pptmaster.builder.themes import get_theme
from pptmaster.builder.template_builder import _blank_prs
from pptmaster.builder.design_system import SLIDE_W, SLIDE_H
from pptmaster.content.builder_content_gen import SLIDE_CATALOG
from pptx.util import Emu
import copy

slide_type = 'SLIDE_TYPE_HERE'
theme_key = 'THEME_HERE'
output_path = 'OUTPUT_HERE'

theme = get_theme(theme_key)
builder_fn = _BUILDER_MAP.get(slide_type)
if not builder_fn:
    print(f'Unknown slide type: {slide_type}')
    sys.exit(1)

prs = _blank_prs()
layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(layout)
builder_fn(slide, theme=theme)
prs.save(output_path)
print(f'Preview saved: {output_path}')
"
```

4. Report the output path and describe what the slide shows.
5. Offer to try a different theme or adjust the content.

This is useful for quickly previewing how a slide type looks in a given theme before committing to a full presentation build.
