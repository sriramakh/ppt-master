List all available slide types in PPT Master, with optional detail on a specific type.

User input: $ARGUMENTS

If $ARGUMENTS is empty, list all 32 selectable slide types:

```bash
cd "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master" && \
source .venv/bin/activate && \
python -c "
from pptmaster.content.builder_content_gen import SLIDE_CATALOG
print(f'{'TYPE':<26} DESCRIPTION')
print('-' * 80)
for slide_type, meta in SLIDE_CATALOG.items():
    print(f'{slide_type:<26} {meta.get(\"description\", \"\")}')
print()
print(f'{len(SLIDE_CATALOG)} selectable types.')
print('Always auto-included: cover, toc, section_divider (x7), thank_you')
print()
print('New in v2 (s31-s40): funnel_diagram, pyramid_hierarchy, venn_diagram,')
print('  hub_spoke, milestone_roadmap, kanban_board, matrix_quadrant,')
print('  gauge_dashboard, icon_grid, risk_matrix')
"
```

If $ARGUMENTS contains a slide type name, show full details for that type:

```bash
cd "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master" && \
source .venv/bin/activate && \
python -c "
import json
from pptmaster.content.builder_content_gen import SLIDE_CATALOG
slide_type = '$ARGUMENTS'.strip()
if slide_type in SLIDE_CATALOG:
    meta = SLIDE_CATALOG[slide_type]
    print(f'Slide type: {slide_type}')
    print(f'Description: {meta.get(\"description\", \"\")}')
    print(f'Content keys: {json.dumps(meta.get(\"content_keys\", []), indent=2)}')
else:
    print(f'Unknown type: {slide_type}')
    print('Run /ppt-list-slides to see all available types.')
"
```

After showing the list, explain that these types can be used with /ppt-build-slide or specified in the content JSON for /ppt-build.
