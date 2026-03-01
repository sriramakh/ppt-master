Browse and find PPT Master icons (144 PNG icons across 14 categories).

User input: $ARGUMENTS
If $ARGUMENTS is a category name or keyword, filter by it. Otherwise list all categories.

Run:

```bash
cd "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master" && \
source .venv/bin/activate && \
python -c "
import json, sys
from pathlib import Path

query = '$ARGUMENTS'.strip().lower()
manifest_path = Path('data/icons/manifest.json')
if not manifest_path.exists():
    print('Icon manifest not found. Run: pptmaster generate-icons')
    sys.exit(1)

manifest = json.loads(manifest_path.read_text())
icons = manifest.get('icons', [])

if query:
    matched = [i for i in icons if query in i['name'].lower() or query in i['category'].lower() or any(query in t.lower() for t in i.get('tags', []))]
    print(f'Icons matching \"{query}\": {len(matched)}')
    for i in matched[:30]:
        print(f'  {i[\"category\"]}/{i[\"name\"]}  ->  {i[\"path\"]}')
    if len(matched) > 30:
        print(f'  ... and {len(matched)-30} more')
else:
    cats = {}
    for i in icons:
        cats.setdefault(i['category'], []).append(i['name'])
    print(f'Total icons: {len(icons)} across {len(cats)} categories')
    print()
    for cat, names in sorted(cats.items()):
        print(f'{cat} ({len(names)}): {', '.join(names[:8])}{' ...' if len(names) > 8 else ''}')
"
```

Categories available: business, technology, finance, healthcare, education, sustainability, communication, people, analytics, logistics, security, creative, real-estate, plus any others generated.

To get the absolute path of a specific icon for use in presentations:
- Run: `get_icon_path(query)` via the PPT Master MCP server
- Or look in: `/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master/data/icons/{category}/{name}.png`

Icons are used automatically by the `icon_grid` slide type when building presentations.
