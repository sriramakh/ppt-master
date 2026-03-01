List all available PPT Master visual themes.

Run the following and display the output to the user:

```bash
cd "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master" && \
source .venv/bin/activate && \
python -c "
from pptmaster.builder.themes import THEME_MAP, get_theme
print(f'{'KEY':<16} {'INDUSTRY':<22} {'STYLE':<14} {'DARK':<6} {'PRIMARY':<10} {'ACCENT':<10} FONT')
print('-' * 90)
for key in sorted(THEME_MAP.keys()):
    t = get_theme(key)
    dark = 'yes' if t.ux_style.dark_mode else 'no'
    print(f'{key:<16} {t.industry:<22} {t.ux_style.name:<14} {dark:<6} {t.primary:<10} {t.accent:<10} {t.font}')
print()
print(f'{len(THEME_MAP)} themes available.')
print('Use any KEY with: /ppt-build, /ppt-build-template, or the ai-build CLI command.')
"
```

After showing the table, briefly describe the 3 specialty themes: academic (SCHOLARLY style), research (LABORATORY dark mode), and report (DASHBOARD dense layout).
