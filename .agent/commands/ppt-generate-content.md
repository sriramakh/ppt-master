Generate AI presentation content as JSON without building the PPTX file.

User input: $ARGUMENTS
Expected: topic and optionally theme, company name, audience, tone

This command generates the structured content JSON that drives a presentation build. Use it when you want to inspect, edit, or customize the AI-generated content before building.

Steps:
1. Parse $ARGUMENTS for: topic, theme_key (default: corporate), company_name (default: "Acme Corp"), audience (default: "executives"), tone (default: "professional").
2. Run:

```bash
cd "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master" && \
source .venv/bin/activate && \
python -c "
import json
from pptmaster.content.builder_content_gen import generate_presentation_content
content = generate_presentation_content(
    topic='TOPIC_HERE',
    company_name='COMPANY_HERE',
    theme_key='THEME_HERE',
    audience='AUDIENCE_HERE',
    tone='TONE_HERE',
)
print(json.dumps(content, indent=2))
" 2>&1 | head -200
```

3. Display the JSON content to the user.
4. Explain the key sections:
   - `selected_slides`: list of slide types chosen by AI (8-23 types)
   - `sections`: section names for divider slides
   - Per-slide content keys with generated text and data

5. Offer to: (a) build the presentation as-is with /ppt-build, (b) let the user edit specific fields and then build with build_from_content(), or (c) regenerate with different parameters.

The content JSON can be passed directly to `build_from_content()` via the MCP server or CLI.
