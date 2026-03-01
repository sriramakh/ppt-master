Build a full AI-generated PPTX presentation using PPT Master.

The user wants to build a presentation. Their request: $ARGUMENTS

Steps:
1. Parse the request for: topic, theme (default: corporate), company name (default: "Acme Corp"), output path (default: "output.pptx").
2. If no topic is clear from "$ARGUMENTS", ask the user what the presentation is about before proceeding.
3. Suggest an appropriate theme if one wasn't specified. Available themes: corporate, healthcare, technology, finance, education, sustainability, luxury, startup, government, realestate, creative, academic, research, report.
4. Run this Python snippet via Bash to build the presentation:

```bash
cd "/Volumes/Seagate Expansion Drive/Seagate AI/PPT Master" && \
source .venv/bin/activate && \
python -c "
from pptmaster.builder.ai_builder import ai_build_presentation
path = ai_build_presentation(
    topic='TOPIC_HERE',
    company_name='COMPANY_HERE',
    theme_key='THEME_HERE',
    output_path='OUTPUT_HERE'
)
print(f'Presentation saved to: {path}')
"
```

5. After building, report the output file path and offer a summary of which slides were included.

PPT Master builds 12–25 slides selected by AI from 32 slide types. Always included: cover, table of contents, section dividers, and thank you slide. The AI selects content-appropriate slides and generates all text/data automatically.
