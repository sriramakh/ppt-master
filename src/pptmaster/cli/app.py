"""Typer CLI for PPT Master."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.table import Table

app = typer.Typer(
    name="pptmaster",
    help="AI-Powered Executive Presentation Generator",
    no_args_is_help=True,
)
console = Console()


@app.command()
def analyze(
    template: Path = typer.Option(..., "-t", "--template", help="Path to PPTX/POTX template"),
    toolkit: Optional[Path] = typer.Option(None, "--toolkit", help="Path to toolkit PPTX with icons"),
) -> None:
    """Analyze a template and display its design DNA."""
    from pptmaster.analyzer.template_analyzer import analyze as run_analyze

    with console.status("[bold green]Analyzing template..."):
        profile = run_analyze(template, toolkit_path=toolkit)

    # Display results
    console.print(f"\n[bold]Template:[/bold] {profile.template_name}")
    console.print(f"[bold]Slide Size:[/bold] {profile.slide_size.width_inches:.1f}\" x {profile.slide_size.height_inches:.1f}\"")

    # Colors
    console.print("\n[bold]Color Scheme:[/bold]")
    color_table = Table(show_header=True)
    color_table.add_column("Slot")
    color_table.add_column("Hex")
    color_table.add_column("Preview")
    for name in ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]:
        hex_val = getattr(profile.colors, name)
        color_table.add_row(name, hex_val, f"[on {hex_val}]      [/]")
    console.print(color_table)

    # Fonts
    console.print(f"\n[bold]Fonts:[/bold]")
    console.print(f"  Headlines: {profile.fonts.major}")
    console.print(f"  Body: {profile.fonts.minor}")

    # Layouts
    console.print(f"\n[bold]Layouts ({len(profile.layouts)} total):[/bold]")
    cats = profile.layout_categories
    layout_table = Table(show_header=True)
    layout_table.add_column("Category")
    layout_table.add_column("Count")
    layout_table.add_column("Layouts")
    for cat, layouts in sorted(cats.items()):
        names = ", ".join(l.name for l in layouts[:3])
        if len(layouts) > 3:
            names += f" (+{len(layouts) - 3})"
        layout_table.add_row(cat, str(len(layouts)), names)
    console.print(layout_table)

    # Icons
    if profile.icons:
        console.print(f"\n[bold]Icons:[/bold] {len(profile.icons)} from toolkit")
        icon_cats = set(i.category for i in profile.icons)
        console.print(f"  Categories: {len(icon_cats)}")

    console.print("\n[green]Analysis complete![/green]")


@app.command()
def generate(
    topic: str = typer.Option(..., "--topic", help="Presentation topic"),
    template: Path = typer.Option(..., "-t", "--template", help="Path to PPTX/POTX template"),
    output: Path = typer.Option("output.pptx", "-o", "--output", help="Output PPTX path"),
    num_slides: int = typer.Option(10, "-n", "--slides", help="Number of slides"),
    toolkit: Optional[Path] = typer.Option(None, "--toolkit", help="Path to toolkit PPTX with icons"),
    audience: str = typer.Option("", "--audience", help="Target audience"),
    api_key: Optional[str] = typer.Option(None, "--api-key", envvar="PPTMASTER_OPENAI_API_KEY"),
) -> None:
    """Generate a presentation from a topic."""
    from pptmaster.analyzer.template_analyzer import analyze as run_analyze
    from pptmaster.content.content_engine import generate as run_generate
    from pptmaster.composer.slide_composer import compose

    with console.status("[bold green]Analyzing template..."):
        profile = run_analyze(template, toolkit_path=toolkit)

    console.print(f"[green]✓[/green] Template analyzed: {len(profile.layouts)} layouts, {len(profile.icons)} icons")

    with console.status("[bold blue]Generating content with GPT-4..."):
        outline, assignments = run_generate(
            profile=profile,
            topic=topic,
            num_slides=num_slides,
            audience=audience,
            api_key=api_key,
        )

    console.print(f"[green]✓[/green] Generated {len(outline.slides)} slides")

    # Show slide overview
    for i, slide in enumerate(outline.slides):
        console.print(f"  {i+1}. [{slide.slide_type.value}] {slide.title}")

    with console.status("[bold magenta]Composing PPTX..."):
        result_path = compose(outline, profile, output, assignments)

    console.print(f"\n[bold green]✓ Presentation saved to {result_path}[/bold green]")


@app.command("from-doc")
def from_doc(
    input_file: Path = typer.Option(..., "-i", "--input", help="Input document (PDF, DOCX, TXT)"),
    template: Path = typer.Option(..., "-t", "--template", help="Path to PPTX/POTX template"),
    output: Path = typer.Option("output.pptx", "-o", "--output", help="Output PPTX path"),
    num_slides: int = typer.Option(10, "-n", "--slides", help="Number of slides"),
    toolkit: Optional[Path] = typer.Option(None, "--toolkit", help="Path to toolkit PPTX with icons"),
    api_key: Optional[str] = typer.Option(None, "--api-key", envvar="PPTMASTER_OPENAI_API_KEY"),
) -> None:
    """Generate a presentation from a document (PDF, DOCX, TXT)."""
    from pptmaster.analyzer.template_analyzer import analyze as run_analyze
    from pptmaster.content.content_engine import generate as run_generate
    from pptmaster.composer.slide_composer import compose

    with console.status("[bold green]Analyzing template..."):
        profile = run_analyze(template, toolkit_path=toolkit)

    with console.status("[bold blue]Processing document and generating content..."):
        outline, assignments = run_generate(
            profile=profile,
            input_source=input_file,
            num_slides=num_slides,
            api_key=api_key,
        )

    console.print(f"[green]✓[/green] Generated {len(outline.slides)} slides from {input_file.name}")

    with console.status("[bold magenta]Composing PPTX..."):
        result_path = compose(outline, profile, output, assignments)

    console.print(f"\n[bold green]✓ Presentation saved to {result_path}[/bold green]")


@app.command("from-data")
def from_data(
    input_file: Path = typer.Option(..., "-i", "--input", help="Input data file (CSV, XLSX)"),
    template: Path = typer.Option(..., "-t", "--template", help="Path to PPTX/POTX template"),
    output: Path = typer.Option("output.pptx", "-o", "--output", help="Output PPTX path"),
    num_slides: int = typer.Option(10, "-n", "--slides", help="Number of slides"),
    toolkit: Optional[Path] = typer.Option(None, "--toolkit", help="Path to toolkit PPTX with icons"),
    api_key: Optional[str] = typer.Option(None, "--api-key", envvar="PPTMASTER_OPENAI_API_KEY"),
) -> None:
    """Generate a presentation from data (CSV, Excel)."""
    from pptmaster.analyzer.template_analyzer import analyze as run_analyze
    from pptmaster.content.content_engine import generate as run_generate
    from pptmaster.composer.slide_composer import compose

    with console.status("[bold green]Analyzing template..."):
        profile = run_analyze(template, toolkit_path=toolkit)

    with console.status("[bold blue]Processing data and generating content..."):
        outline, assignments = run_generate(
            profile=profile,
            input_source=input_file,
            num_slides=num_slides,
            style_notes="Focus on data visualizations — use charts and tables where possible.",
            api_key=api_key,
        )

    console.print(f"[green]✓[/green] Generated {len(outline.slides)} slides from {input_file.name}")

    with console.status("[bold magenta]Composing PPTX..."):
        result_path = compose(outline, profile, output, assignments)

    console.print(f"\n[bold green]✓ Presentation saved to {result_path}[/bold green]")


@app.command("from-url")
def from_url(
    url: str = typer.Option(..., "--url", help="URL to extract content from"),
    template: Path = typer.Option(..., "-t", "--template", help="Path to PPTX/POTX template"),
    output: Path = typer.Option("output.pptx", "-o", "--output", help="Output PPTX path"),
    num_slides: int = typer.Option(10, "-n", "--slides", help="Number of slides"),
    toolkit: Optional[Path] = typer.Option(None, "--toolkit", help="Path to toolkit PPTX with icons"),
    api_key: Optional[str] = typer.Option(None, "--api-key", envvar="PPTMASTER_OPENAI_API_KEY"),
) -> None:
    """Generate a presentation from a web page URL."""
    from pptmaster.analyzer.template_analyzer import analyze as run_analyze
    from pptmaster.content.content_engine import generate as run_generate
    from pptmaster.composer.slide_composer import compose

    with console.status("[bold green]Analyzing template..."):
        profile = run_analyze(template, toolkit_path=toolkit)

    with console.status("[bold blue]Fetching URL and generating content..."):
        outline, assignments = run_generate(
            profile=profile,
            input_source=url,
            num_slides=num_slides,
            api_key=api_key,
        )

    console.print(f"[green]✓[/green] Generated {len(outline.slides)} slides from URL")

    with console.status("[bold magenta]Composing PPTX..."):
        result_path = compose(outline, profile, output, assignments)

    console.print(f"\n[bold green]✓ Presentation saved to {result_path}[/bold green]")


@app.command("build-template")
def build_template_cmd(
    output: Path = typer.Option("corporate_template_30.pptx", "-o", "--output", help="Output PPTX path"),
    company: str = typer.Option("[Company Name]", "--company", help="Company name for placeholders"),
    theme_name: str = typer.Option("corporate", "--theme", help="Theme name (corporate, healthcare, technology, finance, education, sustainability, luxury, startup, government, realestate, creative)"),
) -> None:
    """Build a 40-slide Fortune 500 corporate template from scratch."""
    from pptmaster.builder.template_builder import build_template
    from pptmaster.builder.themes import get_theme

    theme = get_theme(theme_name)

    with console.status(f"[bold magenta]Building 40-slide {theme.industry} template..."):
        result_path = build_template(output, company, theme=theme)

    console.print(f"\n[bold green]✓ Template saved to {result_path}[/bold green]")
    console.print(f"  40 slides | {theme.industry} | {theme.company_name}")


@app.command("ai-build")
def ai_build(
    topic: str = typer.Option(..., "--topic", help="Presentation topic"),
    company: str = typer.Option("Acme Corp", "--company", help="Company name"),
    industry: str = typer.Option("", "--industry", help="Industry context"),
    audience: str = typer.Option("", "--audience", help="Target audience"),
    theme: str = typer.Option("corporate", "--theme", help="Theme key (corporate, healthcare, technology, finance, education, sustainability, luxury, startup, government, realestate, creative)"),
    output: Path = typer.Option("output.pptx", "-o", "--output", help="Output PPTX path"),
    context: str = typer.Option("", "--context", help="Additional context or instructions for the AI"),
    provider: str = typer.Option("minimax", "--provider", help="LLM provider (minimax, openai)"),
) -> None:
    """AI-generate a 40-slide presentation from a topic."""
    from pptmaster.builder.ai_builder import ai_build_presentation

    console.print(f"[bold blue]Topic:[/bold blue] {topic}")
    console.print(f"[bold blue]Company:[/bold blue] {company}")
    console.print(f"[bold blue]Theme:[/bold blue] {theme}")
    if audience:
        console.print(f"[bold blue]Audience:[/bold blue] {audience}")

    with console.status("[bold green]Generating content with AI and building presentation..."):
        result_path = ai_build_presentation(
            topic=topic,
            company_name=company,
            industry=industry,
            audience=audience,
            theme_key=theme,
            output_path=output,
            additional_context=context,
            provider=provider,
        )

    console.print(f"\n[bold green]✓ AI presentation saved to {result_path}[/bold green]")
    console.print("  40 slides | AI-generated content | Open in PowerPoint to review")


@app.command("list-themes")
def list_themes_cmd() -> None:
    """List all available template themes with their visual styles."""
    from pptmaster.builder.ai_builder import list_available_themes

    themes = list_available_themes()
    table = Table(title="Available Themes", show_header=True, header_style="bold cyan")
    table.add_column("Key", style="bold")
    table.add_column("Industry")
    table.add_column("UX Style")
    table.add_column("Primary")
    table.add_column("Accent")
    table.add_column("Font")

    for t in themes:
        table.add_row(
            t["key"], t["industry"], t["ux_style"],
            f'[on {t["primary_color"]}]  [/] {t["primary_color"]}',
            f'[on {t["accent_color"]}]  [/] {t["accent_color"]}',
            t["font"],
        )

    console.print(table)


@app.command("build-all-templates")
def build_all_templates_cmd(
    output_dir: Path = typer.Option(".", "-o", "--output-dir", help="Output directory for all templates"),
) -> None:
    """Build all 11 industry-themed templates (corporate + 10 industries)."""
    from pptmaster.builder.template_builder import build_all_templates

    with console.status("[bold magenta]Building 11 themed templates..."):
        paths = build_all_templates(output_dir)

    console.print(f"\n[bold green]✓ Built {len(paths)} templates:[/bold green]")
    for p in paths:
        console.print(f"  {p}")


@app.command("generate-icons")
def generate_icons_cmd(
    output_dir: str = typer.Option("data/icons", "--output-dir", "-o", help="Output directory for icons"),
    concurrency: int = typer.Option(5, "--concurrency", "-c", help="Max concurrent API calls"),
) -> None:
    """Generate ~150 professional PNG icons using GPT image API.

    Requires OPENAI_API_KEY environment variable.
    Cost: ~$1.65 for 150 icons at 'low' quality.
    """
    from pptmaster.assets.icon_generator import generate_all_icons
    generate_all_icons(output_dir=output_dir, concurrency=concurrency)


@app.command("build-icon-template")
def build_icon_template_cmd(
    output: str = typer.Option("icon_toolkit.pptx", "--output", "-o", help="Output PPTX path"),
    icon_dir: str = typer.Option(None, "--icon-dir", help="Custom icon directory"),
) -> None:
    """Build a reference PPTX showing all available icons.

    Creates one slide per category with icons in a grid layout.
    """
    from pptmaster.builder.icon_template_builder import build_icon_template
    path = build_icon_template(output_path=output, icon_dir=icon_dir)
    typer.echo(f"Icon template saved: {path}")


@app.command("chat")
def chat_cmd(
    topic: Optional[str] = typer.Option(None, "--topic", help="Presentation topic"),
    company: str = typer.Option("Acme Corp", "--company", help="Company name"),
    industry: str = typer.Option("", "--industry", help="Industry context"),
    audience: str = typer.Option("", "--audience", help="Target audience"),
    theme: str = typer.Option("corporate", "--theme", help="Theme key"),
    output: Path = typer.Option("output.pptx", "-o", "--output", help="Output PPTX path"),
    context: str = typer.Option("", "--context", help="Additional context for initial generation"),
    provider: str = typer.Option("minimax", "--provider", help="LLM provider (minimax, openai, or any OpenAI-compatible provider)"),
    chat_model: Optional[str] = typer.Option(None, "--chat-model", help="Override the model used in the chat loop (e.g. gpt-4o, MiniMax-M2.5)"),
    load: Optional[Path] = typer.Option(None, "--load", help="Load a previously saved content JSON to continue editing"),
) -> None:
    """Build a presentation then refine it interactively through conversation.

    Generates an initial presentation (same as ai-build), then enters a chat loop
    where you can describe changes in plain English. The LLM uses tools to add/remove/
    reorder slides, regenerate content, switch themes, and more. The PPTX is rebuilt
    after each turn.

    The same --provider is used for both initial generation and the chat loop.
    Use --chat-model to override the model for the chat session only.

    Requires the appropriate API key in your .env file:
      PPTMASTER_OPENAI_API_KEY   (for openai)
      PPTMASTER_MINIMAX_API_KEY  (for minimax)
    """
    import json as _json

    from pptmaster.chat.loop import run_chat_loop
    from pptmaster.chat.session import PresentationSession

    # ── Load or generate initial content ─────────────────────────────
    if load:
        if not load.exists():
            console.print(f"[red]File not found:[/red] {load}")
            raise typer.Exit(1)
        content_dict = _json.loads(load.read_text())
        gen_result = content_dict
        resolved_topic = topic or content_dict.get("content", {}).get("cover_title", "Presentation")
        console.print(f"[green]✓[/green] Loaded content from [cyan]{load}[/cyan]")
    else:
        if not topic:
            console.print("[red]Error:[/red] --topic is required (or use --load to edit an existing presentation)")
            raise typer.Exit(1)

        from pptmaster.content.builder_content_gen import generate_builder_content

        console.print("[bold blue]Generating initial presentation...[/bold blue]")
        console.print(f"  Topic:    {topic}")
        console.print(f"  Company:  {company}")
        console.print(f"  Theme:    {theme}")
        console.print(f"  Provider: {provider}")

        with console.status("[bold green]Generating content with AI..."):
            gen_result = generate_builder_content(
                topic=topic,
                company_name=company,
                industry=industry,
                audience=audience,
                additional_context=context,
                provider=provider,
            )

        resolved_topic = topic
        console.print(
            f"  [green]✓[/green] AI selected {len(gen_result['selected_slides'])} slides "
            f"in {len(gen_result['sections'])} sections"
        )

    # ── Create session and build initial PPTX ────────────────────────
    session = PresentationSession(
        gen_result=gen_result,
        theme_key=theme,
        company_name=company,
        topic=resolved_topic,
        output_path=output,
        provider=provider,
    )

    with console.status("[bold magenta]Building initial PPTX..."):
        result_path = session.rebuild()
    console.print(f"  [green]✓[/green] Saved to [cyan]{result_path}[/cyan]")

    # ── Enter chat loop ───────────────────────────────────────────────
    run_chat_loop(session, provider=provider, model=chat_model)


@app.command("from-content")
def from_content_cmd(
    input_file: str = typer.Option(..., "--input", "-i", help="JSON file with content dict"),
    output: str = typer.Option("output.pptx", "--output", "-o", help="Output PPTX path"),
    company: str = typer.Option("Acme Corp", "--company", help="Company name"),
    theme: str = typer.Option("corporate", "--theme", help="Theme key"),
) -> None:
    """Build a presentation from a pre-made content JSON file.

    The JSON should contain: selected_slides, sections (optional), and content keys.
    No LLM call needed — pass your own or AI-generated content directly.
    """
    import json
    content_dict = json.loads(Path(input_file).read_text())
    from pptmaster.builder.ai_builder import build_from_content
    path = build_from_content(
        content_dict=content_dict,
        company_name=company,
        theme_key=theme,
        output_path=output,
    )
    console.print(f"[green]Presentation saved:[/green] {path}")


if __name__ == "__main__":
    app()
