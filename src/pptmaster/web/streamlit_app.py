"""Streamlit web interface for PPT Master."""

from __future__ import annotations

import io
import json
import tempfile
from pathlib import Path

import streamlit as st


def main() -> None:
    st.set_page_config(
        page_title="PPT Master",
        page_icon="📊",
        layout="wide",
    )

    st.title("PPT Master")
    st.caption("AI-Powered Executive Presentation Generator")

    # ── Sidebar: Template Upload & Design DNA ──────────────────────────

    with st.sidebar:
        st.header("Template")

        template_file = st.file_uploader(
            "Upload PPTX/POTX template",
            type=["pptx", "potx"],
            key="template",
        )

        toolkit_file = st.file_uploader(
            "Upload Icons toolkit (optional)",
            type=["pptx"],
            key="toolkit",
        )

        api_key = st.text_input(
            "OpenAI API Key",
            type="password",
            value=st.session_state.get("api_key", ""),
            key="api_key_input",
        )
        if api_key:
            st.session_state["api_key"] = api_key

        # Show design DNA if template is loaded
        if "profile" in st.session_state:
            _show_design_dna(st.session_state["profile"])

    # ── Analyze template on upload ─────────────────────────────────────

    if template_file and "profile" not in st.session_state:
        _analyze_template(template_file, toolkit_file)
    elif template_file is None:
        st.session_state.pop("profile", None)
        st.info("Upload a template to get started.")
        return

    if "profile" not in st.session_state:
        return

    # ── Main: Tabbed Interface ─────────────────────────────────────────

    tab_gen, tab_preview, tab_history = st.tabs(["Generate", "Preview & Edit", "History"])

    with tab_gen:
        _generation_tab()

    with tab_preview:
        _preview_tab()

    with tab_history:
        _history_tab()


def _analyze_template(template_file, toolkit_file) -> None:
    """Analyze uploaded template and store profile in session state."""
    from pptmaster.analyzer.template_analyzer import analyze

    with st.spinner("Analyzing template..."):
        # Save uploaded files to temp paths
        with tempfile.NamedTemporaryFile(suffix=f".{template_file.name.split('.')[-1]}", delete=False) as tmp:
            tmp.write(template_file.getvalue())
            template_path = tmp.name

        toolkit_path = None
        if toolkit_file:
            with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
                tmp.write(toolkit_file.getvalue())
                toolkit_path = tmp.name

        try:
            profile = analyze(template_path, toolkit_path=toolkit_path, use_cache=False)
            st.session_state["profile"] = profile
            st.session_state["template_path"] = template_path
            st.success(f"Template analyzed: {len(profile.layouts)} layouts, {len(profile.icons)} icons")
        except Exception as e:
            st.error(f"Failed to analyze template: {e}")


def _show_design_dna(profile) -> None:
    """Display design DNA in the sidebar."""
    st.subheader("Design DNA")

    # Colors
    st.write("**Colors**")
    cols = st.columns(6)
    accent_names = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
    for i, name in enumerate(accent_names):
        color = getattr(profile.colors, name)
        cols[i].color_picker(name, color, key=f"color_{name}", disabled=True)

    # Fonts
    st.write(f"**Headlines:** {profile.fonts.major}")
    st.write(f"**Body:** {profile.fonts.minor}")

    # Layout counts
    st.write("**Layouts**")
    cats = profile.layout_categories
    for cat, layouts in sorted(cats.items()):
        st.write(f"- {cat}: {len(layouts)}")

    if profile.icons:
        st.write(f"**Icons:** {len(profile.icons)}")


def _generation_tab() -> None:
    """The main generation interface."""
    profile = st.session_state["profile"]

    # Input mode selection
    input_mode = st.radio(
        "Input mode",
        ["Text prompt", "Document upload", "Data file", "URL"],
        horizontal=True,
    )

    topic = ""
    input_source = None

    if input_mode == "Text prompt":
        topic = st.text_area(
            "Describe your presentation",
            placeholder="e.g., Global Biodiversity Trends 2025 — a data-driven overview for conservation stakeholders",
            height=120,
        )

    elif input_mode == "Document upload":
        doc_file = st.file_uploader(
            "Upload document",
            type=["pdf", "docx", "txt", "md"],
            key="doc_upload",
        )
        if doc_file:
            with tempfile.NamedTemporaryFile(suffix=f".{doc_file.name.split('.')[-1]}", delete=False) as tmp:
                tmp.write(doc_file.getvalue())
                input_source = tmp.name
        topic = st.text_input("Topic (optional)", key="doc_topic")

    elif input_mode == "Data file":
        data_file = st.file_uploader(
            "Upload data",
            type=["csv", "xlsx", "xls"],
            key="data_upload",
        )
        if data_file:
            with tempfile.NamedTemporaryFile(suffix=f".{data_file.name.split('.')[-1]}", delete=False) as tmp:
                tmp.write(data_file.getvalue())
                input_source = tmp.name
        topic = st.text_input("Topic (optional)", key="data_topic")

    elif input_mode == "URL":
        url = st.text_input("Enter URL", placeholder="https://example.com/article")
        if url:
            input_source = url
        topic = st.text_input("Topic (optional)", key="url_topic")

    # Options
    col1, col2 = st.columns(2)
    with col1:
        num_slides = st.slider("Number of slides", 5, 25, 10)
    with col2:
        audience = st.text_input("Target audience (optional)")

    # Generate button
    if st.button("Generate Presentation", type="primary", use_container_width=True):
        api_key = st.session_state.get("api_key", "")
        if not api_key:
            st.error("Please enter your OpenAI API key in the sidebar.")
            return

        if not topic and not input_source:
            st.error("Please provide a topic or input source.")
            return

        _run_generation(profile, topic, input_source, num_slides, audience, api_key)


def _run_generation(profile, topic, input_source, num_slides, audience, api_key) -> None:
    """Run the full generation pipeline."""
    from pptmaster.content.content_engine import generate
    from pptmaster.composer.slide_composer import compose

    # Step 1: Generate content
    with st.spinner("Generating content with GPT-4..."):
        try:
            outline, assignments = generate(
                profile=profile,
                topic=topic,
                input_source=input_source,
                num_slides=num_slides,
                audience=audience,
                api_key=api_key,
            )
        except Exception as e:
            st.error(f"Content generation failed: {e}")
            return

    st.success(f"Generated {len(outline.slides)} slides")
    st.session_state["outline"] = outline
    st.session_state["assignments"] = assignments

    # Step 2: Compose PPTX
    with st.spinner("Composing PPTX..."):
        try:
            with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
                output_path = tmp.name
            compose(outline, profile, output_path, assignments)
            st.session_state["output_path"] = output_path

            # Read the file for download
            with open(output_path, "rb") as f:
                st.session_state["output_bytes"] = f.read()
        except Exception as e:
            st.error(f"Composition failed: {e}")
            return

    st.success("Presentation ready!")

    # Download button
    st.download_button(
        label="Download PPTX",
        data=st.session_state["output_bytes"],
        file_name=f"{outline.title[:50].replace(' ', '_')}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True,
    )

    # Save to history
    if "history" not in st.session_state:
        st.session_state["history"] = []
    st.session_state["history"].append({
        "title": outline.title,
        "slides": len(outline.slides),
        "outline": outline,
    })


def _preview_tab() -> None:
    """Slide-by-slide preview and per-slide regeneration."""
    if "outline" not in st.session_state:
        st.info("Generate a presentation first to see the preview.")
        return

    outline = st.session_state["outline"]

    st.subheader(outline.title)
    if outline.subtitle:
        st.caption(outline.subtitle)

    for i, slide in enumerate(outline.slides):
        with st.expander(f"Slide {i+1}: {slide.title}", expanded=(i == 0)):
            col1, col2 = st.columns([3, 1])

            with col1:
                st.write(f"**Type:** {slide.slide_type.value}")
                st.write(f"**Title:** {slide.title}")

                if slide.body:
                    st.write(f"**Body:** {slide.body}")

                if slide.bullets:
                    st.write("**Bullets:**")
                    for bullet in slide.bullets:
                        st.write(f"  - {bullet}")

                if slide.columns:
                    st.write("**Columns:**")
                    cols = st.columns(len(slide.columns))
                    for j, col in enumerate(slide.columns):
                        with cols[j]:
                            if col.heading:
                                st.write(f"**{col.heading}**")
                            if col.body:
                                st.write(col.body)
                            for b in col.bullets:
                                st.write(f"  - {b}")

                if slide.metrics:
                    st.write("**Metrics:**")
                    mcols = st.columns(len(slide.metrics))
                    for j, metric in enumerate(slide.metrics):
                        with mcols[j]:
                            st.metric(metric.label, metric.value)

                if slide.chart:
                    st.write(f"**Chart:** {slide.chart.chart_type} — {slide.chart.title}")

                if slide.table:
                    import pandas as pd
                    if slide.table.headers and slide.table.rows:
                        df = pd.DataFrame(slide.table.rows, columns=slide.table.headers)
                        st.dataframe(df, use_container_width=True)

            with col2:
                if slide.speaker_notes:
                    st.write("**Notes:**")
                    st.caption(slide.speaker_notes)

    # Download button (repeated for convenience)
    if "output_bytes" in st.session_state:
        st.download_button(
            label="Download PPTX",
            data=st.session_state["output_bytes"],
            file_name=f"{outline.title[:50].replace(' ', '_')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )


def _history_tab() -> None:
    """Show generation history for the session."""
    if "history" not in st.session_state or not st.session_state["history"]:
        st.info("No presentations generated yet.")
        return

    for i, entry in enumerate(reversed(st.session_state["history"])):
        st.write(f"**{i+1}. {entry['title']}** — {entry['slides']} slides")


if __name__ == "__main__":
    main()
