"""Tool definitions and executor for the presentation chat loop.

Tools are defined in the OpenAI function-calling format, which is supported
by any OpenAI-compatible provider (OpenAI, MiniMax, Groq, Ollama, etc.).
"""

from __future__ import annotations

from typing import Any

from pptmaster.chat.session import PresentationSession
from pptmaster.content.builder_content_gen import SLIDE_CATALOG

# ── Tool definitions (OpenAI function-calling format) ────────────────

TOOLS: list[dict[str, Any]] = [
    {
        "type": "function",
        "function": {
            "name": "get_state",
            "description": (
                "Get the current state of the presentation: topic, company, theme, "
                "sections, and the list of content slides in each section."
            ),
            "parameters": {
                "type": "object",
                "properties": {},
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_slide_types",
            "description": (
                "List all 32 available slide types that can be added to the presentation, "
                "with a short description of each."
            ),
            "parameters": {
                "type": "object",
                "properties": {},
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "add_slide",
            "description": (
                "Add a new slide type to the presentation. "
                "Content is automatically generated via LLM. "
                "Use 'instruction' to guide what the content should focus on."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_type": {
                        "type": "string",
                        "description": (
                            "The slide type key, e.g. 'bar_chart', 'team_leadership', "
                            "'swot_matrix'. Use list_slide_types to see all options."
                        ),
                    },
                    "after": {
                        "type": "string",
                        "description": (
                            "Insert the new slide immediately after this existing slide type. "
                            "Omit to append at the end."
                        ),
                    },
                    "section_title": {
                        "type": "string",
                        "description": (
                            "Add to this section (matched by title). "
                            "Omit to place automatically."
                        ),
                    },
                    "instruction": {
                        "type": "string",
                        "description": (
                            "Optional content instruction, "
                            "e.g. 'focus on cost savings' or 'use Q3 data'."
                        ),
                    },
                },
                "required": ["slide_type"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "remove_slide",
            "description": "Remove a slide type from the presentation.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_type": {
                        "type": "string",
                        "description": "The slide type key to remove.",
                    },
                },
                "required": ["slide_type"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "move_slide",
            "description": "Move a slide to a different position in the presentation.",
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_type": {
                        "type": "string",
                        "description": "The slide type to move.",
                    },
                    "after": {
                        "type": "string",
                        "description": (
                            "Place after this existing slide type. "
                            "Omit to move to the very beginning."
                        ),
                    },
                },
                "required": ["slide_type"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "regenerate_slide",
            "description": (
                "Regenerate the content of a specific slide using the LLM, "
                "optionally with new instructions. Use this to improve or redirect "
                "the content of an existing slide."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_type": {
                        "type": "string",
                        "description": "The slide type to regenerate content for.",
                    },
                    "instruction": {
                        "type": "string",
                        "description": (
                            "Guidance for the regeneration, e.g. "
                            "'make it more optimistic', 'focus on Q3 results'."
                        ),
                    },
                },
                "required": ["slide_type"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "update_content",
            "description": (
                "Directly update specific content fields for a slide without calling the LLM. "
                "The 'updates' object maps content key names to their new values."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "slide_type": {
                        "type": "string",
                        "description": "The slide type whose content fields to update.",
                    },
                    "updates": {
                        "type": "object",
                        "description": (
                            "Key-value pairs where keys are content field names "
                            "(e.g. 'exec_title', 'swot') and values are their new content."
                        ),
                    },
                },
                "required": ["slide_type", "updates"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "set_theme",
            "description": "Change the visual theme of the entire presentation.",
            "parameters": {
                "type": "object",
                "properties": {
                    "theme_key": {
                        "type": "string",
                        "description": (
                            "Theme key. One of: corporate, healthcare, technology, "
                            "finance, education, sustainability, luxury, startup, "
                            "government, realestate, creative, academic, research, report"
                        ),
                    },
                },
                "required": ["theme_key"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "rename_section",
            "description": "Rename a section in the presentation.",
            "parameters": {
                "type": "object",
                "properties": {
                    "old_title": {
                        "type": "string",
                        "description": "Current section title (case-insensitive match).",
                    },
                    "new_title": {
                        "type": "string",
                        "description": "New section title.",
                    },
                },
                "required": ["old_title", "new_title"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "add_section",
            "description": (
                "Add a new empty section to the presentation. "
                "You can then use add_slide with section_title to populate it."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "title": {
                        "type": "string",
                        "description": "Section title.",
                    },
                    "subtitle": {
                        "type": "string",
                        "description": "Optional section subtitle.",
                    },
                },
                "required": ["title"],
            },
        },
    },
]


# ── Executor ─────────────────────────────────────────────────────────


def execute_tool(
    session: PresentationSession,
    tool_name: str,
    tool_input: dict[str, Any],
) -> str:
    """Execute a tool call against the session. Returns result as a plain string."""

    if tool_name == "get_state":
        return session.get_state()

    elif tool_name == "list_slide_types":
        lines = ["Available slide types:\n"]
        for key, entry in SLIDE_CATALOG.items():
            lines.append(f"  {key:<25s} — {entry['description']}")
        return "\n".join(lines)

    elif tool_name == "add_slide":
        return session.add_slide(
            slide_type=tool_input["slide_type"],
            after=tool_input.get("after"),
            section_title=tool_input.get("section_title"),
            instruction=tool_input.get("instruction", ""),
        )

    elif tool_name == "remove_slide":
        return session.remove_slide(tool_input["slide_type"])

    elif tool_name == "move_slide":
        return session.move_slide(
            slide_type=tool_input["slide_type"],
            after=tool_input.get("after"),
        )

    elif tool_name == "regenerate_slide":
        return session.regenerate_slide(
            slide_type=tool_input["slide_type"],
            instruction=tool_input.get("instruction", ""),
        )

    elif tool_name == "update_content":
        return session.update_content(
            slide_type=tool_input["slide_type"],
            updates=tool_input["updates"],
        )

    elif tool_name == "set_theme":
        return session.set_theme(tool_input["theme_key"])

    elif tool_name == "rename_section":
        return session.rename_section(tool_input["old_title"], tool_input["new_title"])

    elif tool_name == "add_section":
        return session.add_section(
            title=tool_input["title"],
            subtitle=tool_input.get("subtitle", ""),
        )

    else:
        return f"Unknown tool: {tool_name}"
