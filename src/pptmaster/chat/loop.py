"""Conversational presentation editing loop.

Uses the OpenAI SDK (same as LLMClient) so it works with any OpenAI-compatible
provider: OpenAI, MiniMax, Groq, Ollama, a local model behind LiteLLM, etc.

Provider resolution mirrors LLMClient — pass the same --provider flag as ai-build.
"""

from __future__ import annotations

import json
import re
import sys
from typing import Any

from openai import APIConnectionError, APIStatusError, AuthenticationError, OpenAI
from rich.console import Console
from rich.panel import Panel

from pptmaster.chat.session import PresentationSession
from pptmaster.chat.tools import TOOLS, execute_tool

console = Console()

_SYSTEM_PROMPT = """\
You are an expert presentation editor for PPT Master. You help users refine their \
AI-generated presentations through natural conversation.

You have access to tools to:
- Inspect the current state (get_state)
- List all 32 available slide types (list_slide_types)
- Add slides with AI-generated content (add_slide)
- Remove slides (remove_slide)
- Reorder slides (move_slide)
- Regenerate slide content with new instructions (regenerate_slide)
- Directly update specific content fields (update_content)
- Change the visual theme (set_theme)
- Rename or add sections (rename_section, add_section)

Guidelines:
- Act on user requests immediately using the appropriate tool(s).
- Call get_state first if you need context about the current structure.
- For vague requests like "improve the intro" use regenerate_slide with a clear instruction.
- Chain multiple tool calls in a single turn when making several changes.
- After making changes, briefly describe what was done and suggest what could come next.
- The PPTX is automatically rebuilt after each turn that makes changes.
"""


# ── Provider resolution ───────────────────────────────────────────────


def _make_client(
    provider: str,
    api_key: str | None = None,
    model: str | None = None,
) -> tuple[OpenAI, str]:
    """Create an OpenAI-compatible client and resolve the model name.

    Mirrors the resolution logic in LLMClient so the same --provider flag works
    for both content generation and the chat loop.
    """
    from pptmaster.config import get_settings

    # Import provider presets from LLMClient to stay in sync
    from pptmaster.content.llm_client import _PROVIDER_PRESETS

    settings = get_settings()
    preset_url, preset_model = _PROVIDER_PRESETS.get(provider, (None, "gpt-4o"))

    if provider == "minimax":
        resolved_key = api_key or settings.minimax_api_key or settings.openai_api_key
        resolved_model = model or settings.minimax_model or preset_model
        resolved_url = settings.minimax_base_url or preset_url
    elif provider == "openai":
        resolved_key = api_key or settings.openai_api_key
        resolved_model = model or settings.model or preset_model
        resolved_url = None
    else:
        # Custom provider — caller must supply api_key and model
        resolved_key = api_key or settings.openai_api_key
        resolved_model = model or preset_model or "gpt-4o"
        resolved_url = preset_url

    client = OpenAI(api_key=resolved_key, base_url=resolved_url)
    return client, resolved_model


# ── Response helpers ───────────────────────────────────────────────────


def _strip_thinking(text: str) -> str:
    """Remove <think>…</think> blocks injected by some providers (e.g. MiniMax)."""
    return re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()


def _call_with_tools(
    client: OpenAI,
    model: str,
    messages: list[dict[str, Any]],
) -> tuple[str, list[dict[str, Any]]]:
    """Call the LLM with tools. Returns (text_response, tool_calls_list).

    Tries streaming first for a better UX (text appears token-by-token).
    Falls back to a single blocking call if the provider doesn't support streaming
    with tool calls.

    tool_calls_list items have shape:
        {"id": str, "type": "function", "function": {"name": str, "arguments": str}}
    """
    text_parts: list[str] = []
    tool_calls_dict: dict[int, dict[str, Any]] = {}
    has_printed_prefix = False

    try:
        stream = client.chat.completions.create(
            model=model,
            messages=messages,
            tools=TOOLS,
            tool_choice="auto",
            stream=True,
        )

        for chunk in stream:
            if not chunk.choices:
                continue
            delta = chunk.choices[0].delta

            if delta.content:
                clean = _strip_thinking(delta.content)
                if clean:
                    if not has_printed_prefix:
                        console.print("\n[bold green]Assistant:[/bold green] ", end="")
                        has_printed_prefix = True
                    sys.stdout.write(clean)
                    sys.stdout.flush()
                    text_parts.append(clean)

            if delta.tool_calls:
                for tc in delta.tool_calls:
                    idx = tc.index
                    if idx not in tool_calls_dict:
                        tool_calls_dict[idx] = {
                            "id": tc.id or "",
                            "type": "function",
                            "function": {
                                "name": (tc.function.name or "") if tc.function else "",
                                "arguments": (tc.function.arguments or "") if tc.function else "",
                            },
                        }
                    else:
                        if tc.id:
                            tool_calls_dict[idx]["id"] = tc.id
                        if tc.function:
                            if tc.function.name:
                                tool_calls_dict[idx]["function"]["name"] += tc.function.name
                            if tc.function.arguments:
                                tool_calls_dict[idx]["function"]["arguments"] += tc.function.arguments

        if has_printed_prefix:
            sys.stdout.write("\n")
            sys.stdout.flush()

    except Exception:
        # Provider doesn't support streaming with tools — fall back to blocking call
        with console.status("[dim]Thinking...[/dim]"):
            resp = client.chat.completions.create(
                model=model,
                messages=messages,
                tools=TOOLS,
                tool_choice="auto",
            )
        msg = resp.choices[0].message
        if msg.content:
            text = _strip_thinking(msg.content)
            if text:
                console.print(f"\n[bold green]Assistant:[/bold green] {text}")
                text_parts.append(text)
        if msg.tool_calls:
            for i, tc in enumerate(msg.tool_calls):
                tool_calls_dict[i] = {
                    "id": tc.id,
                    "type": "function",
                    "function": {
                        "name": tc.function.name,
                        "arguments": tc.function.arguments,
                    },
                }

    text = "".join(text_parts)
    tool_calls = [tool_calls_dict[i] for i in sorted(tool_calls_dict)]
    return text, tool_calls


# ── Main loop ─────────────────────────────────────────────────────────


def run_chat_loop(
    session: PresentationSession,
    provider: str = "openai",
    api_key: str | None = None,
    model: str | None = None,
) -> None:
    """Run the interactive presentation editing loop.

    Args:
        session:  PresentationSession holding current state.
        provider: LLM provider key — same options as ai-build (openai, minimax, …).
        api_key:  Optional override for the provider's API key.
        model:    Optional model override (e.g. "gpt-4o", "gpt-4-turbo").
    """
    client, resolved_model = _make_client(provider, api_key, model)

    # Conversation history (system message always first)
    messages: list[dict[str, Any]] = [
        {"role": "system", "content": _SYSTEM_PROMPT}
    ]

    console.print()
    console.print(
        Panel(
            f"[bold]Chat mode[/bold]  •  provider: [cyan]{provider}[/cyan]  "
            f"•  model: [cyan]{resolved_model}[/cyan]\n"
            f"Presentation: [cyan]{session.output_path}[/cyan]\n"
            f"Topic: {session.topic}  |  "
            f"Theme: {session.theme_key}  |  "
            f"Slides: {len(session.selected_slides)}\n\n"
            "Describe what you'd like to change.\n"
            'Type [bold]exit[/bold] or press [bold]Ctrl+C[/bold] to save and exit.',
            title="PPT Master Chat",
            border_style="blue",
        )
    )
    console.print()

    while True:
        # ── User input ────────────────────────────────────────────────
        try:
            user_input = console.input("[bold cyan]You:[/bold cyan] ").strip()
        except (EOFError, KeyboardInterrupt):
            console.print("\n[yellow]Exiting chat...[/yellow]")
            break

        if not user_input:
            continue
        if user_input.lower() in ("exit", "quit", "q", "done"):
            break

        messages.append({"role": "user", "content": user_input})

        # ── Agentic loop: keep going until the LLM stops calling tools ─
        while True:
            try:
                text, tool_calls = _call_with_tools(client, resolved_model, messages)
            except AuthenticationError:
                console.print(
                    "\n[red]Authentication failed.[/red] "
                    f"Check your API key for provider '{provider}'."
                )
                return
            except APIConnectionError:
                console.print(
                    "\n[red]Connection error.[/red] Check your network / base URL."
                )
                # Remove last user message so they can retry
                if messages[-1]["role"] == "user":
                    messages.pop()
                break
            except APIStatusError as exc:
                console.print(f"\n[red]API error {exc.status_code}:[/red] {exc.message}")
                if messages[-1]["role"] == "user":
                    messages.pop()
                break

            # Append assistant message (content + tool_calls) to history
            assistant_msg: dict[str, Any] = {"role": "assistant", "content": text}
            if tool_calls:
                assistant_msg["tool_calls"] = tool_calls
            messages.append(assistant_msg)

            if not tool_calls:
                break  # No tools requested — turn is complete

            # ── Execute tool calls ─────────────────────────────────────
            for tc in tool_calls:
                tool_name = tc["function"]["name"]
                raw_args = tc["function"]["arguments"]

                # Always json.loads() the arguments string
                try:
                    tool_input = json.loads(raw_args) if raw_args else {}
                except json.JSONDecodeError:
                    tool_input = {}

                args_preview = ", ".join(
                    f"{k}={repr(v)[:50]}" for k, v in tool_input.items()
                )
                console.print(f"  [dim]→ {tool_name}({args_preview})[/dim]")

                needs_spinner = tool_name in ("add_slide", "regenerate_slide")
                if needs_spinner:
                    with console.status(
                        f"[dim]  Generating content for "
                        f"{tool_input.get('slide_type', '')}...[/dim]"
                    ):
                        result = execute_tool(session, tool_name, tool_input)
                else:
                    result = execute_tool(session, tool_name, tool_input)

                # Feed result back as a "tool" role message
                messages.append(
                    {
                        "role": "tool",
                        "tool_call_id": tc["id"],
                        "content": result,
                    }
                )
            # Loop back so the LLM can respond to the tool results

        # ── Rebuild PPTX if state changed ─────────────────────────────
        if session._dirty:
            with console.status("[bold magenta]Rebuilding presentation...[/bold magenta]"):
                try:
                    result_path = session.rebuild()
                    console.print(
                        f"  [green]✓[/green] Saved → [cyan]{result_path}[/cyan]  "
                        f"({len(session.selected_slides)} content slides)"
                    )
                except Exception as exc:
                    console.print(f"  [red]✗ Build failed:[/red] {exc}")
                    session._dirty = True  # Retry next turn

        console.print()

    # ── Final save on exit ────────────────────────────────────────────
    if session._dirty:
        with console.status("[bold magenta]Saving final presentation...[/bold magenta]"):
            try:
                result_path = session.rebuild()
                console.print(
                    f"\n[bold green]✓ Final presentation saved → {result_path}[/bold green]"
                )
            except Exception as exc:
                console.print(f"\n[red]Final save failed:[/red] {exc}")
    else:
        console.print(
            f"\n[bold green]✓ Presentation ready: {session.output_path}[/bold green]"
        )
