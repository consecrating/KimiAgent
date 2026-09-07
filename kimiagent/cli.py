"""Command-line interface for KimiAgent.

Examples
--------
Generate a deck from a topic (no API key needed — uses the offline generator)::

    kimiagent generate --topic "The State of Renewable Energy in 2026" -o energy.pptx

Give it an outline to structure the sections::

    kimiagent generate --topic "Q3 Business Review" \\
        --outline "Highlights; Revenue; Product; Risks; Next quarter" \\
        --theme corporate -o q3.pptx

Render a hand-written / previously exported deck spec::

    kimiagent from-spec my_deck.json -o my_deck.pptx

Export the generated spec as JSON (to tweak, then re-render)::

    kimiagent generate --topic "AI in Healthcare" --emit-spec spec.json -o ai.pptx

List available themes::

    kimiagent themes
"""

from __future__ import annotations

import argparse
import json
import os
import sys

from .agent import DeckAgent
from .renderer import render_deck
from .models import Deck
from .themes import theme_names, THEMES


def _slugify(text: str) -> str:
    keep = [c.lower() if c.isalnum() else "-" for c in text.strip()]
    slug = "".join(keep)
    while "--" in slug:
        slug = slug.replace("--", "-")
    return slug.strip("-") or "deck"


def cmd_generate(args: argparse.Namespace) -> int:
    agent = DeckAgent(api_key=args.api_key, model=args.model, base_url=args.base_url)

    use_llm = None
    if args.offline:
        use_llm = False
    elif args.llm:
        use_llm = True

    print(f"[kimiagent] Building deck for: {args.topic!r}")
    if use_llm is None and agent.api_key is None:
        print("[kimiagent] No API key found — using the built-in offline generator.")

    deck = agent.build(
        topic=args.topic,
        slide_count=args.slides,
        audience=args.audience or "",
        tone=args.tone or "",
        extra=args.outline or "",
        theme=args.theme,
        use_llm=use_llm,
    )
    deck.theme = args.theme or deck.theme

    if args.emit_spec:
        emit_parent = os.path.dirname(os.path.abspath(args.emit_spec))
        os.makedirs(emit_parent, exist_ok=True)
        with open(args.emit_spec, "w", encoding="utf-8") as f:
            json.dump(deck.to_dict(), f, indent=2, ensure_ascii=False)
        print(f"[kimiagent] Wrote deck spec -> {args.emit_spec}")

    output = args.output or f"{_slugify(deck.title)}.pptx"
    render_deck(deck, output)
    print(f"[kimiagent] Rendered {len(deck.slides)} slides -> {output}")
    return 0


def cmd_from_spec(args: argparse.Namespace) -> int:
    with open(args.spec, "r", encoding="utf-8") as f:
        data = json.load(f)
    deck = Deck.from_dict(data)
    if args.theme:
        deck.theme = args.theme
    output = args.output or f"{_slugify(deck.title)}.pptx"
    render_deck(deck, output)
    print(f"[kimiagent] Rendered {len(deck.slides)} slides -> {output}")
    return 0


def cmd_themes(_args: argparse.Namespace) -> int:
    print("Available themes:\n")
    for name in theme_names():
        t = THEMES[name]
        print(f"  {name:<10}  primary #{t.primary}  accent #{t.secondary}  ({t.heading_font})")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="kimiagent",
        description="KimiAgent — turn a topic or outline into a native, designer-quality PowerPoint deck.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    g = sub.add_parser("generate", help="Generate a deck from a topic.")
    g.add_argument("--topic", "-t", required=True, help="What the deck is about.")
    g.add_argument("--output", "-o", help="Output .pptx path (default: derived from title).")
    g.add_argument("--outline", help="Optional section outline, ';' or newline separated.")
    g.add_argument("--slides", type=int, default=12, help="Approximate slide count (default 12).")
    g.add_argument("--theme", default="kimi", choices=theme_names(), help="Visual theme.")
    g.add_argument("--audience", help="Target audience (used on the cover + tone).")
    g.add_argument("--tone", help="Desired tone, e.g. 'executive', 'friendly'.")
    g.add_argument("--emit-spec", help="Also write the generated deck spec as JSON to this path.")
    g.add_argument("--offline", action="store_true", help="Force the offline generator (no API).")
    g.add_argument("--llm", action="store_true", help="Force the LLM path (requires an API key).")
    g.add_argument("--api-key", help="Kimi/Moonshot API key (else read from env).")
    g.add_argument("--model", help="Model name (default kimi-k2 preview).")
    g.add_argument("--base-url", help="OpenAI-compatible base URL (default Moonshot).")
    g.set_defaults(func=cmd_generate)

    fs = sub.add_parser("from-spec", help="Render a deck from a JSON spec file.")
    fs.add_argument("spec", help="Path to a deck spec JSON file.")
    fs.add_argument("--output", "-o", help="Output .pptx path.")
    fs.add_argument("--theme", choices=theme_names(), help="Override the theme in the spec.")
    fs.set_defaults(func=cmd_from_spec)

    th = sub.add_parser("themes", help="List available visual themes.")
    th.set_defaults(func=cmd_themes)

    return parser


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except FileNotFoundError as exc:
        print(f"[kimiagent] File not found: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover
        print(f"[kimiagent] Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
