"""Command-line interface for KimiAgent.

Examples
--------
Generate a deck from a topic — no API key needed, and you get both formats::

    kimiagent generate --topic "The State of Renewable Energy in 2026" \\
        --theme slate --format both -o energy

Give it an outline to structure the sections::

    kimiagent generate --topic "Q3 Business Review" \\
        --outline "Highlights; Revenue; Product; Risks; Next quarter" \\
        --audience "Board of directors" --theme corporate -o q3.pptx

Render a hand-written or previously exported spec::

    kimiagent from-spec my_deck.json -o my_deck.pdf

Convert an existing spec to PDF only::

    kimiagent from-spec my_deck.json --format pdf -o my_deck.pdf

Inspect the design health of a deck without rendering it::

    kimiagent critique my_deck.json

List themes (with a rendered sample of each)::

    kimiagent themes
    kimiagent themes --sample build/samples
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import List, Optional, Tuple

from .agent import DeckAgent, critique
from .design import DENSITY
from .models import Deck
from .pdf_renderer import render_pdf, soffice_available
from .renderer import render_deck
from .themes import THEME_GROUPS, THEMES, theme_names


def _slugify(text: str) -> str:
    keep = [c.lower() if c.isalnum() else "-" for c in (text or "").strip()]
    slug = "".join(keep)
    while "--" in slug:
        slug = slug.replace("--", "-")
    return slug.strip("-") or "deck"


def _resolve_outputs(output: Optional[str], fmt: str, deck: Deck) -> List[Tuple[str, str]]:
    """Work out which files to write.

    Accepts a bare stem (``-o report``), an explicit filename
    (``-o report.pdf``) or nothing at all, and reconciles it with
    ``--format``. An explicit extension wins over ``--format`` when only one
    format was requested, because that is clearly what the user meant.
    """
    fmt = (fmt or "pptx").lower()
    formats = ["pptx", "pdf"] if fmt in ("both", "all") else [fmt]

    if not output:
        stem = _slugify(deck.title)
    else:
        root, ext = os.path.splitext(output)
        ext = ext.lower().lstrip(".")
        if ext in ("pptx", "pdf"):
            if len(formats) == 1:
                # A single explicit filename: honour its extension.
                return [(ext, output)]
            stem = root
        else:
            stem = output

    return [(f, f"{stem}.{f}") for f in formats]


def _write(deck: Deck, targets: List[Tuple[str, str]], pdf_engine: str) -> List[str]:
    written: List[str] = []
    for kind, path in targets:
        if kind == "pdf":
            written.append(render_pdf(deck, path, engine=pdf_engine))
        else:
            written.append(render_deck(deck, path))
    return written


def _report(deck: Deck) -> None:
    hist = deck.type_histogram()
    print(f"[kimiagent] {deck.slide_count()} slides · "
          f"{deck.variety_score():.0%} non-bullet content slides")
    order = sorted(hist.items(), key=lambda kv: (-kv[1], kv[0]))
    print("[kimiagent] layouts: " + ", ".join(f"{k}×{v}" for k, v in order))
    notes = critique(deck)
    if notes:
        print("[kimiagent] design notes:")
        for n in notes:
            print(f"  · {n}")
    else:
        print("[kimiagent] design notes: none — deck looks healthy.")


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
    if args.density:
        deck.density = args.density
    if args.aspect:
        deck.aspect = args.aspect
    if args.author:
        deck.author = args.author
    if args.footer_note:
        deck.footer_note = args.footer_note

    if args.emit_spec:
        parent = os.path.dirname(os.path.abspath(args.emit_spec))
        os.makedirs(parent, exist_ok=True)
        with open(args.emit_spec, "w", encoding="utf-8") as f:
            json.dump(deck.to_dict(), f, indent=2, ensure_ascii=False)
        print(f"[kimiagent] Wrote deck spec -> {args.emit_spec}")

    for path in _write(deck, _resolve_outputs(args.output, args.format, deck), args.pdf_engine):
        print(f"[kimiagent] Rendered {deck.slide_count()} slides -> {path}")
    if args.report:
        _report(deck)
    return 0


def cmd_from_spec(args: argparse.Namespace) -> int:
    with open(args.spec, "r", encoding="utf-8") as f:
        data = json.load(f)
    deck = Deck.from_dict(data)
    if args.theme:
        deck.theme = args.theme
    if args.density:
        deck.density = args.density
    if args.aspect:
        deck.aspect = args.aspect

    for path in _write(deck, _resolve_outputs(args.output, args.format, deck), args.pdf_engine):
        print(f"[kimiagent] Rendered {deck.slide_count()} slides -> {path}")
    if args.report:
        _report(deck)
    return 0


def cmd_critique(args: argparse.Namespace) -> int:
    """Report on a deck's design health without rendering it."""
    with open(args.spec, "r", encoding="utf-8") as f:
        deck = Deck.from_dict(json.load(f))
    _report(deck)
    return 0


def cmd_themes(args: argparse.Namespace) -> int:
    print("Available themes:\n")
    for group, names in THEME_GROUPS.items():
        print(f"  {group.upper()}")
        for name in names:
            t = THEMES.get(name)
            if not t:
                continue
            print(f"    {name:<11} {t.blurb}")
            print(f"    {'':<11} primary #{t.primary}  accent #{t.secondary}  "
                  f"radius {t.radius:g}  {t.heading_font}/{t.body_font}")
        print()

    if getattr(args, "sample", None):
        # Rendering a deck per theme is the only honest way to choose one.
        print(f"\nRendering samples into {args.sample}/ ...")
        _theme_sample(args.sample)
    return 0


def _theme_sample(target_dir: str) -> None:
    """Render the bundled showcase deck in every theme."""
    from .agent import DeckAgent

    os.makedirs(target_dir, exist_ok=True)
    agent = DeckAgent(api_key=None)
    for name in theme_names():
        deck = agent.build_offline("Design system showcase", slide_count=20, theme=name)
        out = os.path.join(target_dir, f"{name}.pdf")
        render_pdf(deck, out)
        print(f"  wrote {out}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="kimiagent",
        description=(
            "KimiAgent — turn a topic or outline into a designer-quality "
            "PowerPoint deck and PDF."
        ),
    )
    sub = parser.add_subparsers(dest="command", required=True)

    def add_output_opts(p: argparse.ArgumentParser, *, spec_mode: bool = False) -> None:
        p.add_argument("--output", "-o", help="Output path or stem (extension optional).")
        p.add_argument(
            "--format", "-f", default="pptx", choices=["pptx", "pdf", "both"],
            help="Which file(s) to write (default pptx).",
        )
        p.add_argument(
            "--pdf-engine", default="native", choices=["native", "soffice", "auto"],
            help=(
                "native (default) draws the PDF directly — no system deps. "
                "soffice converts the .pptx via LibreOffice. auto prefers "
                "LibreOffice when installed."
            ),
        )
        p.add_argument("--theme", choices=theme_names(),
                       help="Visual theme." if spec_mode else "Visual theme.")
        p.add_argument("--density", choices=sorted(DENSITY), help="Layout density.")
        p.add_argument("--aspect", choices=["16:9", "4:3"], help="Slide aspect ratio.")
        p.add_argument("--report", action="store_true",
                       help="Print layout stats and design warnings.")

    g = sub.add_parser("generate", help="Generate a deck from a topic.")
    g.add_argument("--topic", "-t", required=True, help="What the deck is about.")
    g.add_argument("--outline", help="Optional section outline, ';' or newline separated.")
    g.add_argument("--slides", type=int, default=14, help="Approximate slide count (default 14).")
    g.add_argument("--audience", help="Target audience (used on the cover + tone).")
    g.add_argument("--tone", help="Desired tone, e.g. 'executive', 'friendly'.")
    g.add_argument("--author", help="Byline shown on the cover.")
    g.add_argument("--footer-note", help="Footer text, e.g. 'Confidential'.")
    g.add_argument("--emit-spec", help="Also write the generated deck spec as JSON to this path.")
    g.add_argument("--offline", action="store_true", help="Force the offline generator (no API).")
    g.add_argument("--llm", action="store_true", help="Force the LLM path (requires an API key).")
    g.add_argument("--api-key", help="Kimi/Moonshot API key (else read from env).")
    g.add_argument("--model", help="Model name (default kimi-k2 preview).")
    g.add_argument("--base-url", help="OpenAI-compatible base URL (default Moonshot).")
    # ``generate`` needs a default theme; ``from-spec`` must not override one.
    g.set_defaults(func=cmd_generate, theme="kimi")
    add_output_opts(g)

    fs = sub.add_parser("from-spec", help="Render a deck from a JSON spec file.")
    fs.add_argument("spec", help="Path to a deck spec JSON file.")
    fs.set_defaults(func=cmd_from_spec)
    add_output_opts(fs, spec_mode=True)

    cr = sub.add_parser("critique", help="Report design warnings for a spec, without rendering.")
    cr.add_argument("spec", help="Path to a deck spec JSON file.")
    cr.set_defaults(func=cmd_critique)

    th = sub.add_parser("themes", help="List available visual themes.")
    th.add_argument("--sample", metavar="DIR",
                    help="Also render a sample PDF per theme into DIR.")
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
    except json.JSONDecodeError as exc:
        print(f"[kimiagent] Invalid JSON in spec: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover
        print(f"[kimiagent] Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
