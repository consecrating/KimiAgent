"""Example: build a deck programmatically with the KimiAgent Python API.

Run from the repository root::

    python examples/generate_example.py

Produces two .pptx files in ./out/ — one from a topic (offline generator)
and one from the hand-written sample spec — with no API key required.
"""

import os
import sys

# Make the package importable when run directly from the repo (no install needed).
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from kimiagent import DeckAgent, Deck, render_deck


def main() -> None:
    os.makedirs("out", exist_ok=True)

    # 1) Generate from a topic using the offline generator (no API key).
    agent = DeckAgent()
    deck = agent.build_offline(
        topic="Scaling a Product Team from 5 to 50",
        audience="Engineering leadership",
        theme="kimi",
    )
    render_deck(deck, "out/scaling-a-product-team.pptx")
    print(f"Built {len(deck.slides)} slides -> out/scaling-a-product-team.pptx")

    # 2) Render the hand-written sample spec.
    import json

    with open(os.path.join(os.path.dirname(__file__), "sample_spec.json"), encoding="utf-8") as f:
        spec = json.load(f)
    deck2 = Deck.from_dict(spec)
    render_deck(deck2, "out/renewable-energy-2026.pptx")
    print(f"Built {len(deck2.slides)} slides -> out/renewable-energy-2026.pptx")


if __name__ == "__main__":
    main()
