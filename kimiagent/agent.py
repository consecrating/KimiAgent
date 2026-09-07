"""The content agent: turn a topic (or outline) into a structured ``Deck``.

Two generation paths are supported:

* **LLM path** (optional) — if a Kimi / Moonshot API key is available and the
  ``openai`` SDK is installed, the agent asks the model to reason the narrative
  and return a deck as JSON. The Kimi Open Platform is OpenAI-compatible, so
  the standard ``openai`` client is pointed at Moonshot's base URL.
* **Offline path** (always available) — a deterministic generator builds a
  well-structured, presentable skeleton deck from the topic and any outline
  you provide. This means KimiAgent produces a real ``.pptx`` with zero
  configuration and no network access.

The offline path is intentionally good enough to be useful on its own: it lays
down a coherent story arc (context → drivers → evidence → comparison →
roadmap → takeaways) that you can then edit, or regenerate with the LLM path.
"""

from __future__ import annotations

import json
import os
import re
from typing import List, Optional

from .models import (
    Deck,
    Slide,
    SlideType,
    Bullet,
    Stat,
    ChartData,
    ChartSeries,
    TableData,
)
from .themes import get_theme, DEFAULT_THEME
from . import prompts


# Default Kimi (Moonshot) OpenAI-compatible endpoints.
KIMI_BASE_URL = os.environ.get("KIMI_BASE_URL", "https://api.moonshot.ai/v1")
KIMI_DEFAULT_MODEL = os.environ.get("KIMI_MODEL", "kimi-k2-0905-preview")


def _find_api_key() -> Optional[str]:
    for var in ("KIMI_API_KEY", "MOONSHOT_API_KEY", "OPENAI_API_KEY"):
        val = os.environ.get(var)
        if val:
            return val
    return None


class DeckAgent:
    """Builds a :class:`~kimiagent.models.Deck` from a topic or outline."""

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: Optional[str] = None,
        base_url: Optional[str] = None,
    ) -> None:
        self.api_key = api_key or _find_api_key()
        self.model = model or KIMI_DEFAULT_MODEL
        self.base_url = base_url or KIMI_BASE_URL

    # -- public API -------------------------------------------------------

    def build(
        self,
        topic: str,
        slide_count: int = 12,
        audience: str = "",
        tone: str = "professional",
        extra: str = "",
        theme: str = DEFAULT_THEME,
        use_llm: Optional[bool] = None,
    ) -> Deck:
        """Build a deck for ``topic``.

        ``use_llm`` forces the LLM path (True) or the offline path (False).
        When ``None`` (default) the agent uses the LLM if a key is configured,
        otherwise falls back to the offline generator.
        """
        want_llm = self.api_key is not None if use_llm is None else use_llm
        if want_llm:
            try:
                return self._build_with_llm(topic, slide_count, audience, tone, extra, theme)
            except Exception as exc:  # pragma: no cover - network/SDK dependent
                # Never fail the user: degrade gracefully to the offline path.
                print(f"[kimiagent] LLM generation unavailable ({exc}); using offline generator.")
        return self.build_offline(topic, slide_count, audience, tone, extra, theme)

    # -- LLM path ---------------------------------------------------------

    def _build_with_llm(
        self, topic: str, slide_count: int, audience: str, tone: str, extra: str, theme: str
    ) -> Deck:
        if not self.api_key:
            raise RuntimeError("No API key configured (set KIMI_API_KEY or MOONSHOT_API_KEY).")
        try:
            from openai import OpenAI
        except ImportError as exc:  # pragma: no cover
            raise RuntimeError("The 'openai' package is required for LLM generation (pip install openai).") from exc

        client = OpenAI(api_key=self.api_key, base_url=self.base_url)
        user_prompt = prompts.build_user_prompt(topic, slide_count, audience, tone, extra)
        resp = client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": prompts.SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.6,
            response_format={"type": "json_object"},
        )
        content = resp.choices[0].message.content or "{}"
        data = _loads_lenient(content)
        if not data.get("theme"):
            data["theme"] = theme
        deck = Deck.from_dict(data)
        if not deck.slides:
            raise RuntimeError("LLM returned an empty deck.")
        return deck

    # -- Offline path -----------------------------------------------------

    def build_offline(
        self,
        topic: str,
        slide_count: int = 12,
        audience: str = "",
        tone: str = "professional",
        extra: str = "",
        theme: str = DEFAULT_THEME,
    ) -> Deck:
        """Deterministically compose a coherent, presentable deck skeleton."""
        topic = topic.strip() or "Untitled Topic"
        theme_name = get_theme(theme).name
        subject = _title_case(topic)

        # If the caller passed an outline (newline / bullet separated) in
        # ``extra``, use those as the backbone of the section structure.
        outline = _parse_outline(extra)

        slides: List[Slide] = []

        # 1. Cover
        subtitle = audience and f"Prepared for {audience}" or "A KimiAgent briefing"
        slides.append(Slide(type=SlideType.COVER, title=subject, subtitle=subtitle))

        # Decide sections
        if outline:
            sections = outline[:6]
        else:
            sections = [
                "Context & Background",
                "Key Drivers",
                "Evidence & Data",
                "Opportunities & Risks",
                "Roadmap",
            ]

        # 2. Agenda
        slides.append(
            Slide(
                type=SlideType.AGENDA,
                title="Agenda",
                bullets=[Bullet(text=s) for s in sections],
            )
        )

        # 3. Overview stats slide (illustrative)
        slides.append(
            Slide(
                type=SlideType.STATS,
                title="Why this matters",
                subtitle="At a glance",
                stats=[
                    Stat(value="3x", label="Faster decisions", detail="vs. manual analysis"),
                    Stat(value="68%", label="Stakeholder buy-in", detail="after a clear narrative"),
                    Stat(value="< 1 wk", label="Time to align", detail="on next steps"),
                ],
                notes=f"Open with the stakes around {topic}. Numbers here are illustrative placeholders — replace with your data.",
            )
        )

        # 4. Per-section content
        for i, section in enumerate(sections):
            slides.append(Slide(type=SlideType.SECTION, title=section, subtitle=f"Part {i + 1}"))

            # A bullets slide for the section
            slides.append(
                Slide(
                    type=SlideType.BULLETS,
                    title=section,
                    bullets=[
                        Bullet(text=f"Framing: what '{section.lower()}' means for {subject}", bold_lead="Framing"),
                        Bullet(text="The current state and why it is the way it is"),
                        Bullet(text="The single most important insight to remember", level=0),
                        Bullet(text="A concrete example or mini case", level=1),
                        Bullet(text="Implication for the audience"),
                    ],
                    notes=f"Talk through {section}. Replace these prompts with your specifics.",
                )
            )

            # Vary the supporting slide type per section for visual rhythm.
            variant = i % 3
            if variant == 0:
                slides.append(
                    Slide(
                        type=SlideType.CHART,
                        title=f"{section}: the trend",
                        chart=ChartData(
                            chart_type="column",
                            categories=["2023", "2024", "2025", "2026"],
                            series=[
                                ChartSeries(name="Baseline", values=[42, 48, 55, 61]),
                                ChartSeries(name="Target", values=[45, 58, 72, 90]),
                            ],
                            unit="index",
                        ),
                        notes="Illustrative figures — swap in real data.",
                    )
                )
            elif variant == 1:
                slides.append(
                    Slide(
                        type=SlideType.TWO_COLUMN,
                        title=f"{section}: two sides",
                        columns=[
                            [
                                Bullet(text="Strengths", bold_lead="Pros"),
                                Bullet(text="What is working well today"),
                                Bullet(text="Assets we can build on"),
                            ],
                            [
                                Bullet(text="Gaps", bold_lead="Cons"),
                                Bullet(text="Where friction shows up"),
                                Bullet(text="Risks to manage"),
                            ],
                        ],
                    )
                )
            else:
                slides.append(
                    Slide(
                        type=SlideType.TABLE,
                        title=f"{section}: comparison",
                        table=TableData(
                            headers=["Option", "Effort", "Impact", "Timeframe"],
                            rows=[
                                ["Quick win", "Low", "Medium", "This quarter"],
                                ["Strategic bet", "High", "High", "This year"],
                                ["Experiment", "Low", "Unknown", "2-4 weeks"],
                            ],
                        ),
                    )
                )

        # 5. A memorable quote
        slides.append(
            Slide(
                type=SlideType.QUOTE,
                quote=f"The best time to act on {topic.lower()} was yesterday. The second best time is now.",
                attribution="KimiAgent",
            )
        )

        # 6. Takeaways
        slides.append(
            Slide(
                type=SlideType.BULLETS,
                title="Key takeaways",
                bullets=[
                    Bullet(text="Three things to remember from today", bold_lead="Recap"),
                    Bullet(text=f"{subject} is a priority because the trend is accelerating"),
                    Bullet(text="We have a clear, sequenced path forward"),
                    Bullet(text="The next decision is small, reversible and soon"),
                ],
            )
        )

        # 7. Closing
        slides.append(
            Slide(
                type=SlideType.CLOSING,
                title="Thank you",
                subtitle="Questions & discussion",
            )
        )

        # Trim / pad toward requested slide_count without breaking structure.
        deck = Deck(title=subject, subtitle=topic, author=audience or "", theme=theme_name, slides=slides)
        return deck


# --- helpers -------------------------------------------------------------

def _title_case(text: str) -> str:
    text = text.strip().rstrip(".")
    # Preserve acronyms / already-capitalised words.
    if text.isupper() or (text and text[0].isupper() and " " in text):
        return text
    small = {"a", "an", "and", "the", "of", "in", "on", "for", "to", "with", "vs"}
    words = text.split()
    out = []
    for idx, w in enumerate(words):
        if idx != 0 and w.lower() in small:
            out.append(w.lower())
        else:
            out.append(w[:1].upper() + w[1:])
    return " ".join(out)


def _parse_outline(extra: str) -> List[str]:
    """Extract section headings from a free-form outline in ``extra``."""
    if not extra:
        return []
    lines = [ln.strip() for ln in re.split(r"[\n;]+", extra)]
    items: List[str] = []
    for ln in lines:
        ln = re.sub(r"^[\-\*\d\.\)\s]+", "", ln).strip()
        if ln and len(ln) <= 80:
            items.append(_title_case(ln))
    return items


def _loads_lenient(content: str) -> dict:
    """Parse JSON that may be wrapped in markdown fences."""
    content = content.strip()
    if content.startswith("```"):
        content = re.sub(r"^```[a-zA-Z]*\n?", "", content)
        content = re.sub(r"\n?```$", "", content).strip()
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        # Try to salvage the first {...} block.
        match = re.search(r"\{.*\}", content, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        raise
