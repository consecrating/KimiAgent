"""The content agent: turn a topic (or outline) into a structured ``Deck``.

Two generation paths are supported:

* **LLM path** (optional) — if a Kimi / Moonshot API key is available and the
  ``openai`` SDK is installed, the agent asks the model to reason the narrative
  and return a deck as JSON. The Kimi Open Platform is OpenAI-compatible, so
  the standard ``openai`` client is pointed at Moonshot's base URL.
* **Offline path** (always available) — a deterministic generator builds a
  well-structured, presentable deck from the topic and any outline you
  provide. This means KimiAgent produces a real deck with zero configuration
  and no network access.

Both paths aim at the same thing: a deck whose *form varies with its content*.
The offline generator deliberately rotates through archetypes — a statement
slide for the thesis, cards for parallel ideas, a timeline for a sequence, a
comparison for a decision — because a deck of nothing but bullet lists looks
generated no matter how good the writing is. :meth:`Deck.variety_score`
measures this, and :func:`critique` reports on it.
"""

from __future__ import annotations

import json
import os
import re
from typing import Dict, List, Optional, Sequence

from . import prompts
from .models import (
    Bullet,
    Card,
    ChartData,
    ChartSeries,
    ComparisonSide,
    Deck,
    ProcessStep,
    Slide,
    SlideType,
    Stat,
    TableData,
    TimelineItem,
)
from .themes import DEFAULT_THEME, get_theme


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
            raise RuntimeError(
                "The 'openai' package is required for LLM generation (pip install openai)."
            ) from exc

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
        return _ensure_structure(deck, topic, audience)

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
        """Deterministically compose a coherent, presentable deck.

        The output is a *scaffold with a real narrative shape*, not filler: a
        thesis, sections that each open with a divider and carry two differently
        shaped content slides, a memorable line, and a close. Every slide
        carries prompts specific enough to edit rather than lorem ipsum.
        """
        topic = (topic or "").strip() or "Untitled Topic"
        theme_name = get_theme(theme).name
        subject = _title_case(topic)
        short = _short_subject(subject)

        outline = _parse_outline(extra)
        explicit = bool(outline)
        sections = outline[:6] if outline else _default_sections()

        # Seven slides are fixed (cover, agenda, thesis, stats, quote, recap,
        # closing); each section costs a divider plus one or two content slides.
        #
        # When the budget is tight something has to give, and the two options
        # trade off differently depending on where the outline came from:
        #   * a user-supplied outline is a contract, so every section is kept
        #     and the supporting slide is dropped instead;
        #   * an auto-generated outline is arbitrary, so sections are dropped
        #     and the supporting slides kept — variety is what makes the deck
        #     worth looking at.
        budget = max(8, int(slide_count or 12))
        fixed = 7
        room = max(0, budget - fixed)

        if explicit and room < len(sections) * 3:
            # Keep every section the user named; give up the supporting slide.
            per_section = 1
        elif room < 6:
            # Too tight for two sections of three slides: keep two sections so
            # the deck still has a middle, but only the prose slide in each.
            per_section = 1
            sections = sections[:2]
        else:
            per_section = 2
            sections = sections[: max(2, min(len(sections), room // 3))]

        slides: List[Slide] = []

        # 1. Cover
        slides.append(
            Slide(
                type=SlideType.COVER,
                title=subject,
                subtitle=_cover_subtitle(short, audience),
                kicker=_kicker_for(audience, tone),
            )
        )

        # 2. Agenda
        slides.append(
            Slide(
                type=SlideType.AGENDA,
                title="What we'll cover",
                bullets=[Bullet(text=s) for s in sections],
            )
        )

        # 3. The thesis, as a statement rather than a bullet. A deck should
        #    say what it thinks before it starts arguing.
        slides.append(
            Slide(
                type=SlideType.BIG_STATEMENT,
                kicker="The argument",
                title=f"{short} is changing faster than our plans for it.",
                subtitle="Replace this with the one sentence you want remembered.",
                notes=(
                    f"State the thesis plainly, then promise the evidence. "
                    f"Everything after this slide should support this claim about {topic}."
                ),
            )
        )

        # 4. Opening evidence: quantified stakes.
        slides.append(
            Slide(
                type=SlideType.STATS,
                kicker="At a glance",
                title="Why this matters now",
                subtitle="Illustrative figures — replace with your data",
                stats=[
                    Stat(value="3x", label="Faster decisions", detail="vs. manual analysis", trend="up"),
                    Stat(value="68%", label="Stakeholder buy-in", detail="after a clear narrative", trend="up"),
                    Stat(value="<1 wk", label="Time to align", detail="on next steps", trend="down"),
                ],
                notes=(
                    f"Open with the stakes around {topic}. These numbers are "
                    "placeholders — swap in real ones before presenting."
                ),
            )
        )

        # 5. Per-section content, rotating the supporting archetype so the deck
        #    has visual rhythm instead of twelve identical bullet slides.
        for i, section in enumerate(sections):
            slides.append(
                Slide(
                    type=SlideType.SECTION,
                    title=section,
                    subtitle=f"{i + 1:02d}",
                )
            )
            slides.append(_section_bullets(section, short))
            if per_section > 1:
                slides.append(_supporting_slide(i, section, short))

        # 6. A line worth remembering.
        slides.append(
            Slide(
                type=SlideType.QUOTE,
                quote=(
                    f"Waiting on {short} stopped being free some time ago. "
                    "The cost just compounds quietly."
                ),
                attribution="Replace with a real quote from your research",
            )
        )

        # 7. Takeaways as cards: three parallel points read better side by side
        #    than stacked as bullets.
        slides.append(
            Slide(
                type=SlideType.CARDS,
                kicker="Recap",
                title="Three things to take away",
                cards=[
                    Card(
                        title="The trend is real",
                        body=f"{short} is moving, and the direction is not in dispute.",
                        icon="1",
                    ),
                    Card(
                        title="We have a path",
                        body="The sequence is clear and the first step is small.",
                        icon="2",
                    ),
                    Card(
                        title="The next decision is soon",
                        body="Reversible, cheap, and worth making this quarter.",
                        icon="3",
                    ),
                ],
                notes="Land these three. If the audience remembers nothing else, this is it.",
            )
        )

        # 8. Closing
        slides.append(
            Slide(
                type=SlideType.CLOSING,
                title="Thank you",
                subtitle="Questions and discussion",
            )
        )

        return Deck(
            title=subject,
            subtitle=topic if topic != subject else "",
            # The audience is not the author; it belongs in the cover subtitle
            # ("Prepared for ...") which :func:`_cover_subtitle` already sets.
            author="",
            theme=theme_name,
            slides=slides,
        )


# --- offline content builders --------------------------------------------


def _default_sections() -> List[str]:
    """A generic but genuinely usable story arc for an unspecified topic."""
    return [
        "Where things stand",
        "What is driving it",
        "What the data shows",
        "Options and trade-offs",
        "What we do next",
    ]


def _section_bullets(section: str, short: str) -> Slide:
    """The prose slide for a section, with editable prompts."""
    return Slide(
        type=SlideType.BULLETS,
        title=section,
        bullets=[
            Bullet(text=f"What '{_lower_first(section)}' means for {short}.", bold_lead="Framing"),
            Bullet(text="The current state, and why it ended up this way."),
            Bullet(text="The single most important thing to remember here."),
            Bullet(text="A concrete example or mini case that proves it.", level=1),
            Bullet(text="What this implies for the people in the room.", bold_lead="So what"),
        ],
        notes=(
            f"Talk through {section}. Replace each prompt with your specifics — "
            "the structure is the point, the words are not."
        ),
    )


def _supporting_slide(index: int, section: str, short: str) -> Slide:
    """Pick a supporting archetype for a section, rotating through the set."""
    variant = index % 5

    if variant == 0:
        return Slide(
            type=SlideType.CHART,
            title=f"{section}: the trend",
            kicker="Evidence",
            chart=ChartData(
                chart_type="line",
                categories=["2023", "2024", "2025", "2026"],
                series=[
                    ChartSeries(name="Baseline", values=[42, 48, 55, 61]),
                    ChartSeries(name="Target", values=[45, 58, 72, 90]),
                ],
                unit="index",
                takeaway="State the conclusion here — do not make the audience read the chart.",
            ),
            notes="Illustrative figures. Lead with the shape of the line, not the values.",
        )

    if variant == 1:
        return Slide(
            type=SlideType.CARDS,
            title=f"{section}: the moving parts",
            cards=[
                Card(title="Driver one", body="What it is and why it matters.", icon="1"),
                Card(title="Driver two", body="The second force at work here.", icon="2"),
                Card(title="Driver three", body="The one people usually miss.", icon="3"),
            ],
        )

    if variant == 2:
        return Slide(
            type=SlideType.COMPARISON,
            title=f"{section}: the choice",
            verdict="State your recommendation here, in one sentence.",
            sides=[
                ComparisonSide(
                    title="Option A",
                    subtitle="The conventional move",
                    points=[
                        Bullet(text="Proven and quick to start"),
                        Bullet(text="Limited upside"),
                        Bullet(text="What could go wrong", bold_lead="Risk"),
                    ],
                ),
                ComparisonSide(
                    title="Option B",
                    subtitle="The bet worth making",
                    preferred=True,
                    points=[
                        Bullet(text="Attacks the real constraint"),
                        Bullet(text="Compounds over time"),
                        Bullet(text="Needs more up-front work", bold_lead="Risk"),
                    ],
                ),
            ],
        )

    if variant == 3:
        return Slide(
            type=SlideType.TIMELINE,
            title=f"{section}: how it unfolds",
            timeline=[
                TimelineItem(period="Now", title="Where we are", detail="The current state.", done=True),
                TimelineItem(period="Next quarter", title="First move", detail="Small and reversible."),
                TimelineItem(period="This year", title="Scale what works", detail="Double down on evidence."),
                TimelineItem(period="Next year", title="The payoff", detail="What success looks like."),
            ],
        )

    return Slide(
        type=SlideType.TABLE,
        title=f"{section}: options compared",
        table=TableData(
            headers=["Option", "Effort", "Impact", "Timeframe"],
            numeric_columns=[],
            rows=[
                ["Quick win", "Low", "Medium", "This quarter"],
                ["Strategic bet", "High", "High", "This year"],
                ["Experiment", "Low", "Unknown", "2–4 weeks"],
            ],
        ),
    )


def _cover_subtitle(short: str, audience: str) -> str:
    if audience:
        return f"Prepared for {audience}"
    return f"A briefing on {short}"


def _kicker_for(audience: str, tone: str) -> str:
    if audience:
        return "Briefing"
    if (tone or "").lower() in ("executive", "board"):
        return "Executive summary"
    return "Overview"


# --- post-processing -----------------------------------------------------


def _ensure_structure(deck: Deck, topic: str, audience: str) -> Deck:
    """Repair an LLM deck that is missing its structural bookends.

    Models reliably produce good *content* and unreliably produce a complete
    *shape* — a deck that opens on a bullet slide or never closes. Rather than
    reject the response, the missing pieces are added, because a repaired deck
    is more useful than an error message.
    """
    if not deck.slides:
        return deck
    if deck.slides[0].type != SlideType.COVER:
        deck.slides.insert(
            0,
            Slide(
                type=SlideType.COVER,
                title=deck.title or _title_case(topic),
                subtitle=deck.subtitle or _cover_subtitle(_short_subject(deck.title or topic), audience),
            ),
        )
    if deck.slides[-1].type != SlideType.CLOSING:
        deck.slides.append(
            Slide(type=SlideType.CLOSING, title="Thank you", subtitle="Questions and discussion")
        )
    return deck


def critique(deck: Deck) -> List[str]:
    """Return human-readable warnings about a deck's design health.

    Surfaced by ``kimiagent generate --report``. These are the checks a
    designer would make on a first pass, and they catch the failure modes that
    survive rendering: a deck that is all bullets, slides with no takeaway,
    charts nobody interpreted, walls of text.
    """
    notes: List[str] = []
    variety = deck.variety_score()
    if variety < 0.35:
        notes.append(
            f"Low layout variety ({variety:.0%} of content slides are non-bullet). "
            "Consider cards, stats, a timeline or a comparison for some of them."
        )

    hist = deck.type_histogram()
    if hist.get("bullets", 0) > max(3, len(deck.slides) // 3):
        notes.append(f"{hist['bullets']} bullet slides — the densest ones may work better as cards.")

    for i, s in enumerate(deck.slides, 1):
        if s.type == SlideType.CHART and s.chart and not s.chart.takeaway:
            notes.append(f"Slide {i}: chart has no takeaway line — state what it shows.")
        if s.type == SlideType.BULLETS and len(s.bullets) > 6:
            notes.append(f"Slide {i}: {len(s.bullets)} bullets; it will be split across slides.")
        for b in s.bullets:
            if len(b.text) > 160:
                notes.append(f"Slide {i}: a bullet runs {len(b.text)} characters — tighten it to a phrase.")
                break
        if s.type == SlideType.COMPARISON and s.sides and not s.verdict:
            notes.append(f"Slide {i}: a comparison without a verdict leaves the decision to the audience.")
        if s.type == SlideType.IMAGE and not s.image_path:
            notes.append(f"Slide {i}: image slide has no image_path; a placeholder will render.")

    if not any(s.notes for s in deck.slides):
        notes.append("No speaker notes anywhere — add them for the slides you will actually talk to.")
    return notes


# --- helpers -------------------------------------------------------------


def _title_case(text: str) -> str:
    text = (text or "").strip().rstrip(".")
    if not text:
        return "Untitled"
    # Preserve acronyms / already-capitalised words.
    if text.isupper() or (text[0].isupper() and " " in text):
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


def _short_subject(subject: str) -> str:
    """A compact noun phrase for use mid-sentence.

    Long titles ("The State of Renewable Energy in 2026") read badly inside a
    generated sentence, so leading articles and trailing date clauses are
    trimmed.
    """
    s = (subject or "").strip()
    s = re.sub(r"^(the|a|an)\s+", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^state of\s+", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+in\s+(19|20)\d{2}$", "", s)
    words = s.split()
    if len(words) > 6:
        s = " ".join(words[:6])
    return s or subject or "the topic"


def _lower_first(text: str) -> str:
    """Lowercase the first word unless it looks like an acronym.

    Intended for *sentence-case* strings such as section headings ("Where
    things stand" -> "where things stand") so they can be quoted mid-sentence.
    Do not use it on title-cased topic names: lowercasing only the first word
    of "Renewable Energy" yields "renewable Energy".
    """
    if not text:
        return text
    first, _, rest = text.partition(" ")
    if first.isupper() or (len(first) > 1 and first[1:].lower() != first[1:]):
        return text
    return first.lower() + (" " + rest if rest else "")


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
    content = (content or "").strip()
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


__all__ = ["DeckAgent", "critique", "KIMI_BASE_URL", "KIMI_DEFAULT_MODEL"]
