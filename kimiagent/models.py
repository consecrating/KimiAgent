"""Data model for a KimiAgent deck.

A ``Deck`` is a serialisable, tool-agnostic description of a presentation — the
"reason first, design later" intermediate representation. The agent produces a
``Deck``; a renderer turns it into a native ``.pptx`` or ``.pdf``.

Everything here is plain dataclasses so a deck round-trips cleanly to and from
JSON, which makes decks easy to inspect, hand-edit, diff and version-control.

Backward compatibility
----------------------
Every field added since the first release has a default, and
:meth:`Slide.from_dict` accepts both the compact shorthand (a bullet as a bare
string) and the full object form. Specs written against the original five slide
types still load and render unchanged.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional


class SlideType(str, Enum):
    """The supported slide archetypes.

    Each archetype maps to a dedicated layout routine in both renderers.

    The set is deliberately wider than "title + bullets" because bullet lists
    are the least effective way to carry most ideas. A sequence in time wants
    a :attr:`TIMELINE`; a set of peer concepts wants :attr:`CARDS`; a single
    claim wants the whole slide (:attr:`BIG_STATEMENT`). Giving the agent
    richer archetypes to choose from is what moves the output from
    "transcribed notes" to "designed deck".
    """

    # -- Structural ------------------------------------------------------
    COVER = "cover"                  # Title slide / deck opener
    AGENDA = "agenda"                # Table of contents
    SECTION = "section"              # Section divider
    CLOSING = "closing"              # Thank-you / call to action

    # -- Prose ------------------------------------------------------------
    BULLETS = "bullets"              # Title + bullet points
    TWO_COLUMN = "two_column"        # Two side-by-side bullet columns
    QUOTE = "quote"                  # Pull quote
    BIG_STATEMENT = "big_statement"  # One oversized claim, full bleed

    # -- Structured -------------------------------------------------------
    STATS = "stats"                  # Big-number KPI / metric callouts
    CARDS = "cards"                  # Grid of titled cards with icons
    TIMELINE = "timeline"            # Milestones along an axis
    PROCESS = "process"              # Numbered / chevroned sequence
    COMPARISON = "comparison"        # Two options weighed, with a verdict

    # -- Data -------------------------------------------------------------
    CHART = "chart"                  # Native, editable chart
    TABLE = "table"                  # Native table
    IMAGE = "image"                  # Full-bleed / captioned image


#: Archetypes that are rendered full-bleed and therefore carry no footer.
FULL_BLEED_TYPES = frozenset(
    {SlideType.COVER, SlideType.SECTION, SlideType.CLOSING, SlideType.BIG_STATEMENT}
)


@dataclass
class Bullet:
    """A single bullet point, optionally nested and optionally with a lead-in."""

    text: str
    level: int = 0                    # indentation level (0 = top level)
    bold_lead: Optional[str] = None    # optional bold prefix, e.g. "Key point:"

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {"text": self.text}
        if self.level:
            data["level"] = self.level
        if self.bold_lead:
            data["bold_lead"] = self.bold_lead
        return data


@dataclass
class Stat:
    """A large headline metric with a short label and optional detail."""

    value: str                      # e.g. "42%", "$3.1B", "1.2M"
    label: str                      # e.g. "YoY growth"
    detail: Optional[str] = None    # small supporting line
    #: Optional directional cue: "up", "down" or "flat". Renderers draw a
    #: small arrow glyph, which lets a reader grasp the trend pre-attentively
    #: instead of parsing the number first.
    trend: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class Card:
    """A titled card in a :attr:`SlideType.CARDS` grid.

    Cards are the workhorse of modern slide design: they let several peer
    ideas share a slide with equal visual weight and clear separation, which
    a bullet list cannot do.
    """

    title: str
    body: str = ""
    #: A short glyph or emoji shown above the title. Kept as text rather than
    #: an image asset so decks stay self-contained and dependency-free.
    icon: str = ""
    #: Optional index into the theme's accent ramp. ``None`` = use position.
    accent: Optional[int] = None

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {"title": self.title}
        if self.body:
            data["body"] = self.body
        if self.icon:
            data["icon"] = self.icon
        if self.accent is not None:
            data["accent"] = self.accent
        return data


@dataclass
class TimelineItem:
    """One milestone on a :attr:`SlideType.TIMELINE`."""

    period: str            # e.g. "Q1 2026", "Phase 1", "2019"
    title: str
    detail: str = ""
    #: Marks the milestone as reached, so renderers can de-emphasise it.
    done: bool = False

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {"period": self.period, "title": self.title}
        if self.detail:
            data["detail"] = self.detail
        if self.done:
            data["done"] = True
        return data


@dataclass
class ProcessStep:
    """One step in a :attr:`SlideType.PROCESS` flow."""

    title: str
    detail: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {"title": self.title}
        if self.detail:
            data["detail"] = self.detail
        return data


@dataclass
class ComparisonSide:
    """One side of a :attr:`SlideType.COMPARISON`."""

    title: str
    points: List[Bullet] = field(default_factory=list)
    subtitle: str = ""
    #: Renders this side as the recommended option.
    preferred: bool = False

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {"title": self.title}
        if self.subtitle:
            data["subtitle"] = self.subtitle
        if self.points:
            data["points"] = [p.to_dict() for p in self.points]
        if self.preferred:
            data["preferred"] = True
        return data


@dataclass
class ChartSeries:
    """One data series inside a chart."""

    name: str
    values: List[float]


@dataclass
class ChartData:
    """Native chart definition rendered as a real (editable) chart."""

    chart_type: str = "bar"   # bar, column, line, area, pie, doughnut, stacked_*
    categories: List[str] = field(default_factory=list)
    series: List[ChartSeries] = field(default_factory=list)
    unit: Optional[str] = None
    #: Print the value on each data point. Best left off for dense series.
    data_labels: bool = False
    #: A short sentence stating what the chart *means*. Rendered beneath the
    #: plot. A chart without a takeaway makes the audience do the analysis.
    takeaway: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {
            "chart_type": self.chart_type,
            "categories": list(self.categories),
            "series": [asdict(s) for s in self.series],
        }
        if self.unit:
            data["unit"] = self.unit
        if self.data_labels:
            data["data_labels"] = True
        if self.takeaway:
            data["takeaway"] = self.takeaway
        return data


@dataclass
class TableData:
    """Native table definition."""

    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    #: Column indices to right-align (numeric columns should be right-aligned
    #: so digits line up by place value).
    numeric_columns: List[int] = field(default_factory=list)
    #: Row index to visually emphasise, e.g. a totals row.
    highlight_row: Optional[int] = None

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {
            "headers": list(self.headers),
            "rows": [list(r) for r in self.rows],
        }
        if self.numeric_columns:
            data["numeric_columns"] = list(self.numeric_columns)
        if self.highlight_row is not None:
            data["highlight_row"] = self.highlight_row
        return data


@dataclass
class Slide:
    """A single slide. Only the fields relevant to ``type`` need to be set."""

    type: SlideType
    title: str = ""
    subtitle: str = ""

    # Prose
    bullets: List[Bullet] = field(default_factory=list)
    columns: List[List[Bullet]] = field(default_factory=list)  # two_column
    quote: str = ""
    attribution: str = ""

    # Structured
    stats: List[Stat] = field(default_factory=list)
    cards: List[Card] = field(default_factory=list)
    timeline: List[TimelineItem] = field(default_factory=list)
    steps: List[ProcessStep] = field(default_factory=list)
    sides: List[ComparisonSide] = field(default_factory=list)
    verdict: str = ""          # conclusion line under a comparison

    # Data
    chart: Optional[ChartData] = None
    table: Optional[TableData] = None
    image_path: str = ""
    caption: str = ""

    # Meta
    notes: str = ""            # speaker notes
    #: A one-line "so what". Rendered as a kicker above the title, giving the
    #: audience the conclusion before the evidence.
    kicker: str = ""
    #: Force a theme accent index for this slide's highlight colour.
    accent: Optional[int] = None

    # -- serialisation ----------------------------------------------------

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {"type": self.type.value}
        if self.title:
            data["title"] = self.title
        if self.subtitle:
            data["subtitle"] = self.subtitle
        if self.kicker:
            data["kicker"] = self.kicker
        if self.bullets:
            data["bullets"] = [b.to_dict() for b in self.bullets]
        if self.columns:
            data["columns"] = [[b.to_dict() for b in col] for col in self.columns]
        if self.stats:
            data["stats"] = [s.to_dict() for s in self.stats]
        if self.cards:
            data["cards"] = [c.to_dict() for c in self.cards]
        if self.timeline:
            data["timeline"] = [t.to_dict() for t in self.timeline]
        if self.steps:
            data["steps"] = [s.to_dict() for s in self.steps]
        if self.sides:
            data["sides"] = [s.to_dict() for s in self.sides]
        if self.verdict:
            data["verdict"] = self.verdict
        if self.chart:
            data["chart"] = self.chart.to_dict()
        if self.table:
            data["table"] = self.table.to_dict()
        if self.quote:
            data["quote"] = self.quote
        if self.attribution:
            data["attribution"] = self.attribution
        if self.image_path:
            data["image_path"] = self.image_path
        if self.caption:
            data["caption"] = self.caption
        if self.accent is not None:
            data["accent"] = self.accent
        if self.notes:
            data["notes"] = self.notes
        return data

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Slide":
        bullets = [_bullet_from(b) for b in data.get("bullets") or []]
        columns = [[_bullet_from(b) for b in col] for col in data.get("columns") or []]
        stats = [_stat_from(s) for s in data.get("stats") or []]
        cards = [_card_from(c) for c in data.get("cards") or []]
        timeline = [_timeline_from(t) for t in data.get("timeline") or []]
        steps = [_step_from(s) for s in data.get("steps") or []]
        sides = [_side_from(s) for s in data.get("sides") or []]

        chart = None
        if data.get("chart"):
            cd = data["chart"]
            chart = ChartData(
                chart_type=cd.get("chart_type", "bar"),
                categories=[str(c) for c in cd.get("categories") or []],
                series=[
                    ChartSeries(
                        name=str(s.get("name", f"Series {i + 1}")),
                        values=[_as_float(v) for v in s.get("values") or []],
                    )
                    for i, s in enumerate(cd.get("series") or [])
                ],
                unit=cd.get("unit"),
                data_labels=bool(cd.get("data_labels", False)),
                takeaway=cd.get("takeaway", "") or "",
            )

        table = None
        if data.get("table"):
            td = data["table"]
            table = TableData(
                headers=[str(h) for h in td.get("headers") or []],
                rows=[[_cell(v) for v in r] for r in td.get("rows") or []],
                numeric_columns=[int(i) for i in td.get("numeric_columns") or []],
                highlight_row=td.get("highlight_row"),
            )

        return cls(
            type=_slide_type_from(data.get("type")),
            title=data.get("title", "") or "",
            subtitle=data.get("subtitle", "") or "",
            kicker=data.get("kicker", "") or "",
            bullets=bullets,
            columns=columns,
            stats=stats,
            cards=cards,
            timeline=timeline,
            steps=steps,
            sides=sides,
            verdict=data.get("verdict", "") or "",
            chart=chart,
            table=table,
            quote=data.get("quote", "") or "",
            attribution=data.get("attribution", "") or "",
            image_path=data.get("image_path", "") or "",
            caption=data.get("caption", "") or "",
            accent=data.get("accent"),
            notes=data.get("notes", "") or "",
        )


# --- coercion helpers ----------------------------------------------------
#
# These are forgiving on purpose. The LLM path produces JSON that is usually
# but not reliably schema-perfect, and a hand-edited spec should not explode on
# a stray integer where a string was expected. Every helper degrades to
# something renderable rather than raising.


def _slide_type_from(value: Any) -> SlideType:
    """Coerce a slide type, tolerating unknown values and aliases."""
    if isinstance(value, SlideType):
        return value
    key = str(value or "").strip().lower().replace("-", "_").replace(" ", "_")
    try:
        return SlideType(key)
    except ValueError:
        return _TYPE_ALIASES.get(key, SlideType.BULLETS)


#: Names a model might plausibly emit, mapped to the real archetype.
_TYPE_ALIASES: Dict[str, SlideType] = {
    "title": SlideType.COVER,
    "title_slide": SlideType.COVER,
    "opening": SlideType.COVER,
    "toc": SlideType.AGENDA,
    "contents": SlideType.AGENDA,
    "outline": SlideType.AGENDA,
    "divider": SlideType.SECTION,
    "section_divider": SlideType.SECTION,
    "bullet": SlideType.BULLETS,
    "bullet_points": SlideType.BULLETS,
    "list": SlideType.BULLETS,
    "content": SlideType.BULLETS,
    "text": SlideType.BULLETS,
    "kpi": SlideType.STATS,
    "kpis": SlideType.STATS,
    "metrics": SlideType.STATS,
    "numbers": SlideType.STATS,
    "statement": SlideType.BIG_STATEMENT,
    "hero": SlideType.BIG_STATEMENT,
    "callout": SlideType.BIG_STATEMENT,
    "feature_cards": SlideType.CARDS,
    "grid": SlideType.CARDS,
    "features": SlideType.CARDS,
    "roadmap": SlideType.TIMELINE,
    "milestones": SlideType.TIMELINE,
    "steps": SlideType.PROCESS,
    "workflow": SlideType.PROCESS,
    "flow": SlideType.PROCESS,
    "versus": SlideType.COMPARISON,
    "vs": SlideType.COMPARISON,
    "pros_cons": SlideType.COMPARISON,
    "graph": SlideType.CHART,
    "picture": SlideType.IMAGE,
    "photo": SlideType.IMAGE,
    "thanks": SlideType.CLOSING,
    "thank_you": SlideType.CLOSING,
    "end": SlideType.CLOSING,
    "cta": SlideType.CLOSING,
}


def _as_float(value: Any) -> float:
    """Coerce a chart value to float, tolerating strings like ``"1,240"``."""
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").replace("%", "").strip() or 0)
    except (TypeError, ValueError):
        return 0.0


def _cell(value: Any) -> str:
    return "" if value is None else str(value)


def _bullet_from(b: Any) -> Bullet:
    """Accept a plain string, a dict, or a nested list of either."""
    if isinstance(b, str):
        return Bullet(text=b)
    if isinstance(b, dict):
        return Bullet(
            text=str(b.get("text", b.get("title", ""))),
            level=int(b.get("level", 0) or 0),
            bold_lead=b.get("bold_lead") or b.get("lead") or None,
        )
    return Bullet(text=_cell(b))


def _stat_from(s: Any) -> Stat:
    if isinstance(s, dict):
        return Stat(
            value=_cell(s.get("value", "")),
            label=_cell(s.get("label", "")),
            detail=s.get("detail") or None,
            trend=(str(s["trend"]).lower() if s.get("trend") else None),
        )
    return Stat(value=_cell(s), label="")


def _card_from(c: Any) -> Card:
    if isinstance(c, str):
        return Card(title=c)
    if isinstance(c, dict):
        return Card(
            title=_cell(c.get("title", "")),
            body=_cell(c.get("body", c.get("detail", "")) or ""),
            icon=_cell(c.get("icon", "") or ""),
            accent=c.get("accent"),
        )
    return Card(title=_cell(c))


def _timeline_from(t: Any) -> TimelineItem:
    if isinstance(t, str):
        return TimelineItem(period="", title=t)
    if isinstance(t, dict):
        return TimelineItem(
            period=_cell(t.get("period", t.get("date", t.get("when", ""))) or ""),
            title=_cell(t.get("title", t.get("label", "")) or ""),
            detail=_cell(t.get("detail", t.get("description", "")) or ""),
            done=bool(t.get("done", False)),
        )
    return TimelineItem(period="", title=_cell(t))


def _step_from(s: Any) -> ProcessStep:
    if isinstance(s, str):
        return ProcessStep(title=s)
    if isinstance(s, dict):
        return ProcessStep(
            title=_cell(s.get("title", s.get("label", "")) or ""),
            detail=_cell(s.get("detail", s.get("description", "")) or ""),
        )
    return ProcessStep(title=_cell(s))


def _side_from(s: Any) -> ComparisonSide:
    if isinstance(s, dict):
        raw = s.get("points", s.get("bullets", [])) or []
        return ComparisonSide(
            title=_cell(s.get("title", "") or ""),
            subtitle=_cell(s.get("subtitle", "") or ""),
            points=[_bullet_from(p) for p in raw],
            preferred=bool(s.get("preferred", False)),
        )
    return ComparisonSide(title=_cell(s))


# Retained under its original name: earlier releases exposed this helper.
_bullet_from_dict = _bullet_from


@dataclass
class Deck:
    """A full presentation."""

    title: str
    subtitle: str = ""
    author: str = ""
    theme: str = "kimi"
    slides: List[Slide] = field(default_factory=list)
    #: Layout density: ``airy``, ``normal`` or ``dense``. Scales the type
    #: ramp so a content-heavy deck stays inside its margins.
    density: str = "normal"
    #: ``"16:9"`` (default) or ``"4:3"``.
    aspect: str = "16:9"
    #: Shown in the footer beside the page number, e.g. "Confidential".
    footer_note: str = ""

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {
            "title": self.title,
            "subtitle": self.subtitle,
            "author": self.author,
            "theme": self.theme,
        }
        if self.density != "normal":
            data["density"] = self.density
        if self.aspect != "16:9":
            data["aspect"] = self.aspect
        if self.footer_note:
            data["footer_note"] = self.footer_note
        data["slides"] = [s.to_dict() for s in self.slides]
        return data

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Deck":
        return cls(
            title=data.get("title", "Untitled Deck") or "Untitled Deck",
            subtitle=data.get("subtitle", "") or "",
            author=data.get("author", "") or "",
            theme=data.get("theme", "kimi") or "kimi",
            density=data.get("density", "normal") or "normal",
            aspect=data.get("aspect", "16:9") or "16:9",
            footer_note=data.get("footer_note", "") or "",
            slides=[Slide.from_dict(s) for s in data.get("slides") or []],
        )

    # -- convenience ------------------------------------------------------

    def slide_count(self) -> int:
        return len(self.slides)

    def type_histogram(self) -> Dict[str, int]:
        """Count slides by archetype.

        Used by :meth:`variety_score` and by the CLI's ``--report`` output.
        """
        hist: Dict[str, int] = {}
        for s in self.slides:
            hist[s.type.value] = hist.get(s.type.value, 0) + 1
        return hist

    def variety_score(self) -> float:
        """Fraction of *content* slides that are not plain bullet lists.

        A blunt but honest proxy for whether a deck is designed or merely
        transcribed. Decks that lean entirely on ``bullets`` score near 0.
        """
        content = [
            s for s in self.slides
            if s.type not in (SlideType.COVER, SlideType.AGENDA, SlideType.SECTION, SlideType.CLOSING)
        ]
        if not content:
            return 0.0
        non_bullet = sum(1 for s in content if s.type != SlideType.BULLETS)
        return non_bullet / len(content)


__all__ = [
    "SlideType",
    "FULL_BLEED_TYPES",
    "Bullet",
    "Stat",
    "Card",
    "TimelineItem",
    "ProcessStep",
    "ComparisonSide",
    "ChartSeries",
    "ChartData",
    "TableData",
    "Slide",
    "Deck",
]
