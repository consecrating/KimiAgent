"""Data model for a KimiAgent deck.

A ``Deck`` is a serialisable, tool-agnostic description of a presentation — the
"reason first, design later" intermediate representation. The agent produces a
``Deck``; the renderer turns it into a native ``.pptx``.

Everything here is plain dataclasses so a deck round-trips cleanly to and from
JSON, which makes decks easy to inspect, hand-edit, diff and version-control.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from enum import Enum
from typing import Any, Dict, List, Optional


class SlideType(str, Enum):
    """The supported slide archetypes.

    Each archetype maps to a dedicated layout routine in the renderer.
    """

    COVER = "cover"            # Title slide / deck opener
    AGENDA = "agenda"          # Table of contents
    SECTION = "section"        # Section divider
    BULLETS = "bullets"        # Title + bullet points
    TWO_COLUMN = "two_column"  # Two side-by-side bullet columns
    STATS = "stats"            # Big-number KPI / metric callouts
    CHART = "chart"            # Native, editable chart
    TABLE = "table"            # Native table
    QUOTE = "quote"            # Pull quote
    IMAGE = "image"            # Full-bleed / captioned image
    CLOSING = "closing"        # Thank-you / call to action


@dataclass
class Bullet:
    """A single bullet point, optionally nested and optionally with a lead-in."""

    text: str
    level: int = 0            # indentation level (0 = top level)
    bold_lead: Optional[str] = None  # optional bold prefix, e.g. "Key point:"

    def to_dict(self) -> Dict[str, Any]:
        return {k: v for k, v in asdict(self).items() if v not in (None, 0)}


@dataclass
class Stat:
    """A large headline metric with a short label and optional detail."""

    value: str                # e.g. "42%", "$3.1B", "1.2M"
    label: str                # e.g. "YoY growth"
    detail: Optional[str] = None  # small supporting line


@dataclass
class ChartSeries:
    """One data series inside a chart."""

    name: str
    values: List[float]


@dataclass
class ChartData:
    """Native chart definition rendered as a real (editable) PowerPoint chart."""

    chart_type: str = "bar"   # one of: bar, column, line, pie, area
    categories: List[str] = field(default_factory=list)
    series: List[ChartSeries] = field(default_factory=list)
    unit: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "chart_type": self.chart_type,
            "categories": list(self.categories),
            "series": [asdict(s) for s in self.series],
            "unit": self.unit,
        }


@dataclass
class TableData:
    """Native table definition."""

    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)


@dataclass
class Slide:
    """A single slide. Only the fields relevant to ``type`` need to be set."""

    type: SlideType
    title: str = ""
    subtitle: str = ""
    bullets: List[Bullet] = field(default_factory=list)
    columns: List[List[Bullet]] = field(default_factory=list)  # for two_column
    stats: List[Stat] = field(default_factory=list)
    chart: Optional[ChartData] = None
    table: Optional[TableData] = None
    quote: str = ""
    attribution: str = ""
    image_path: str = ""
    caption: str = ""
    notes: str = ""  # speaker notes

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {"type": self.type.value}
        if self.title:
            data["title"] = self.title
        if self.subtitle:
            data["subtitle"] = self.subtitle
        if self.bullets:
            data["bullets"] = [b.to_dict() for b in self.bullets]
        if self.columns:
            data["columns"] = [[b.to_dict() for b in col] for col in self.columns]
        if self.stats:
            data["stats"] = [asdict(s) for s in self.stats]
        if self.chart:
            data["chart"] = self.chart.to_dict()
        if self.table:
            data["table"] = asdict(self.table)
        if self.quote:
            data["quote"] = self.quote
        if self.attribution:
            data["attribution"] = self.attribution
        if self.image_path:
            data["image_path"] = self.image_path
        if self.caption:
            data["caption"] = self.caption
        if self.notes:
            data["notes"] = self.notes
        return data

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Slide":
        bullets = [_bullet_from_dict(b) for b in data.get("bullets", [])]
        columns = [[_bullet_from_dict(b) for b in col] for col in data.get("columns", [])]
        stats = [Stat(**s) for s in data.get("stats", [])]

        chart = None
        if data.get("chart"):
            cd = data["chart"]
            chart = ChartData(
                chart_type=cd.get("chart_type", "bar"),
                categories=list(cd.get("categories", [])),
                series=[ChartSeries(**s) for s in cd.get("series", [])],
                unit=cd.get("unit"),
            )

        table = None
        if data.get("table"):
            td = data["table"]
            table = TableData(headers=list(td.get("headers", [])), rows=[list(r) for r in td.get("rows", [])])

        return cls(
            type=SlideType(data["type"]),
            title=data.get("title", ""),
            subtitle=data.get("subtitle", ""),
            bullets=bullets,
            columns=columns,
            stats=stats,
            chart=chart,
            table=table,
            quote=data.get("quote", ""),
            attribution=data.get("attribution", ""),
            image_path=data.get("image_path", ""),
            caption=data.get("caption", ""),
            notes=data.get("notes", ""),
        )


def _bullet_from_dict(b: Any) -> Bullet:
    """Accept either a plain string or a full bullet dict."""
    if isinstance(b, str):
        return Bullet(text=b)
    return Bullet(text=b.get("text", ""), level=b.get("level", 0), bold_lead=b.get("bold_lead"))


@dataclass
class Deck:
    """A full presentation."""

    title: str
    subtitle: str = ""
    author: str = ""
    theme: str = "kimi"
    slides: List[Slide] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "title": self.title,
            "subtitle": self.subtitle,
            "author": self.author,
            "theme": self.theme,
            "slides": [s.to_dict() for s in self.slides],
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Deck":
        return cls(
            title=data.get("title", "Untitled Deck"),
            subtitle=data.get("subtitle", ""),
            author=data.get("author", ""),
            theme=data.get("theme", "kimi"),
            slides=[Slide.from_dict(s) for s in data.get("slides", [])],
        )
