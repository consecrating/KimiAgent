"""KimiAgent — turn a topic or outline into a designer-quality deck and PDF.

The design philosophy is inspired by Kimi / Moonshot AI's document-generation
skills: reason the narrative first, then design; and emit *native* output (real
charts, tables and shapes) rather than flat text boxes or a filled-in template.

Typical usage::

    from kimiagent import DeckAgent, render_deck, render_pdf

    deck = DeckAgent().build(topic="The State of Renewable Energy in 2026")
    render_deck(deck, "renewables.pptx")
    render_pdf(deck, "renewables.pdf")

The package is layered so each piece can be used on its own:

``kimiagent.design``
    The design system: a 12-column grid, a modular type scale, dependency-free
    font metrics for measuring and fitting text, and WCAG colour maths.
``kimiagent.themes``
    Colour/typography/elevation token sets, including contrast-safe gradients.
``kimiagent.models``
    The ``Deck`` / ``Slide`` spec — plain dataclasses that round-trip to JSON.
``kimiagent.layout``
    Slide zones plus the pagination that stops content overflowing.
``kimiagent.agent``
    Content generation (offline generator + optional Kimi API path).
``kimiagent.renderer`` / ``kimiagent.pdf_renderer``
    The two backends, both driven by the same design system.
"""

from .agent import DeckAgent, critique
from .design import GRID, TYPE, Grid, TypeScale
from .layout import paginate
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
from .pdf_renderer import PDFRenderer, render_pdf, soffice_available
from .renderer import DeckRenderer, render_deck
from .themes import THEMES, Theme, get_theme, register_theme, theme_names

__all__ = [
    # spec
    "Deck",
    "Slide",
    "SlideType",
    "Bullet",
    "Stat",
    "Card",
    "TimelineItem",
    "ProcessStep",
    "ComparisonSide",
    "ChartData",
    "ChartSeries",
    "TableData",
    # generation
    "DeckAgent",
    "critique",
    # rendering
    "render_deck",
    "render_pdf",
    "DeckRenderer",
    "PDFRenderer",
    "soffice_available",
    # design system
    "THEMES",
    "Theme",
    "get_theme",
    "theme_names",
    "register_theme",
    "Grid",
    "TypeScale",
    "GRID",
    "TYPE",
    "paginate",
]

__version__ = "0.2.0"
