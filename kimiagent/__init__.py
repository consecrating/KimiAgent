"""KimiAgent — turn a topic or outline into a native, designer-quality PowerPoint deck.

The design philosophy is inspired by Kimi / Moonshot AI's document-generation
skills: reason the narrative first, then design; and emit *native* PowerPoint
(real charts, tables and shapes) rather than flat text boxes or a filled-in
template.

Typical usage::

    from kimiagent import DeckAgent, render_deck

    deck = DeckAgent().build(topic="The State of Renewable Energy in 2026")
    render_deck(deck, "renewables.pptx")
"""

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
from .agent import DeckAgent
from .renderer import render_deck
from .themes import THEMES, get_theme, Theme

__all__ = [
    "Deck",
    "Slide",
    "SlideType",
    "Bullet",
    "Stat",
    "ChartData",
    "ChartSeries",
    "TableData",
    "DeckAgent",
    "render_deck",
    "THEMES",
    "get_theme",
    "Theme",
]

__version__ = "0.1.0"
