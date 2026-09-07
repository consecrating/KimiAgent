"""Visual themes for KimiAgent decks.

Each theme is a small palette + typography spec. The renderer reads these to
keep every slide visually consistent. Palettes lean on muted, professional
colour systems (including Morandi-inspired tones, a nod to Kimi's design
sensibility) rather than the harsh primaries of default templates.

Colours are plain ``RRGGBB`` hex strings (no leading ``#``) so they can be fed
straight to ``pptx.dml.color.RGBColor.from_string``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List


@dataclass
class Theme:
    """A colour + type system for a deck."""

    name: str
    # Core colours
    primary: str        # dominant brand colour (titles, accents, section bg)
    secondary: str      # supporting accent
    background: str     # normal slide background
    surface: str        # cards / panels
    text: str           # body text on light backgrounds
    text_muted: str     # captions, footers
    text_on_primary: str  # text drawn on top of the primary colour
    # Sequence used for chart series / stat callouts, cycled as needed
    accents: List[str] = field(default_factory=list)
    # Typography
    heading_font: str = "Calibri"
    body_font: str = "Calibri"

    def accent(self, index: int) -> str:
        """Return an accent colour, cycling through the palette."""
        if not self.accents:
            return self.primary
        return self.accents[index % len(self.accents)]


# --- Palette definitions -------------------------------------------------

THEMES: Dict[str, Theme] = {
    # Signature theme: deep ink + warm coral accent, calm and modern.
    "kimi": Theme(
        name="kimi",
        primary="1B2A4A",       # deep navy ink
        secondary="E8623B",     # coral
        background="FFFFFF",
        surface="F4F6FA",
        text="1F2733",
        text_muted="6B7688",
        text_on_primary="FFFFFF",
        accents=["E8623B", "3B82C4", "3FA796", "E0A458", "8E7CC3"],
        heading_font="Calibri",
        body_font="Calibri",
    ),
    # Morandi: muted, dusty, gallery-like tones.
    "morandi": Theme(
        name="morandi",
        primary="6E6A63",       # warm grey
        secondary="A8998A",     # taupe
        background="F5F2EC",
        surface="EAE4DA",
        text="4A453E",
        text_muted="8B8477",
        text_on_primary="F7F4EE",
        accents=["A3B0A0", "C3A9A0", "9FA8B0", "C8B79B", "8E9E9C"],
        heading_font="Georgia",
        body_font="Calibri",
    ),
    # Ink-wash: monochrome, editorial, high contrast.
    "inkwash": Theme(
        name="inkwash",
        primary="222222",
        secondary="7A7A7A",
        background="FBFBF9",
        surface="EEEEEC",
        text="1A1A1A",
        text_muted="777777",
        text_on_primary="FBFBF9",
        accents=["444444", "888888", "B0B0B0", "5C5C5C", "9A9A9A"],
        heading_font="Georgia",
        body_font="Calibri",
    ),
    # Corporate blue: trustworthy, conservative, boardroom-ready.
    "corporate": Theme(
        name="corporate",
        primary="0F4C81",
        secondary="2E9CCA",
        background="FFFFFF",
        surface="EEF4FA",
        text="1C2A36",
        text_muted="5B6B7A",
        text_on_primary="FFFFFF",
        accents=["2E9CCA", "0F4C81", "5BC0BE", "F2A65A", "6C7A89"],
        heading_font="Calibri",
        body_font="Calibri",
    ),
    # Forest: sustainability / energy / nature decks.
    "forest": Theme(
        name="forest",
        primary="1E3B2F",
        secondary="4C9A6C",
        background="FBFDFB",
        surface="EAF2EC",
        text="1B2B22",
        text_muted="5E6F65",
        text_on_primary="FFFFFF",
        accents=["4C9A6C", "88C999", "2F6F52", "D9A566", "6FA8A0"],
        heading_font="Calibri",
        body_font="Calibri",
    ),
    # Midnight: dark background, vivid accents — great for product/tech.
    "midnight": Theme(
        name="midnight",
        primary="0B1220",
        secondary="4F8CFF",
        background="0E1626",
        surface="18233A",
        text="E6ECF5",
        text_muted="9AA7BD",
        text_on_primary="FFFFFF",
        accents=["4F8CFF", "38D9A9", "FF6B6B", "FFD166", "B692FF"],
        heading_font="Calibri",
        body_font="Calibri",
    ),
}

DEFAULT_THEME = "kimi"


def get_theme(name: str) -> Theme:
    """Look up a theme by name, falling back to the default if unknown."""
    return THEMES.get((name or "").lower(), THEMES[DEFAULT_THEME])


def theme_names() -> List[str]:
    return list(THEMES.keys())
