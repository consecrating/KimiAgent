"""Visual themes for KimiAgent decks.

A theme is a complete, self-consistent design token set: colour roles,
gradients, corner radius, elevation and a type pairing. The renderers read
*only* these tokens, never hardcoded colours — so adding a theme here
restyles every slide archetype in both the ``.pptx`` and ``.pdf`` backends
with no renderer changes.

Design notes
------------
The palettes follow the constraint that current presentation-design guidance
keeps converging on: **one neutral, one accent, one background**, with a small
supporting accent ramp reserved for data. Decks look generated when they use
six competing hues on one slide; they look designed when a single accent does
all the pointing.

Colour *roles* matter more than colour values. ``primary`` is structure,
``secondary`` is the one thing that draws the eye, ``surface`` is elevation,
and ``accents`` exists for charts. Because these are roles, a renderer can ask
for "the accent for series 3" without knowing anything about the theme.

Every colour is a plain ``RRGGBB`` hex string with no leading ``#``, so it
feeds straight into ``RGBColor.from_string`` (pptx) and ``HexColor`` (PDF).

Contrast is *computed*, not assumed: see :meth:`Theme.on` and
:meth:`Theme.ink_on`, which pick a legible foreground for any background using
WCAG relative luminance. That is what prevents the classic generated-deck bug
of white text on a pale amber accent.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Dict, List, Optional, Tuple

from .design import (
    contrast_ratio,
    darken,
    ensure_contrast,
    is_dark,
    lighten,
    mix,
    readable_text_on,
    tint_sequence,
)


@dataclass
class Theme:
    """A colour + type + elevation system for a deck.

    New tokens all carry defaults, so a theme (or a hand-written spec) that
    only sets the original fields still renders correctly.
    """

    name: str

    # -- Core colour roles ------------------------------------------------
    primary: str            # dominant structural colour (titles, hero fills)
    secondary: str          # the single attention-drawing accent
    background: str         # normal slide background
    surface: str            # cards / panels raised off the background
    text: str               # body text on ``background``
    text_muted: str         # captions, footers, axis labels
    text_on_primary: str    # text drawn on top of ``primary``

    #: Chart/data accent ramp, cycled as needed.
    accents: List[str] = field(default_factory=list)

    # -- Typography -------------------------------------------------------
    heading_font: str = "Segoe UI"
    body_font: str = "Segoe UI"
    #: True when the heading face is a serif, so text metrics widen slightly.
    heading_serif: bool = False
    body_serif: bool = False
    #: Extra weight cue for headings; PowerPoint has no variable weights, so
    #: the renderers emulate "light" headings by using a muted colour.
    heading_bold: bool = True

    # -- Shape language ---------------------------------------------------
    #: Corner radius in points. 0 = hard edges (editorial), 10-16 = modern
    #: product feel. This single number changes a deck's character more than
    #: any colour choice.
    radius: float = 12.0
    #: Whether cards get a soft outer shadow (elevation).
    shadow: bool = True
    #: Whether cards get a hairline border. Usually the alternative to shadow,
    #: not a companion to it.
    hairline: bool = False

    # -- Gradient ---------------------------------------------------------
    #: Hero gradient endpoints. When ``None`` the renderers derive a subtle
    #: gradient from ``primary``, so every theme gets a usable hero.
    gradient_from: Optional[str] = None
    gradient_to: Optional[str] = None
    #: Gradient direction in degrees. 0 = left-to-right, 90 = bottom-to-top.
    gradient_angle: float = 45.0

    # -- Derived / optional overrides -------------------------------------
    surface_alt: Optional[str] = None   # second elevation level (zebra rows)
    border: Optional[str] = None        # hairline colour
    chart_grid: Optional[str] = None    # gridline colour

    #: True for dark-background themes. Auto-detected when left as ``None``.
    dark: Optional[bool] = None

    #: One-line description, shown by ``kimiagent themes``.
    blurb: str = ""

    # -- Normalisation ----------------------------------------------------

    def __post_init__(self) -> None:
        if self.dark is None:
            self.dark = is_dark(self.background)
        if self.surface_alt is None:
            # A second elevation step, nudged further from the background.
            self.surface_alt = (
                lighten(self.surface, 0.06) if self.dark else darken(self.surface, 0.035)
            )
        if self.border is None:
            self.border = (
                lighten(self.surface, 0.14) if self.dark else darken(self.surface, 0.10)
            )
        if self.chart_grid is None:
            self.chart_grid = (
                lighten(self.background, 0.14) if self.dark else darken(self.background, 0.09)
            )
        if not self.accents:
            self.accents = [self.secondary]
        if self.gradient_from is None:
            self.gradient_from = self.primary
        if self.gradient_to is None:
            # A hero gradient that leans toward the accent reads as intentional;
            # a pure tint of one colour reads as a rendering artefact.
            self.gradient_to = mix(self.primary, self.secondary, 0.42)

    # -- Colour helpers ---------------------------------------------------

    def accent(self, index: int) -> str:
        """Return an accent colour, cycling through the ramp."""
        if not self.accents:
            return self.secondary
        return self.accents[index % len(self.accents)]

    def series_colors(self, count: int, monochrome: bool = False) -> List[str]:
        """Colours for ``count`` chart series.

        A single series gets a monochrome ramp of the accent rather than one
        flat colour per bar, which reads as considered rather than arbitrary.
        Multiple series get distinct accents so they stay tellable apart.
        """
        if monochrome or count == 1:
            return tint_sequence(self.secondary, max(count, 1))
        return [self.accent(i) for i in range(count)]

    def on(self, background: str) -> str:
        """Pick a legible text colour for any background fill."""
        return readable_text_on(background, light="FFFFFF", dark=self.ink)

    def ink_on(self, background: str, minimum: float = 4.5) -> str:
        """Body-text colour on ``background``, guaranteed to meet WCAG AA."""
        return ensure_contrast(self.text, background, minimum)

    def muted_on(self, background: str, minimum: float = 3.0) -> str:
        """Secondary-text colour on ``background`` with a relaxed threshold.

        Captions and footers are allowed 3:1 rather than 4.5:1 — they are
        supporting text, and forcing full contrast flattens the hierarchy the
        muted tone exists to create.
        """
        return ensure_contrast(self.text_muted, background, minimum)

    def accent_on(self, background: str, minimum: float = 3.0) -> str:
        """The accent colour, darkened/lightened until it is legible as *large* text.

        WCAG allows 3:1 for large text (>=24pt bold), which is exactly what
        the oversized KPI numbers and display type are. Several attractive
        accents — ambers, mints, sky blues — sit around 2.5:1 on a near-white
        background and turn to mush at any distance; this pulls them just far
        enough to read without abandoning the hue.
        """
        return ensure_contrast(self.secondary, background, minimum)

    @property
    def ink(self) -> str:
        """The darkest text colour, for use on light fills."""
        return self.text if not self.dark else "111418"

    @property
    def gradient(self) -> Tuple[str, str, float]:
        """``(from, to, angle)`` for the hero gradient, made text-safe.

        A gradient is only as legible as its *worst* point. Several appealing
        ramps (deep rust to bright amber, navy to sky blue) span so much
        luminance that no single text colour clears WCAG AA along the whole
        sweep: white works at the dark end and disappears at the light end.

        Rather than accept that, the lighter endpoint is pulled toward the
        darker one until one text colour is legible across both ends. The hue
        and the sense of motion survive; the illegibility does not. Doing this
        here means custom themes registered by users get the same guarantee.
        """
        start = self.gradient_from or self.primary
        end = self.gradient_to or self.primary
        text, start, end = self._safe_gradient(start, end)
        return start, end, self.gradient_angle

    def _safe_gradient(self, start: str, end: str, minimum: float = 4.5) -> Tuple[str, str, str]:
        """Return ``(text_colour, start, end)`` guaranteed legible together.

        Both possible text colours are tried, each with the ramp compressed
        just enough to clear ``minimum``. The winner is whichever *moved the
        palette least* — measured as RGB displacement from the author's
        original endpoints. That criterion matters: darkening a vivid amber
        toward its deep-rust partner keeps the gradient saturated, whereas
        lightening the rust toward the amber turns it into mud. Choosing by
        displacement picks the rich option automatically.
        """

        def worst(text: str, a: str, b: str) -> float:
            # Contrast varies monotonically between endpoints, so sampling the
            # ends plus the midpoint bounds the whole sweep.
            return min(
                contrast_ratio(text, a),
                contrast_ratio(text, b),
                contrast_ratio(text, mix(a, b, 0.5)),
            )

        def displacement(a: str, b: str) -> float:
            from .design import to_rgb

            o1, o2 = to_rgb(start), to_rgb(end)
            n1, n2 = to_rgb(a), to_rgb(b)
            return sum(abs(o1[i] - n1[i]) + abs(o2[i] - n2[i]) for i in range(3))

        def compress(text: str) -> Tuple[str, str]:
            """Nudge the failing endpoint until ``text`` is legible throughout."""
            toward_dark = text == "FFFFFF"
            a, b = start, end
            for _ in range(28):
                if worst(text, a, b) >= minimum:
                    break
                if contrast_ratio(text, a) < contrast_ratio(text, b):
                    a = darken(a, 0.07) if toward_dark else lighten(a, 0.07)
                else:
                    b = darken(b, 0.07) if toward_dark else lighten(b, 0.07)
            return a, b

        options = []
        for text in ("FFFFFF", self.ink):
            a, b = compress(text)
            options.append((worst(text, a, b) >= minimum, displacement(a, b), text, a, b))

        # Prefer options that actually reach the threshold, then least distortion.
        options.sort(key=lambda o: (not o[0], o[1]))
        _, _, text, a, b = options[0]
        return text, a, b

    @property
    def hero_text(self) -> str:
        """Text colour for the hero gradient, chosen by measured contrast."""
        start = self.gradient_from or self.primary
        end = self.gradient_to or self.primary
        return self._safe_gradient(start, end)[0]

    def variant(self, **overrides) -> "Theme":
        """Return a copy of this theme with tokens overridden.

        Lets a caller tweak one token (``theme.variant(radius=0)``) without
        redefining a palette.
        """
        return replace(self, **overrides)


# --------------------------------------------------------------------------
# Palettes
# --------------------------------------------------------------------------
#
# Font pairings are chosen from faces that actually ship with Windows, macOS
# and Office, because a beautiful pairing that falls back to Times New Roman
# on the audience's laptop is worse than a plain one that survives. "Segoe UI"
# resolves on Windows/Office, degrades to Helvetica on macOS and Arial
# elsewhere — all metric-compatible with the embedded width tables.

THEMES: Dict[str, Theme] = {
    # ---------------- Signature ----------------
    "kimi": Theme(
        name="kimi",
        primary="16243D",
        secondary="F0643B",
        background="FFFFFF",
        surface="F5F7FA",
        text="1C2432",
        text_muted="69758A",
        text_on_primary="FFFFFF",
        accents=["F0643B", "3B82C4", "3FA796", "E0A458", "8E7CC3"],
        heading_font="Segoe UI",
        body_font="Segoe UI",
        radius=12,
        shadow=True,
        gradient_from="16243D",
        gradient_to="2C4A6E",
        gradient_angle=38,
        blurb="Deep navy ink + warm coral. Calm, modern, the default.",
    ),
    # ---------------- New: contemporary ----------------
    "aurora": Theme(
        name="aurora",
        primary="1A1038",
        secondary="8B5CF6",
        background="0F0A24",
        surface="1C1442",
        text="ECE9F7",
        text_muted="A79FC8",
        text_on_primary="FFFFFF",
        accents=["8B5CF6", "38BDF8", "F472B6", "34D399", "FBBF24"],
        heading_font="Segoe UI",
        body_font="Segoe UI",
        radius=16,
        shadow=True,
        gradient_from="2E1065",
        gradient_to="0EA5E9",
        gradient_angle=35,
        blurb="Indigo-to-cyan aurora on deep violet. Dark, vivid, keynote-ready.",
    ),
    "slate": Theme(
        name="slate",
        primary="1E293B",
        secondary="2563EB",
        background="FFFFFF",
        surface="F1F5F9",
        text="0F172A",
        text_muted="64748B",
        text_on_primary="FFFFFF",
        accents=["2563EB", "0EA5E9", "6366F1", "14B8A6", "F59E0B"],
        heading_font="Segoe UI",
        body_font="Segoe UI",
        radius=10,
        shadow=False,
        hairline=True,
        gradient_from="1E293B",
        gradient_to="2563EB",
        gradient_angle=40,
        blurb="Neutral slate + electric blue, hairline cards. Clean SaaS/product.",
    ),
    "sunset": Theme(
        name="sunset",
        primary="7C2D12",
        secondary="F97316",
        background="FFFBF7",
        surface="FFF1E6",
        text="2A1810",
        text_muted="8A6A55",
        text_on_primary="FFFFFF",
        accents=["F97316", "E11D48", "F59E0B", "A855F7", "0D9488"],
        heading_font="Segoe UI",
        body_font="Segoe UI",
        radius=18,
        shadow=True,
        gradient_from="9A3412",
        gradient_to="F59E0B",
        gradient_angle=30,
        blurb="Warm amber-to-coral gradient on cream. Energetic, human.",
    ),
    "mint": Theme(
        name="mint",
        primary="0F3D3E",
        secondary="10B981",
        background="FBFEFC",
        surface="ECFDF5",
        text="10241F",
        text_muted="5B7A72",
        text_on_primary="FFFFFF",
        accents=["10B981", "0891B2", "84CC16", "F59E0B", "6366F1"],
        heading_font="Segoe UI",
        body_font="Segoe UI",
        radius=14,
        shadow=True,
        gradient_from="0F3D3E",
        gradient_to="10B981",
        gradient_angle=42,
        blurb="Fresh teal + mint. Health, climate, wellbeing.",
    ),
    "noir": Theme(
        name="noir",
        primary="0A0A0A",
        secondary="D4A537",
        background="121212",
        surface="1E1E1E",
        text="F2F0EC",
        text_muted="9C968C",
        text_on_primary="FFFFFF",
        accents=["D4A537", "C2703E", "8E9B8A", "B9A2C4", "6E97A8"],
        heading_font="Georgia",
        body_font="Segoe UI",
        heading_serif=True,
        radius=0,
        shadow=False,
        hairline=True,
        gradient_from="0A0A0A",
        gradient_to="2A2118",
        gradient_angle=50,
        blurb="Near-black + antique gold, serif headings. Luxury, editorial.",
    ),
    "editorial": Theme(
        name="editorial",
        primary="1A1A1A",
        secondary="C2410C",
        background="FAF7F2",
        surface="F0EAE0",
        text="1A1A1A",
        text_muted="6B6255",
        text_on_primary="FAF7F2",
        accents=["C2410C", "1A1A1A", "78716C", "A16207", "44403C"],
        heading_font="Georgia",
        body_font="Georgia",
        heading_serif=True,
        body_serif=True,
        radius=0,
        shadow=False,
        hairline=True,
        gradient_from="1A1A1A",
        gradient_to="44403C",
        gradient_angle=90,
        blurb="Magazine serif on warm paper, hard edges. Long-form, essayistic.",
    ),
    "glacier": Theme(
        name="glacier",
        primary="0C4A6E",
        secondary="0EA5E9",
        background="F8FBFE",
        surface="E8F2FA",
        text="0B2637",
        text_muted="5A7689",
        text_on_primary="FFFFFF",
        accents=["0EA5E9", "0C4A6E", "38BDF8", "6366F1", "94A3B8"],
        heading_font="Segoe UI",
        body_font="Segoe UI",
        radius=14,
        shadow=True,
        gradient_from="0C4A6E",
        gradient_to="38BDF8",
        gradient_angle=35,
        blurb="Icy blues on near-white. Airy, analytical, data-heavy decks.",
    ),
    # ---------------- Original palettes, retuned ----------------
    "morandi": Theme(
        name="morandi",
        primary="6E6A63",
        secondary="A8998A",
        background="F5F2EC",
        surface="EAE4DA",
        text="413C35",
        text_muted="857E71",
        text_on_primary="F7F4EE",
        accents=["A3B0A0", "C3A9A0", "9FA8B0", "C8B79B", "8E9E9C"],
        heading_font="Georgia",
        body_font="Segoe UI",
        heading_serif=True,
        radius=8,
        shadow=False,
        hairline=True,
        gradient_from="6E6A63",
        gradient_to="A8998A",
        gradient_angle=60,
        blurb="Muted, dusty, gallery-like tones.",
    ),
    "inkwash": Theme(
        name="inkwash",
        primary="222222",
        secondary="6E6E6E",
        background="FBFBF9",
        surface="EEEEEC",
        text="1A1A1A",
        text_muted="6E6E6E",
        text_on_primary="FBFBF9",
        accents=["444444", "888888", "B0B0B0", "5C5C5C", "9A9A9A"],
        heading_font="Georgia",
        body_font="Segoe UI",
        heading_serif=True,
        radius=2,
        shadow=False,
        hairline=True,
        gradient_from="222222",
        gradient_to="555555",
        gradient_angle=90,
        blurb="Monochrome, editorial, high contrast.",
    ),
    "corporate": Theme(
        name="corporate",
        primary="0F4C81",
        secondary="2E9CCA",
        background="FFFFFF",
        surface="EEF4FA",
        text="16232E",
        text_muted="566A7A",
        text_on_primary="FFFFFF",
        accents=["2E9CCA", "0F4C81", "5BC0BE", "E08A3C", "6C7A89"],
        heading_font="Segoe UI",
        body_font="Segoe UI",
        radius=6,
        shadow=False,
        hairline=True,
        gradient_from="0F4C81",
        gradient_to="2E9CCA",
        gradient_angle=40,
        blurb="Trustworthy boardroom blue.",
    ),
    "forest": Theme(
        name="forest",
        primary="1E3B2F",
        secondary="4C9A6C",
        background="FBFDFB",
        surface="EAF2EC",
        text="16261E",
        text_muted="5A6B61",
        text_on_primary="FFFFFF",
        accents=["4C9A6C", "2F6F52", "88C999", "C08A3E", "5F938C"],
        heading_font="Segoe UI",
        body_font="Segoe UI",
        radius=12,
        shadow=True,
        gradient_from="1E3B2F",
        gradient_to="4C9A6C",
        gradient_angle=42,
        blurb="Greens for sustainability, energy, nature.",
    ),
    "midnight": Theme(
        name="midnight",
        primary="0B1220",
        secondary="4F8CFF",
        background="0E1626",
        surface="18233A",
        text="E6ECF5",
        text_muted="94A3BD",
        text_on_primary="FFFFFF",
        accents=["4F8CFF", "38D9A9", "FF6B6B", "FFD166", "B692FF"],
        heading_font="Segoe UI",
        body_font="Segoe UI",
        radius=12,
        shadow=True,
        gradient_from="0B1220",
        gradient_to="1E3A8A",
        gradient_angle=40,
        blurb="Dark canvas, vivid accents. Tech and product.",
    ),
}

DEFAULT_THEME = "kimi"

#: Themes grouped for the CLI listing, so related looks sit together.
THEME_GROUPS: Dict[str, List[str]] = {
    "signature": ["kimi"],
    "light": ["slate", "glacier", "mint", "sunset", "editorial", "corporate", "forest", "morandi", "inkwash"],
    "dark": ["aurora", "midnight", "noir"],
}


def get_theme(name: str) -> Theme:
    """Look up a theme by name, falling back to the default if unknown."""
    return THEMES.get((name or "").strip().lower(), THEMES[DEFAULT_THEME])


def theme_names() -> List[str]:
    return list(THEMES.keys())


def dark_themes() -> List[str]:
    return [n for n, t in THEMES.items() if t.dark]


def register_theme(theme: Theme) -> Theme:
    """Add a custom theme to the registry.

    Lets a user define a brand palette in a few lines and then use it by name
    from the CLI or the spec, without forking the package::

        from kimiagent.themes import Theme, register_theme
        register_theme(Theme(name="acme", primary="12203A", secondary="FF5A1F",
                             background="FFFFFF", surface="F4F6F9",
                             text="111827", text_muted="6B7280",
                             text_on_primary="FFFFFF"))
    """
    THEMES[theme.name.lower()] = theme
    return theme


__all__ = [
    "Theme",
    "THEMES",
    "THEME_GROUPS",
    "DEFAULT_THEME",
    "get_theme",
    "theme_names",
    "dark_themes",
    "register_theme",
]
