"""The KimiAgent design system.

This module is the shared "taste layer" behind both renderers (``.pptx`` and
``.pdf``). It holds no PowerPoint or PDF specifics — only pure geometry,
typography and colour maths — so the two backends can lay out a slide
*identically* and stay visually in sync.

It exists because the single biggest reason generated decks look
machine-made is not the colour palette: it is **broken spatial rhythm**.
Text overflows its box, cards are sized by guesswork, headings collide with
body copy, and nothing lines up to a common grid. Everything here is aimed at
that problem.

Four things live here:

``Grid``
    A 12-column layout grid with a fixed margin and gutter, so every element
    on every slide snaps to the same vertical and horizontal rhythm.

``TypeScale``
    A modular type scale. Sizes are derived from a base by a fixed ratio
    rather than hand-picked, which is what makes hierarchy feel deliberate.

Text metrics
    :func:`wrap_text`, :func:`text_height` and :func:`fit_font_size` measure
    strings using an embedded Helvetica/Arial width table (values in 1/1000
    em, the standard AFM unit). This means text can be *fitted before it is
    drawn* with no font files, no rendering engine and no dependencies — so a
    long bullet shrinks to fit instead of spilling off the slide.

Colour
    sRGB-correct :func:`mix`, :func:`lighten`, :func:`darken` plus WCAG
    relative-luminance :func:`contrast_ratio` and :func:`readable_text_on`,
    so foreground colours are *computed* to be legible on whatever sits
    behind them instead of assumed.

Units: all geometry is in **points** (1/72 inch). Both backends convert at
the boundary — python-pptx wants EMU, reportlab already works in points.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, List, Sequence, Tuple

# --------------------------------------------------------------------------
# Units
# --------------------------------------------------------------------------

PT_PER_INCH = 72.0


def inches(value: float) -> float:
    """Inches -> points."""
    return value * PT_PER_INCH


#: Canonical 16:9 slide canvas, in points (13.333in x 7.5in).
SLIDE_W = inches(13.333)
SLIDE_H = inches(7.5)

#: 4:3 canvas, for decks that must play on legacy projectors.
SLIDE_W_4_3 = inches(10.0)
SLIDE_H_4_3 = inches(7.5)


# --------------------------------------------------------------------------
# Layout grid
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Grid:
    """A 12-column layout grid.

    Modern editorial and product design both lean on a column grid because it
    turns "where should this go?" into arithmetic. Every card, chart and
    column in the renderers is positioned with :meth:`col` / :meth:`span`, so
    elements on different slides line up with each other automatically.

    The default margins are deliberately generous. Research on slide design
    converges on leaving roughly 40% of a content slide (and more on a hero
    slide) as empty space; tight margins are the fastest way to make a deck
    feel cramped and anxious.
    """

    width: float = SLIDE_W
    height: float = SLIDE_H
    margin_x: float = inches(0.78)
    margin_top: float = inches(0.62)
    margin_bottom: float = inches(0.62)
    columns: int = 12
    gutter: float = inches(0.22)

    # -- horizontal ------------------------------------------------------

    @property
    def content_w(self) -> float:
        return self.width - 2 * self.margin_x

    @property
    def content_h(self) -> float:
        return self.height - self.margin_top - self.margin_bottom

    @property
    def col_w(self) -> float:
        """Width of a single column."""
        total_gutter = self.gutter * (self.columns - 1)
        return (self.content_w - total_gutter) / self.columns

    def col(self, index: int) -> float:
        """Left edge (x) of column ``index`` (0-based)."""
        return self.margin_x + index * (self.col_w + self.gutter)

    def span(self, count: int) -> float:
        """Width of ``count`` adjacent columns, gutters included."""
        count = max(1, count)
        return count * self.col_w + (count - 1) * self.gutter

    def cell(self, index: int, count: int) -> Tuple[float, float]:
        """``(left, width)`` for ``count`` columns starting at ``index``."""
        return self.col(index), self.span(count)

    # -- vertical --------------------------------------------------------

    @property
    def left(self) -> float:
        return self.margin_x

    @property
    def right(self) -> float:
        return self.width - self.margin_x

    @property
    def top(self) -> float:
        return self.margin_top

    @property
    def bottom(self) -> float:
        return self.height - self.margin_bottom

    def tracks(self, count: int, gap: float, total: float) -> List[Tuple[float, float]]:
        """Split ``total`` into ``count`` equal tracks separated by ``gap``.

        Returns a list of ``(offset, size)``. Used for card rows, stat rows
        and any other evenly distributed run of elements.
        """
        count = max(1, count)
        size = (total - gap * (count - 1)) / count
        return [(i * (size + gap), size) for i in range(count)]

    def row(self, count: int, gap: float | None = None) -> List[Tuple[float, float]]:
        """Evenly split the content width into ``count`` columns.

        Returns absolute ``(left, width)`` pairs, so this is the shortcut for
        laying out a row of N equal cards.
        """
        gap = self.gutter if gap is None else gap
        return [(self.margin_x + off, size) for off, size in self.tracks(count, gap, self.content_w)]


#: The default grid used unless a renderer is told otherwise.
GRID = Grid()


# --------------------------------------------------------------------------
# Typography
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class TypeScale:
    """A modular type scale, in points.

    Sizes step by a constant ratio rather than being chosen individually.
    That single constraint is most of what separates a deck that reads as
    "designed" from one that reads as "assembled": hierarchy becomes
    unambiguous because each level is visibly, proportionally larger than the
    one below it.

    ``display`` is reserved for hero moments (cover, big statement, KPI
    numbers) where oversized type is the design.
    """

    display: float = 66
    hero: float = 46
    h1: float = 32
    h2: float = 24
    h3: float = 19
    body: float = 17
    body_sm: float = 14.5
    caption: float = 12
    micro: float = 9.5

    #: Multiplied by font size to get line height (leading).
    line_height: float = 1.30
    #: Tighter leading for large display type, which needs less air.
    line_height_tight: float = 1.10

    def scaled(self, factor: float) -> "TypeScale":
        """Return a copy with every size multiplied by ``factor``.

        Used by the ``--density`` option: a denser deck needs slightly
        smaller type so more content fits without the layout breaking.
        """
        return TypeScale(
            display=self.display * factor,
            hero=self.hero * factor,
            h1=self.h1 * factor,
            h2=self.h2 * factor,
            h3=self.h3 * factor,
            body=self.body * factor,
            body_sm=self.body_sm * factor,
            caption=self.caption * factor,
            micro=self.micro * factor,
            line_height=self.line_height,
            line_height_tight=self.line_height_tight,
        )


TYPE = TypeScale()

#: Density presets. ``airy`` trades content for whitespace; ``dense`` does the
#: reverse. Exposed on the CLI as ``--density``.
DENSITY: Dict[str, float] = {"airy": 1.06, "normal": 1.0, "dense": 0.90}


# --------------------------------------------------------------------------
# Spacing scale
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Space:
    """A geometric spacing scale, in points.

    Padding and gaps are drawn from this fixed set so that whitespace is
    consistent across slide types instead of being improvised per layout.
    """

    xxs: float = 3
    xs: float = 6
    sm: float = 10
    md: float = 16
    lg: float = 24
    xl: float = 36
    xxl: float = 54


SPACE = Space()


# --------------------------------------------------------------------------
# Text metrics
# --------------------------------------------------------------------------

# Character advance widths in 1/1000 em for Helvetica (the metric-compatible
# stand-in for Arial, Inter, Noto Sans and the other UI sans-serifs these
# decks use). Covers ASCII 32..126. Extracted from the standard AFM tables.
_ASCII_START = 32

_W_REGULAR: Tuple[int, ...] = (
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278,
    278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584,
    584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556,
    833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278,
    278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222,
    500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500,
    500, 334, 260, 334, 584,
)

_W_BOLD: Tuple[int, ...] = (
    278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278,
    278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584,
    584, 611, 975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611,
    833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333,
    278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278,
    556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556,
    500, 389, 280, 389, 584,
)

# Fallback advance for anything outside ASCII. CJK and full-width punctuation
# are effectively one em wide; other non-ASCII gets a lowercase-ish average.
_W_FALLBACK = 556
_W_WIDE = 1000

# Serif faces (Georgia, Playfair) run a little wider than Helvetica at the
# same nominal size; a flat correction keeps fitting conservative.
_SERIF_FACTOR = 1.06


def _is_wide(ch: str) -> bool:
    """True for characters that occupy a full em (CJK, Hangul, kana, etc.)."""
    cp = ord(ch)
    return (
        0x1100 <= cp <= 0x115F
        or 0x2E80 <= cp <= 0xA4CF
        or 0xAC00 <= cp <= 0xD7A3
        or 0xF900 <= cp <= 0xFAFF
        or 0xFE30 <= cp <= 0xFE6F
        or 0xFF00 <= cp <= 0xFF60
        or 0xFFE0 <= cp <= 0xFFE6
        or 0x20000 <= cp <= 0x3FFFD
    )


def char_width(ch: str, size: float, bold: bool = False, serif: bool = False) -> float:
    """Advance width of a single character at ``size`` points."""
    cp = ord(ch)
    if _ASCII_START <= cp <= 126:
        table = _W_BOLD if bold else _W_REGULAR
        w = table[cp - _ASCII_START]
    elif _is_wide(ch):
        w = _W_WIDE
    else:
        w = _W_FALLBACK
    width = w / 1000.0 * size
    return width * _SERIF_FACTOR if serif else width


def text_width(text: str, size: float, bold: bool = False, serif: bool = False) -> float:
    """Width of ``text`` on one line, in points."""
    if not text:
        return 0.0
    table = _W_BOLD if bold else _W_REGULAR
    total = 0
    extra = 0.0
    for ch in text:
        cp = ord(ch)
        if _ASCII_START <= cp <= 126:
            total += table[cp - _ASCII_START]
        elif _is_wide(ch):
            total += _W_WIDE
        else:
            total += _W_FALLBACK
    width = total / 1000.0 * size + extra
    return width * _SERIF_FACTOR if serif else width


def wrap_text(
    text: str,
    max_width: float,
    size: float,
    bold: bool = False,
    serif: bool = False,
) -> List[str]:
    """Greedy word wrap of ``text`` into lines no wider than ``max_width``.

    Mirrors how PowerPoint and reportlab break lines (greedy, at spaces),
    which is what makes :func:`fit_font_size` predictive rather than
    hopeful. Explicit newlines are honoured. A single word longer than the
    line is split character-wise so it can never bleed out of its box.
    """
    if not text:
        return [""]
    if max_width <= 0:
        return [text]

    lines: List[str] = []
    for para in text.split("\n"):
        if not para.strip():
            lines.append("")
            continue
        words = para.split()
        current = ""
        for word in words:
            candidate = f"{current} {word}" if current else word
            if text_width(candidate, size, bold, serif) <= max_width:
                current = candidate
                continue
            if current:
                lines.append(current)
                current = ""
            # The word alone may still be too long: hard-break it.
            if text_width(word, size, bold, serif) <= max_width:
                current = word
            else:
                chunk = ""
                for ch in word:
                    if text_width(chunk + ch, size, bold, serif) <= max_width:
                        chunk += ch
                    else:
                        if chunk:
                            lines.append(chunk)
                        chunk = ch
                current = chunk
        if current:
            lines.append(current)
    return lines or [""]


def line_count(
    text: str,
    max_width: float,
    size: float,
    bold: bool = False,
    serif: bool = False,
) -> int:
    """Number of lines ``text`` wraps to at ``size``."""
    return len(wrap_text(text, max_width, size, bold, serif))


def text_height(
    text: str,
    max_width: float,
    size: float,
    bold: bool = False,
    serif: bool = False,
    line_height: float = TYPE.line_height,
) -> float:
    """Rendered height of wrapped ``text``, in points."""
    return line_count(text, max_width, size, bold, serif) * size * line_height


def longest_word_width(text: str, size: float, bold: bool = False, serif: bool = False) -> float:
    """Width of the widest single word in ``text``."""
    words = text.split()
    if not words:
        return 0.0
    return max(text_width(w, size, bold, serif) for w in words)


def fit_font_size(
    text: str,
    max_width: float,
    max_height: float,
    ideal: float,
    minimum: float = 9.0,
    bold: bool = False,
    serif: bool = False,
    line_height: float = TYPE.line_height,
    step: float = 0.5,
    avoid_word_break: bool = True,
) -> float:
    """Largest font size <= ``ideal`` at which ``text`` fits the box.

    This is the core anti-overflow primitive. python-pptx's own
    ``fit_text``/autofit needs a font file on disk and only supports a 12pt
    baseline, and PowerPoint's own autofit is computed by the *renderer* —
    meaning nothing is actually resized until a human opens the file. By
    measuring here we guarantee the geometry is correct in the saved
    document, in every viewer, and in the PDF.

    ``avoid_word_break`` additionally requires the *widest single word* to fit
    on one line. Without it, a narrow column produces breaks mid-word
    ("Interconnec / t") because the total height still fits — which looks far
    worse than simply setting the label a point or two smaller.

    Returns ``minimum`` if even that does not fit; callers that must not clip
    should pair this with :func:`truncate_to_fit`.
    """
    if not text:
        return ideal
    size = ideal
    while size > minimum:
        fits_height = text_height(text, max_width, size, bold, serif, line_height) <= max_height
        fits_words = (
            not avoid_word_break
            or longest_word_width(text, size, bold, serif) <= max_width
        )
        if fits_height and fits_words:
            return size
        size -= step
    return minimum


def truncate_to_width(
    text: str,
    max_width: float,
    size: float,
    bold: bool = False,
    serif: bool = False,
    ellipsis: str = "\u2026",
) -> str:
    """Trim ``text`` to a *single line* no wider than ``max_width``.

    For fixed one-line furniture — footers, badges, page labels — where
    wrapping to a second line would overrun a deliberately short box. Shrinking
    the type is wrong here because the size is doing hierarchical work.
    """
    if not text or text_width(text, size, bold, serif) <= max_width:
        return text
    if max_width <= 0:
        return ""
    ell = text_width(ellipsis, size, bold, serif)
    out = ""
    for ch in text:
        if text_width(out + ch, size, bold, serif) + ell > max_width:
            break
        out += ch
    return (out.rstrip() + ellipsis) if out.strip() else ellipsis


def truncate_to_fit(
    text: str,
    max_width: float,
    max_height: float,
    size: float,
    bold: bool = False,
    serif: bool = False,
    line_height: float = TYPE.line_height,
    ellipsis: str = "\u2026",
) -> str:
    """Trim ``text`` (on a word boundary) until it fits, adding an ellipsis.

    The last line of defence. Preferred behaviour is to shrink type or flow
    onto a continuation slide; this is for the cases where neither is
    possible, and it is still strictly better than silently overflowing.
    """
    if text_height(text, max_width, size, bold, serif, line_height) <= max_height:
        return text
    words = text.split()
    while words:
        words.pop()
        candidate = " ".join(words) + ellipsis
        if text_height(candidate, max_width, size, bold, serif, line_height) <= max_height:
            return candidate
    return ellipsis


def max_lines(max_height: float, size: float, line_height: float = TYPE.line_height) -> int:
    """How many lines of ``size`` type fit in ``max_height``."""
    if size <= 0:
        return 0
    return max(0, int(max_height // (size * line_height)))


# --------------------------------------------------------------------------
# Colour
# --------------------------------------------------------------------------


def _clamp(v: float, lo: float = 0.0, hi: float = 1.0) -> float:
    return max(lo, min(hi, v))


def to_rgb(hex_str: str) -> Tuple[int, int, int]:
    """``"1B2A4A"`` / ``"#1B2A4A"`` -> ``(27, 42, 74)``."""
    h = hex_str.lstrip("#").strip()
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    if len(h) != 6:
        raise ValueError(f"Not a 6-digit hex colour: {hex_str!r}")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def to_hex(rgb: Sequence[float]) -> str:
    """``(27, 42, 74)`` -> ``"1B2A4A"`` (no leading ``#``)."""
    return "".join(f"{int(round(_clamp(c, 0, 255))):02X}" for c in rgb[:3])


def mix(color_a: str, color_b: str, weight: float = 0.5) -> str:
    """Blend two colours. ``weight`` is how much of ``color_b`` to use."""
    a, b = to_rgb(color_a), to_rgb(color_b)
    w = _clamp(weight)
    return to_hex([a[i] + (b[i] - a[i]) * w for i in range(3)])


def lighten(color: str, amount: float = 0.2) -> str:
    """Move ``color`` toward white."""
    return mix(color, "FFFFFF", amount)


def darken(color: str, amount: float = 0.2) -> str:
    """Move ``color`` toward black."""
    return mix(color, "000000", amount)


def relative_luminance(color: str) -> float:
    """WCAG 2.x relative luminance of ``color`` (0 = black, 1 = white)."""
    def channel(c: int) -> float:
        s = c / 255.0
        return s / 12.92 if s <= 0.04045 else ((s + 0.055) / 1.055) ** 2.4

    r, g, b = to_rgb(color)
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def contrast_ratio(fg: str, bg: str) -> float:
    """WCAG contrast ratio between two colours (1.0 to 21.0)."""
    l1, l2 = relative_luminance(fg), relative_luminance(bg)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def is_dark(color: str) -> bool:
    """True if ``color`` is dark enough to need light text on top."""
    return relative_luminance(color) < 0.45


def readable_text_on(
    background: str,
    light: str = "FFFFFF",
    dark: str = "111418",
) -> str:
    """Pick whichever of ``light``/``dark`` reads better on ``background``.

    Used wherever text sits on a themed fill — accent chips, gradient
    heroes, coloured table headers, KPI cards. Choosing this by contrast
    ratio instead of hardcoding a colour per theme is what stops the classic
    generated-deck failure of white text on a pale yellow accent.
    """
    return light if contrast_ratio(light, background) >= contrast_ratio(dark, background) else dark


def ensure_contrast(fg: str, bg: str, minimum: float = 4.5) -> str:
    """Nudge ``fg`` until it meets ``minimum`` contrast against ``bg``.

    4.5:1 is the WCAG AA threshold for body text. Darkens or lightens the
    foreground (whichever direction helps) in small steps, falling back to
    the best available black or white.
    """
    if contrast_ratio(fg, bg) >= minimum:
        return fg
    toward = "000000" if not is_dark(bg) else "FFFFFF"
    candidate = fg
    for _ in range(20):
        candidate = mix(candidate, toward, 0.1)
        if contrast_ratio(candidate, bg) >= minimum:
            return candidate
    return readable_text_on(bg)


def alpha_over(color: str, backdrop: str, alpha: float) -> str:
    """Flatten ``color`` at ``alpha`` over an opaque ``backdrop``.

    The PDF backend can use real transparency, but flattening is needed
    wherever a solid colour must stand in for a translucent overlay (for
    example chart fills, which PowerPoint renders inconsistently when given
    an alpha channel).
    """
    return mix(backdrop, color, _clamp(alpha))


def tint_sequence(base: str, count: int) -> List[str]:
    """``count`` tints of ``base``, light to dark.

    A monochrome ramp for single-series charts, which reads as far more
    considered than cycling unrelated accent hues over one data series.
    """
    if count <= 1:
        return [base]
    return [mix(lighten(base, 0.45), darken(base, 0.18), i / (count - 1)) for i in range(count)]


__all__ = [
    "PT_PER_INCH",
    "inches",
    "SLIDE_W",
    "SLIDE_H",
    "SLIDE_W_4_3",
    "SLIDE_H_4_3",
    "Grid",
    "GRID",
    "TypeScale",
    "TYPE",
    "DENSITY",
    "Space",
    "SPACE",
    "char_width",
    "text_width",
    "wrap_text",
    "line_count",
    "text_height",
    "longest_word_width",
    "fit_font_size",
    "truncate_to_width",
    "truncate_to_fit",
    "max_lines",
    "to_rgb",
    "to_hex",
    "mix",
    "lighten",
    "darken",
    "relative_luminance",
    "contrast_ratio",
    "is_dark",
    "readable_text_on",
    "ensure_contrast",
    "alpha_over",
    "tint_sequence",
]
