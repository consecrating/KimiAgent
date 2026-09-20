"""Render a :class:`~kimiagent.models.Deck` straight to a designed ``.pdf``.

Why a second renderer instead of converting the ``.pptx``
---------------------------------------------------------
Converting PowerPoint to PDF requires LibreOffice or Office to be installed,
which makes PDF output an environment problem rather than a feature: it fails
on most CI runners, most containers and most laptops without Office. It also
re-flows type through a different layout engine, so the PDF quietly stops
matching the deck.

So PDF is a first-class output here, drawn directly with ``reportlab`` from the
same :mod:`kimiagent.design` tokens and :mod:`kimiagent.layout` geometry as the
PPTX backend. Same grid, same type scale, same pagination, same computed
contrast — which is what keeps the two in agreement.

If LibreOffice *is* available and byte-level fidelity to the ``.pptx`` matters
more than portability, :func:`render_pdf` will use it when asked
(``engine="soffice"``).

Coordinate systems
------------------
:class:`~kimiagent.layout.Rect` uses a top-left origin to match PowerPoint;
reportlab uses bottom-left. The flip happens in exactly one place,
:meth:`PDFRenderer._y`, so no layout code has to think about it.
"""

from __future__ import annotations

import math
import os
import shutil
import subprocess
import tempfile
from typing import List, Optional, Sequence, Tuple

from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as _canvas

from . import design as D
from .design import DENSITY, TYPE, Grid, TypeScale
from .layout import (
    Rect,
    Zones,
    card_grid,
    column_card_height,
    paginate,
    stat_grid,
    title_line_count,
    zones,
)
from .models import (
    Bullet,
    Card,
    ChartData,
    ComparisonSide,
    Deck,
    ProcessStep,
    Slide,
    SlideType,
    Stat,
    TableData,
    TimelineItem,
    FULL_BLEED_TYPES,
)
from .themes import Theme, get_theme


# --------------------------------------------------------------------------
# Fonts
# --------------------------------------------------------------------------
#
# Core PDF fonts (Helvetica, Times) need no embedding and are metric-identical
# to the tables in :mod:`kimiagent.design`, so PPTX and PDF wrap text at the
# same points. Their downside is Latin-1: they cannot show CJK or most symbols.
# A Unicode TTF is therefore registered *only when the deck needs one*, which
# keeps the common case byte-small and pixel-consistent.

_UNICODE_FONT_CANDIDATES = (
    "/usr/share/fonts/google-noto/NotoSans-Regular.ttf",
    "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans.ttf",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    "C:/Windows/Fonts/arial.ttf",
)
_UNICODE_BOLD_CANDIDATES = (
    "/usr/share/fonts/google-noto/NotoSans-Bold.ttf",
    "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans-Bold.ttf",
    "C:/Windows/Fonts/arialbd.ttf",
)

#: Characters the core fonts lack, mapped to close Latin-1 equivalents. Used
#: only on the core-font path, so nothing is lost when a Unicode face is in use.
_FALLBACK_GLYPHS = {
    "\u2022": "\u2022",   # bullet (present in WinAnsi)
    "\u2013": "\u2013",
    "\u2014": "\u2014",
    "\u2018": "'", "\u2019": "'",
    "\u201c": "\u201c", "\u201d": "\u201d",
    "\u2026": "...",
    "\u2192": "->", "\u2190": "<-", "\u2191": "^", "\u2193": "v",
    "\u2713": "+", "\u2714": "+", "\u2717": "x",
    "\u25aa": "\u2022", "\u25cf": "\u2022", "\u25c6": "\u2022", "\u00b7": "\u00b7",
    "\u2039": "<", "\u203a": ">", "\u00ab": "\u00ab", "\u00bb": "\u00bb",
    "\u20ac": "\u20ac", "\u00a3": "\u00a3", "\u00a5": "\u00a5",
    "\u2011": "-", "\u2212": "-", "\u00a0": " ",
    "\u2009": " ", "\u200a": " ", "\u202f": " ",
}


class _FontBook:
    """Resolves the four faces a deck needs and measures text with them.

    Measurement goes through :func:`reportlab.pdfbase.pdfmetrics.stringWidth`
    for whichever font is actually in use, so wrapping in the PDF is exact
    rather than approximate — the PDF can't overflow even if a Unicode
    fallback face with different metrics gets registered.
    """

    def __init__(self, unicode_needed: bool = False) -> None:
        self.unicode = False
        self.sans = "Helvetica"
        self.sans_bold = "Helvetica-Bold"
        self.serif = "Times-Roman"
        self.serif_bold = "Times-Bold"

        if unicode_needed:
            self._try_register_unicode()

    def _try_register_unicode(self) -> None:
        regular = next((p for p in _UNICODE_FONT_CANDIDATES if os.path.exists(p)), None)
        if not regular:
            return
        bold = next((p for p in _UNICODE_BOLD_CANDIDATES if os.path.exists(p)), regular)
        try:
            pdfmetrics.registerFont(TTFont("KimiSans", regular))
            pdfmetrics.registerFont(TTFont("KimiSans-Bold", bold))
            self.sans = self.serif = "KimiSans"
            self.sans_bold = self.serif_bold = "KimiSans-Bold"
            self.unicode = True
        except Exception:
            # Registration failure is not fatal: fall back to core fonts and
            # let :meth:`sanitize` substitute what cannot be encoded.
            self.unicode = False

    def name(self, bold: bool = False, serif: bool = False) -> str:
        if serif:
            return self.serif_bold if bold else self.serif
        return self.sans_bold if bold else self.sans

    def sanitize(self, text: str) -> str:
        """Replace characters the active font cannot encode."""
        if self.unicode or not text:
            return text
        out = []
        for ch in text:
            if ch in _FALLBACK_GLYPHS:
                out.append(_FALLBACK_GLYPHS[ch])
                continue
            try:
                ch.encode("cp1252")
                out.append(ch)
            except UnicodeEncodeError:
                # Unrepresentable and unmapped: a middot beats a black box.
                out.append("\u00b7" if not ch.isspace() else " ")
        return "".join(out)

    def width(self, text: str, size: float, bold: bool = False, serif: bool = False) -> float:
        try:
            return pdfmetrics.stringWidth(text, self.name(bold, serif), size)
        except Exception:
            return D.text_width(text, size, bold, serif)

    def wrap(
        self, text: str, max_width: float, size: float, bold: bool = False, serif: bool = False
    ) -> List[str]:
        """Greedy wrap using this font's real metrics."""
        if not text:
            return [""]
        if max_width <= 0:
            return [text]
        lines: List[str] = []
        for para in text.split("\n"):
            if not para.strip():
                lines.append("")
                continue
            current = ""
            for word in para.split():
                candidate = f"{current} {word}" if current else word
                if self.width(candidate, size, bold, serif) <= max_width:
                    current = candidate
                    continue
                if current:
                    lines.append(current)
                    current = ""
                if self.width(word, size, bold, serif) <= max_width:
                    current = word
                else:
                    chunk = ""
                    for ch in word:
                        if self.width(chunk + ch, size, bold, serif) <= max_width:
                            chunk += ch
                        else:
                            if chunk:
                                lines.append(chunk)
                            chunk = ch
                    current = chunk
            if current:
                lines.append(current)
        return lines or [""]

    def height(
        self, text: str, max_width: float, size: float, bold: bool = False,
        serif: bool = False, line_height: float = TYPE.line_height,
    ) -> float:
        return len(self.wrap(text, max_width, size, bold, serif)) * size * line_height

    def fit(
        self, text: str, max_width: float, max_height: float, ideal: float,
        minimum: float = 8.0, bold: bool = False, serif: bool = False,
        line_height: float = TYPE.line_height, step: float = 0.5,
        avoid_word_break: bool = True,
    ) -> float:
        """Largest size <= ``ideal`` at which ``text`` fits the box.

        ``avoid_word_break`` mirrors
        :func:`kimiagent.design.fit_font_size`: the widest word must also fit
        on one line, so narrow cards shrink their label instead of splitting a
        word across two lines.
        """
        if not text:
            return ideal
        words = text.split()
        size = ideal
        while size > minimum:
            fits_height = self.height(text, max_width, size, bold, serif, line_height) <= max_height
            fits_words = not avoid_word_break or not words or max(
                self.width(w, size, bold, serif) for w in words
            ) <= max_width
            if fits_height and fits_words:
                return size
            size -= step
        return minimum


def _needs_unicode(deck: Deck) -> bool:
    """True when any deck text falls outside what core PDF fonts can encode."""
    def check(value: str) -> bool:
        for ch in value or "":
            if ch in _FALLBACK_GLYPHS:
                continue
            try:
                ch.encode("cp1252")
            except UnicodeEncodeError:
                return True
        return False

    fields = [deck.title, deck.subtitle, deck.author, deck.footer_note]
    for s in deck.slides:
        fields += [s.title, s.subtitle, s.kicker, s.quote, s.attribution, s.caption, s.verdict]
        fields += [b.text for b in s.bullets] + [b.bold_lead or "" for b in s.bullets]
        for col in s.columns:
            fields += [b.text for b in col]
        fields += [f"{x.value}{x.label}{x.detail or ''}" for x in s.stats]
        fields += [f"{c.title}{c.body}{c.icon}" for c in s.cards]
        fields += [f"{t.period}{t.title}{t.detail}" for t in s.timeline]
        fields += [f"{p.title}{p.detail}" for p in s.steps]
        for side in s.sides:
            fields += [side.title, side.subtitle] + [b.text for b in side.points]
        if s.chart:
            fields += list(s.chart.categories) + [x.name for x in s.chart.series] + [s.chart.takeaway]
        if s.table:
            fields += list(s.table.headers) + [c for row in s.table.rows for c in row]
    return any(check(f) for f in fields)


# --------------------------------------------------------------------------
# Renderer
# --------------------------------------------------------------------------

#: Alignment constants, mirroring the PPTX side without importing pptx.
ALIGN_LEFT, ALIGN_CENTER, ALIGN_RIGHT = "left", "center", "right"
ANCHOR_TOP, ANCHOR_MIDDLE, ANCHOR_BOTTOM = "top", "middle", "bottom"


class PDFRenderer:
    """Draws a deck onto a reportlab canvas, one page per slide."""

    def __init__(self, deck: Deck, theme: Optional[Theme] = None) -> None:
        self.deck = deck
        self.theme: Theme = theme or get_theme(deck.theme)

        if (deck.aspect or "16:9").replace(" ", "") in ("4:3", "4x3"):
            width, height = D.SLIDE_W_4_3, D.SLIDE_H_4_3
        else:
            width, height = D.SLIDE_W, D.SLIDE_H
        self.grid = Grid(width=width, height=height)
        self.scale: TypeScale = TYPE.scaled(DENSITY.get(deck.density or "normal", 1.0))
        self.W, self.H = width, height
        self.fonts = _FontBook(unicode_needed=_needs_unicode(deck))
        self.c: Optional[_canvas.Canvas] = None

    # -- coordinates -----------------------------------------------------

    def _y(self, top: float, height: float = 0.0) -> float:
        """Convert a top-left y to reportlab's bottom-left y."""
        return self.H - top - height

    # -- entry point -----------------------------------------------------

    def render(self, path: str) -> str:
        c = _canvas.Canvas(path, pagesize=(self.W, self.H))
        self.c = c
        c.setTitle(self.deck.title or "Deck")
        if self.deck.author:
            c.setAuthor(self.deck.author)
        if self.deck.subtitle:
            c.setSubject(self.deck.subtitle)
        c.setCreator("KimiAgent")

        dispatch = {
            SlideType.COVER: self._cover,
            SlideType.AGENDA: self._agenda,
            SlideType.SECTION: self._section,
            SlideType.BULLETS: self._bullets,
            SlideType.TWO_COLUMN: self._two_column,
            SlideType.STATS: self._stats,
            SlideType.CARDS: self._cards,
            SlideType.TIMELINE: self._timeline,
            SlideType.PROCESS: self._process,
            SlideType.COMPARISON: self._comparison,
            SlideType.BIG_STATEMENT: self._big_statement,
            SlideType.CHART: self._chart,
            SlideType.TABLE: self._table,
            SlideType.QUOTE: self._quote,
            SlideType.IMAGE: self._image,
            SlideType.CLOSING: self._closing,
        }

        flowed = paginate(self.deck, self.grid, self.scale, self.theme.body_serif)
        total = len(flowed.slides)

        for idx, slide in enumerate(flowed.slides):
            dispatch.get(slide.type, self._bullets)(slide)
            if slide.type not in FULL_BLEED_TYPES:
                self._footer(idx + 1, total)
            # A PDF bookmark per slide: the deck becomes navigable in any
            # reader, which is the main thing a PDF can offer that a deck can't.
            label = slide.title or slide.quote or slide.type.value.title()
            key = f"slide-{idx}"
            try:
                c.bookmarkPage(key)
                c.addOutlineEntry(self.fonts.sanitize(label[:80]), key, level=0)
            except Exception:
                pass
            c.showPage()

        try:
            c.showOutline()
        except Exception:
            pass
        c.save()
        self.c = None
        return path

    # ------------------------------------------------------------------
    # Primitives
    # ------------------------------------------------------------------

    def _fill(self, color: str, alpha: float = 1.0) -> None:
        self.c.setFillColor(HexColor(f"#{color.lstrip('#')}"))
        self.c.setFillAlpha(alpha)

    def _stroke(self, color: str, width: float = 1.0, alpha: float = 1.0) -> None:
        self.c.setStrokeColor(HexColor(f"#{color.lstrip('#')}"))
        self.c.setLineWidth(width)
        self.c.setStrokeAlpha(alpha)

    def _bg(self, color: str) -> None:
        self._fill(color)
        self.c.rect(0, 0, self.W, self.H, stroke=0, fill=1)
        self.c.setFillAlpha(1)

    def _bg_gradient(self) -> Tuple[str, str]:
        """Paint the text-safe hero gradient across the whole page."""
        start, end, angle = self.theme.gradient
        c = self.c
        c.saveState()
        p = c.beginPath()
        p.rect(0, 0, self.W, self.H)
        c.clipPath(p, stroke=0, fill=0)
        # Translate the design-system angle (0 = left-to-right, CCW positive)
        # into the two points reportlab sweeps between.
        rad = math.radians(angle)
        dx, dy = math.cos(rad), math.sin(rad)
        cx, cy = self.W / 2, self.H / 2
        half = (abs(dx) * self.W + abs(dy) * self.H) / 2
        c.linearGradient(
            cx - dx * half, cy - dy * half, cx + dx * half, cy + dy * half,
            (HexColor(f"#{start}"), HexColor(f"#{end}")), extend=True,
        )
        c.restoreState()
        return start, end

    def _rect(
        self,
        r: Rect,
        fill: Optional[str] = None,
        *,
        alpha: float = 1.0,
        radius: Optional[float] = None,
        line: Optional[str] = None,
        line_w: float = 1.0,
        shadow: bool = False,
    ) -> None:
        c = self.c
        y = self._y(r.y, r.h)

        if shadow:
            self._soft_shadow(r, radius)

        if fill:
            self._fill(fill, alpha)
        if line:
            self._stroke(line, line_w)
        do_fill = 1 if fill else 0
        do_stroke = 1 if line else 0
        if not do_fill and not do_stroke:
            return
        if radius and radius > 0:
            c.roundRect(r.x, y, r.w, r.h, min(radius, min(r.w, r.h) / 2),
                        stroke=do_stroke, fill=do_fill)
        else:
            c.rect(r.x, y, r.w, r.h, stroke=do_stroke, fill=do_fill)
        c.setFillAlpha(1)
        c.setStrokeAlpha(1)

    def _soft_shadow(self, r: Rect, radius: Optional[float]) -> None:
        """Approximate a blurred shadow with stacked translucent rounds.

        reportlab has no blur, so the falloff is built from a handful of
        progressively larger, progressively fainter shapes. Four layers is
        enough to read as a soft shadow rather than as an outline.
        """
        c = self.c
        layers = 4
        for i in range(layers, 0, -1):
            spread = i * 1.6
            alpha = 0.055 * (1.0 - (i - 1) / layers)
            self._fill("000000", alpha)
            yy = self._y(r.y + 2.0 + spread * 0.35, r.h)
            rr = (min(radius + spread, min(r.w, r.h) / 2) if radius else 0)
            if rr > 0:
                c.roundRect(r.x - spread, yy - spread, r.w + spread * 2,
                            r.h + spread * 2, rr, stroke=0, fill=1)
            else:
                c.rect(r.x - spread, yy - spread, r.w + spread * 2,
                       r.h + spread * 2, stroke=0, fill=1)
        c.setFillAlpha(1)

    def _card(self, r: Rect, *, fill: Optional[str] = None) -> str:
        """A surface panel following the theme's shape language."""
        t = self.theme
        fill = fill or t.surface
        self._rect(
            r, fill,
            radius=t.radius if t.radius > 0 else None,
            line=t.border if t.hairline else None,
            line_w=0.75,
            shadow=t.shadow,
        )
        return fill

    def _text(
        self,
        r: Rect,
        text: str,
        size: float,
        color: str,
        *,
        bold: bool = False,
        serif: bool = False,
        align: str = ALIGN_LEFT,
        anchor: str = ANCHOR_TOP,
        line_spacing: Optional[float] = None,
        fit: bool = True,
        min_size: float = 8.0,
        tracking: float = 0.0,
        caps: bool = False,
    ) -> float:
        """Draw wrapped text, shrinking to fit. Returns the height used."""
        if not text:
            return 0.0
        text = self.fonts.sanitize(text.upper() if caps else text)
        lh = line_spacing or self.scale.line_height

        if fit:
            size = self.fonts.fit(text, r.w, r.h, size, min_size, bold, serif, lh)
        lines = self.fonts.wrap(text, r.w, size, bold, serif)
        if fit:
            # Trim overflow rather than letting it run past the box.
            room = max(1, int(r.h // (size * lh)))
            if len(lines) > room:
                lines = lines[:room]
                if lines:
                    lines[-1] = self._ellipsize(lines[-1], r.w, size, bold, serif)

        total = len(lines) * size * lh
        if anchor == ANCHOR_MIDDLE:
            top = r.y + (r.h - total) / 2
        elif anchor == ANCHOR_BOTTOM:
            top = r.bottom - total
        else:
            top = r.y

        c = self.c
        self._fill(color)
        font = self.fonts.name(bold, serif)
        c.setFont(font, size)
        if tracking:
            try:
                c.setCharSpace(tracking)
            except Exception:
                tracking = 0.0

        for i, line in enumerate(lines):
            # 0.78em puts the baseline where a text box would: cap height
            # below the top edge, leaving room for descenders.
            baseline_top = top + i * size * lh + size * 0.78
            w = self.fonts.width(line, size, bold, serif)
            if tracking:
                w += tracking * max(0, len(line) - 1)
            if align == ALIGN_CENTER:
                x = r.x + (r.w - w) / 2
            elif align == ALIGN_RIGHT:
                x = r.right - w
            else:
                x = r.x
            c.drawString(x, self._y(baseline_top), line)

        if tracking:
            try:
                c.setCharSpace(0)
            except Exception:
                pass
        c.setFillAlpha(1)
        return total

    def _ellipsize(self, line: str, width: float, size: float, bold: bool, serif: bool) -> str:
        ell = "..." if not self.fonts.unicode else "\u2026"
        words = line.split()
        while words:
            candidate = " ".join(words) + ell
            if self.fonts.width(candidate, size, bold, serif) <= width:
                return candidate
            words.pop()
        return ell

    # -- furniture -------------------------------------------------------

    def _zones(self, s: Slide) -> Zones:
        return zones(
            self.grid, self.scale,
            has_kicker=bool(s.kicker),
            has_subtitle=bool(s.subtitle),
            title_lines=title_line_count(s.title, self.grid, self.scale, self.theme.heading_serif),
        )

    def _accent_for(self, s: Slide, index: int = 0) -> str:
        if s.accent is not None:
            return self.theme.accent(s.accent + index)
        return self.theme.accent(index) if index else self.theme.secondary

    def _heading(self, s: Slide, *, on: Optional[str] = None) -> Zones:
        t = self.theme
        bg = on or t.background
        z = self._zones(s)
        if s.kicker:
            self._text(z.kicker, s.kicker, self.scale.caption, t.accent_on(bg),
                       bold=True, caps=True, tracking=1.4, fit=False)
        if s.title:
            self._text(z.title, s.title, self.scale.h1, t.ink_on(bg), bold=t.heading_bold,
                       serif=t.heading_serif, line_spacing=self.scale.line_height_tight)
        self._rect(z.rule, self._accent_for(s))
        if s.subtitle:
            self._text(z.subtitle, s.subtitle, self.scale.body_sm, t.muted_on(bg),
                       serif=t.body_serif)
        return z

    def _footer(self, page: int, total: int) -> None:
        t, g = self.theme, self.grid
        z = zones(g, self.scale)
        label_w = z.footer.w * 0.7
        label = D.truncate_to_width(
            self.deck.footer_note or self.deck.title, label_w, self.scale.micro
        )
        self._text(Rect(z.footer.x, z.footer.y + 4, label_w, z.footer.h),
                   label, self.scale.micro, t.muted_on(t.background), fit=False)
        self._text(Rect(g.right - 90, z.footer.y + 4, 90, z.footer.h),
                   f"{page} / {total}", self.scale.micro, t.muted_on(t.background),
                   align=ALIGN_RIGHT, fit=False)
        y = g.height - 3.0
        self._rect(Rect(0, y, g.width, 3.0), t.border, alpha=0.55)
        if total:
            self._rect(Rect(0, y, g.width * (page / total), 3.0), t.secondary)

    def _hero_decor(self, accent: str) -> None:
        g = self.grid
        r1 = min(g.width, g.height) * 0.62
        self._ellipse(Rect(g.width - r1 * 0.55, g.height - r1 * 0.62, r1, r1), accent, 0.14)
        r2 = r1 * 0.58
        self._ellipse(Rect(g.width - r2 * 0.30, g.height - r2 * 1.35, r2, r2), accent, 0.20)

    def _ellipse(self, r: Rect, fill: str, alpha: float = 1.0,
                 line: Optional[str] = None, line_w: float = 1.0) -> None:
        c = self.c
        y = self._y(r.y, r.h)
        if fill:
            self._fill(fill, alpha)
        if line:
            self._stroke(line, line_w)
        c.ellipse(r.x, y, r.x + r.w, y + r.h, stroke=1 if line else 0, fill=1 if fill else 0)
        c.setFillAlpha(1)
        c.setStrokeAlpha(1)

    def _triangle(self, r: Rect, fill: str, down: bool = False) -> None:
        c = self.c
        y = self._y(r.y, r.h)
        self._fill(fill)
        p = c.beginPath()
        if down:
            p.moveTo(r.x, y + r.h)
            p.lineTo(r.x + r.w, y + r.h)
            p.lineTo(r.x + r.w / 2, y)
        else:
            p.moveTo(r.x, y)
            p.lineTo(r.x + r.w, y)
            p.lineTo(r.x + r.w / 2, y + r.h)
        p.close()
        c.drawPath(p, stroke=0, fill=1)
        c.setFillAlpha(1)

    # ------------------------------------------------------------------
    # Slide builders
    # ------------------------------------------------------------------

    def _cover(self, s: Slide) -> None:
        t, g = self.theme, self.grid
        self._bg_gradient()
        self._hero_decor(t.secondary)
        ink = t.hero_text
        muted = D.mix(ink, D.mix(*t.gradient[:2], 0.5), 0.28)

        left, width = g.left, g.content_w * 0.82
        y = g.height * 0.30
        kicker = s.kicker or self.deck.subtitle or ""
        if kicker:
            y += self._text(Rect(left, y, width, 22), kicker, self.scale.caption, ink,
                            bold=True, caps=True, tracking=1.8, fit=False) + D.SPACE.sm

        y += self._text(
            Rect(left, y, width, self.scale.display * 2.6), s.title or self.deck.title,
            self.scale.display, ink, bold=True, serif=t.heading_serif,
            line_spacing=self.scale.line_height_tight, min_size=self.scale.h2,
        ) + D.SPACE.md

        self._rect(Rect(left, y, 62, 4), t.secondary)
        y += 4 + D.SPACE.md
        if s.subtitle:
            self._text(Rect(left, y, width * 0.88, self.scale.h3 * 3), s.subtitle,
                       self.scale.h3, muted, serif=t.body_serif)

        byline = " \u00b7 ".join(x for x in (self.deck.author, _today()) if x)
        if byline:
            self._text(Rect(left, g.bottom - 24, width, 22), byline,
                       self.scale.body_sm, muted, fit=False)

    def _closing(self, s: Slide) -> None:
        t, g = self.theme, self.grid
        self._bg_gradient()
        self._hero_decor(t.secondary)
        ink = t.hero_text
        box_w = g.content_w * 0.78
        x = (g.width - box_w) / 2
        y = g.height * 0.34
        y += self._text(Rect(x, y, box_w, self.scale.hero * 2.4), s.title or "Thank you",
                        self.scale.hero, ink, bold=True, align=ALIGN_CENTER,
                        serif=t.heading_serif,
                        line_spacing=self.scale.line_height_tight) + D.SPACE.md
        self._rect(Rect(g.width / 2 - 31, y, 62, 4), t.secondary)
        y += 4 + D.SPACE.md
        if s.subtitle:
            self._text(Rect(x, y, box_w, self.scale.h3 * 3), s.subtitle, self.scale.h3,
                       D.mix(ink, D.mix(*t.gradient[:2], 0.5), 0.25),
                       align=ALIGN_CENTER, serif=t.body_serif)

    def _section(self, s: Slide) -> None:
        t, g = self.theme, self.grid
        bg = t.primary if not t.dark else t.surface
        self._bg(bg)
        ink = t.on(bg)
        accent = self._accent_for(s)

        number = (s.subtitle or "").strip()
        if number:
            self._text(Rect(g.right - 340, g.height * 0.16, 340, 260), number, 190,
                       D.mix(bg, ink, 0.16), bold=True, align=ALIGN_RIGHT,
                       serif=t.heading_serif, fit=False, line_spacing=1.0)

        y = g.height * 0.40
        if s.kicker:
            self._text(Rect(g.left, y - 30, g.content_w, 20), s.kicker, self.scale.caption,
                       accent, bold=True, caps=True, tracking=1.6, fit=False)
        self._rect(Rect(g.left, y, 62, 4), accent)
        y += 4 + D.SPACE.lg
        self._text(Rect(g.left, y, g.content_w * 0.74, self.scale.hero * 2.2), s.title,
                   self.scale.hero, ink, bold=True, serif=t.heading_serif,
                   line_spacing=self.scale.line_height_tight)

    def _big_statement(self, s: Slide) -> None:
        t, g = self.theme, self.grid
        self._bg_gradient()
        ink = t.hero_text
        accent = self._accent_for(s)
        box_w = g.content_w * 0.86
        x = (g.width - box_w) / 2

        if s.kicker:
            self._text(Rect(x, g.height * 0.22, box_w, 22), s.kicker, self.scale.caption,
                       ink, bold=True, caps=True, tracking=1.8, align=ALIGN_CENTER, fit=False)
        band = Rect(x, g.height * 0.28, box_w, g.height * 0.44)
        used = self._text(band, s.title or s.quote, self.scale.hero * 1.12, ink,
                          bold=True, align=ALIGN_CENTER, anchor=ANCHOR_MIDDLE,
                          serif=t.heading_serif,
                          line_spacing=self.scale.line_height_tight,
                          min_size=self.scale.h2)
        y = band.y + (band.h + used) / 2 + D.SPACE.lg
        self._rect(Rect(g.width / 2 - 31, y, 62, 4), accent)
        if s.subtitle or s.attribution:
            self._text(Rect(x, y + 4 + D.SPACE.md, box_w, 40),
                       s.subtitle or f"\u2014 {s.attribution}",
                       self.scale.body, D.mix(ink, D.mix(*t.gradient[:2], 0.5), 0.25),
                       align=ALIGN_CENTER, serif=t.body_serif)

    # -- prose -----------------------------------------------------------

    def _bullets(self, s: Slide) -> None:
        self._bg(self.theme.background)
        z = self._heading(s)
        self._render_bullets(s.bullets, z.body, accent=self._accent_for(s))

    def _render_bullets(
        self,
        bullets: Sequence[Bullet],
        r: Rect,
        *,
        accent: Optional[str] = None,
        on: Optional[str] = None,
        size: Optional[float] = None,
        compact: bool = False,
    ) -> float:
        """Draw a bullet list with hand-placed glyphs and hanging indents."""
        if not bullets:
            return 0.0
        t = self.theme
        bg = on or t.background
        accent = accent or t.secondary
        base = size or self.scale.body
        gap = D.SPACE.xs if compact else D.SPACE.sm
        lh = self.scale.line_height

        def measure(shrink: float) -> float:
            total = 0.0
            for b in bullets:
                sz = (base if b.level == 0 else base * 0.87) * shrink
                indent = 16.0 + b.level * 16.0
                text = f"{b.bold_lead}: {b.text}" if b.bold_lead else b.text
                total += self.fonts.height(text, r.w - indent, sz, False, t.body_serif, lh) + gap
            return total

        shrink = 1.0
        while measure(shrink) > r.h and shrink > 0.62:
            shrink -= 0.04

        y = r.y
        for b in bullets:
            sz = (base if b.level == 0 else base * 0.87) * shrink
            indent = 16.0 + b.level * 16.0
            marker = "\u2022" if b.level == 0 else "\u2013"
            color = t.ink_on(bg) if b.level == 0 else t.muted_on(bg)

            # The glyph is drawn on the first line's baseline, and the text
            # block starts at the indent, which is what gives the hanging
            # indent the PPTX side gets from native bullet formatting.
            self._text(Rect(r.x + b.level * 16.0, y, 14, sz * lh), marker, sz,
                       accent if b.level == 0 else t.muted_on(bg), bold=True,
                       fit=False, line_spacing=lh)

            block = Rect(r.x + indent, y, r.w - indent, max(sz * lh, r.bottom - y))
            if b.bold_lead:
                lead = self.fonts.sanitize(f"{b.bold_lead}: ")
                lead_w = self.fonts.width(lead, sz, True, t.body_serif)
                self._text(Rect(block.x, y, lead_w + 1, sz * lh), lead, sz,
                           t.ink_on(bg), bold=True, serif=t.body_serif,
                           fit=False, line_spacing=lh)
                # Continuation of the same visual line: first line is inset by
                # the lead width, remaining lines return to the block edge.
                used = self._flow_after_lead(b.text, block, sz, color, lead_w, lh)
            else:
                used = self._text(block, b.text, sz, color, serif=t.body_serif,
                                  fit=False, line_spacing=lh)
            y += used + gap
            if y > r.bottom:
                break
        return y - r.y

    def _flow_after_lead(
        self, text: str, block: Rect, size: float, color: str, lead_w: float, lh: float
    ) -> float:
        """Lay out ``text`` where the first line starts after a bold lead-in."""
        t = self.theme
        words = self.fonts.sanitize(text).split()
        lines: List[Tuple[float, str]] = []
        current, offset = "", lead_w
        for word in words:
            candidate = f"{current} {word}" if current else word
            if self.fonts.width(candidate, size, False, t.body_serif) <= block.w - offset:
                current = candidate
            else:
                lines.append((offset, current))
                current, offset = word, 0.0
        if current:
            lines.append((offset, current))
        for i, (off, line) in enumerate(lines):
            self._text(Rect(block.x + off, block.y + i * size * lh, block.w - off, size * lh),
                       line, size, color, serif=t.body_serif, fit=False, line_spacing=lh)
        return max(1, len(lines)) * size * lh

    def _two_column(self, s: Slide) -> None:
        t = self.theme
        self._bg(t.background)
        z = self._heading(s)
        cols = (s.columns + [[], []])[:2]
        rects = z.body.split_h(2, D.SPACE.md)
        card_h = column_card_height(cols, rects[0].w, z.body, self.scale, t.body_serif)
        for i, (col, r) in enumerate(zip(cols, rects)):
            r = r.with_h(card_h)
            self._card(r)
            accent = self._accent_for(s, i)
            self._rect(Rect(r.x, r.y, r.w, 4), accent)
            inner = Rect(r.x + D.SPACE.lg, r.y + D.SPACE.lg + 4,
                         r.w - D.SPACE.lg * 2, r.h - D.SPACE.lg * 2 - 4)
            self._render_bullets(col, inner, accent=accent, on=t.surface)

    def _quote(self, s: Slide) -> None:
        t, g = self.theme, self.grid
        self._bg(t.surface)
        accent = self._accent_for(s)
        self._text(Rect(g.left - 6, g.top + 6, 200, 220), "\u201c", 200,
                   D.mix(t.surface, accent, 0.30), bold=True, serif=True,
                   fit=False, line_spacing=1.0)
        box = Rect(g.left + 54, g.height * 0.24, g.content_w - 100, g.height * 0.42)
        used = self._text(box, s.quote or s.title, self.scale.h1 * 1.05,
                          t.ink_on(t.surface), bold=True, anchor=ANCHOR_MIDDLE,
                          serif=t.heading_serif,
                          line_spacing=self.scale.line_height_tight,
                          min_size=self.scale.h3)
        y = box.y + (box.h + used) / 2 + D.SPACE.lg
        self._rect(Rect(box.x, y, 40, 3), accent)
        if s.attribution:
            self._text(Rect(box.x, y + 3 + D.SPACE.md, box.w, 40), s.attribution,
                       self.scale.body_sm, t.muted_on(t.surface), caps=True,
                       tracking=1.2, bold=True, fit=False)

    # -- structured ------------------------------------------------------

    def _stats(self, s: Slide) -> None:
        t = self.theme
        self._bg(t.background)
        z = self._heading(s)
        if not s.stats:
            return
        for i, (st, r) in enumerate(zip(s.stats, stat_grid(z.body, len(s.stats), D.SPACE.md))):
            accent = self._accent_for(s, i) if s.accent is not None else t.accent(i)
            self._card(r)
            self._rect(Rect(r.x, r.y, r.w, 4), accent)
            inner = r.inset(D.SPACE.md, D.SPACE.md)
            value_h = inner.h * 0.46
            self._text(Rect(inner.x, r.y + D.SPACE.lg, inner.w, value_h), st.value,
                       self.scale.display * 0.86, D.ensure_contrast(accent, t.surface, 3.0),
                       bold=True, align=ALIGN_CENTER, anchor=ANCHOR_MIDDLE,
                       serif=t.heading_serif, line_spacing=1.0, min_size=self.scale.h2)
            y = r.y + D.SPACE.lg + value_h + D.SPACE.xs
            if st.trend:
                self._trend(Rect(inner.cx - 6, y, 12, 10), st.trend, accent)
                y += 14
            y += self._text(Rect(inner.x, y, inner.w, 44), st.label, self.scale.body_sm,
                            t.ink_on(t.surface), bold=True, align=ALIGN_CENTER,
                            serif=t.body_serif)
            if st.detail:
                self._text(Rect(inner.x, y + D.SPACE.xxs, inner.w,
                                max(14.0, r.bottom - y - D.SPACE.md)),
                           st.detail, self.scale.caption, t.muted_on(t.surface),
                           align=ALIGN_CENTER, serif=t.body_serif)

    def _trend(self, r: Rect, trend: str, accent: str) -> None:
        trend = (trend or "").lower()
        if trend in ("up", "rise", "increase", "+"):
            self._triangle(r, accent, down=False)
        elif trend in ("down", "fall", "decrease", "-"):
            self._triangle(r, accent, down=True)
        else:
            self._rect(Rect(r.x, r.cy - 1.5, r.w, 3),
                       self.theme.muted_on(self.theme.surface))

    def _cards(self, s: Slide) -> None:
        t = self.theme
        self._bg(t.background)
        z = self._heading(s)
        if not s.cards:
            return
        for i, (card, r) in enumerate(zip(s.cards, card_grid(z.body, len(s.cards), D.SPACE.md))):
            accent = t.accent(card.accent) if card.accent is not None else t.accent(i)
            self._card(r)
            inner = r.inset(D.SPACE.lg, D.SPACE.lg)
            y = inner.y
            if card.icon:
                tile = Rect(inner.x, y, 34, 34)
                tile_bg = D.mix(t.surface, accent, 0.20)
                self._rect(tile, tile_bg, radius=9 if t.radius else None)
                self._text(tile, card.icon, 15,
                           D.ensure_contrast(accent, tile_bg, 3.0), align=ALIGN_CENTER,
                           anchor=ANCHOR_MIDDLE, fit=False, line_spacing=1.0)
                y += tile.h + D.SPACE.sm
            else:
                self._rect(Rect(inner.x, y + 4, 26, 3), accent)
                y += D.SPACE.md
            y += self._text(Rect(inner.x, y, inner.w, self.scale.h3 * 2.6), card.title,
                            self.scale.h3, t.ink_on(t.surface), bold=True,
                            serif=t.heading_serif,
                            line_spacing=self.scale.line_height_tight)
            if card.body:
                self._text(Rect(inner.x, y + D.SPACE.xs, inner.w,
                                max(16.0, inner.bottom - y - D.SPACE.xs)),
                           card.body, self.scale.body_sm, t.muted_on(t.surface),
                           serif=t.body_serif)

    def _timeline(self, s: Slide) -> None:
        t = self.theme
        self._bg(t.background)
        z = self._heading(s)
        items = s.timeline
        if not items:
            return
        axis_y = z.body.y + z.body.h * 0.5
        n = len(items)
        usable = z.body.w * (1 - 1.0 / max(2, n))
        x0 = z.body.x + (z.body.w - usable) / 2
        step = usable / max(1, n - 1) if n > 1 else 0

        self._rect(Rect(z.body.x, axis_y - 1, z.body.w, 2), t.border)
        done = [i for i, it in enumerate(items) if it.done]
        if done:
            end_x = x0 + step * max(done)
            self._rect(Rect(z.body.x, axis_y - 1.5, end_x - z.body.x, 3), t.secondary)

        label_h = z.body.h * 0.40
        col_w = max(90.0, (step if n > 1 else z.body.w) * 1.5)
        for i, it in enumerate(items):
            cx = x0 + step * i if n > 1 else z.body.cx
            accent = t.accent(i)
            above = i % 2 == 0
            node = 15.0
            self._ellipse(Rect(cx - node / 2, axis_y - node / 2, node, node),
                          accent if it.done else t.background, 1.0, line=accent, line_w=2.0)
            if it.done:
                self._ellipse(Rect(cx - 3, axis_y - 3, 6, 6), t.on(accent))

            stem = 16.0
            sy = axis_y - node / 2 - stem if above else axis_y + node / 2
            self._rect(Rect(cx - 0.75, sy, 1.5, stem), t.border)

            bx = max(z.body.x, min(cx - col_w / 2, z.body.right - col_w))
            by = sy - label_h if above else sy + stem
            block = Rect(bx, by, col_w, label_h)
            y = block.bottom - self._measure_timeline(it, block.w) if above else block.y

            if it.period:
                y += self._text(Rect(block.x, y, block.w, 18), it.period, self.scale.caption,
                                D.ensure_contrast(accent, t.background, 3.0), bold=True,
                                caps=True, tracking=1.0, align=ALIGN_CENTER,
                                fit=False) + D.SPACE.xxs
            y += self._text(Rect(block.x, y, block.w, self.scale.h3 * 2.4), it.title,
                            self.scale.body, t.ink_on(t.background), bold=True,
                            align=ALIGN_CENTER, serif=t.heading_serif,
                            line_spacing=self.scale.line_height_tight)
            if it.detail:
                self._text(Rect(block.x, y + D.SPACE.xxs, block.w,
                                max(14.0, block.bottom - y - D.SPACE.xxs)),
                           it.detail, self.scale.caption, t.muted_on(t.background),
                           align=ALIGN_CENTER, serif=t.body_serif)

    def _measure_timeline(self, it: TimelineItem, w: float) -> float:
        t = self.theme
        h = 18 + D.SPACE.xxs if it.period else 0.0
        h += self.fonts.height(it.title, w, self.scale.body, True, t.heading_serif,
                               self.scale.line_height_tight)
        if it.detail:
            h += D.SPACE.xxs + self.fonts.height(it.detail, w, self.scale.caption,
                                                 False, t.body_serif)
        return h

    def _process(self, s: Slide) -> None:
        t = self.theme
        self._bg(t.background)
        z = self._heading(s)
        steps = s.steps
        if not steps:
            return
        n = len(steps)
        gap, arrow = D.SPACE.md, 16.0
        w = (z.body.w - (n - 1) * (gap + arrow)) / n
        h = min(z.body.h, 210.0)
        top = z.body.y + (z.body.h - h) / 2

        for i, st in enumerate(steps):
            x = z.body.x + i * (w + gap + arrow)
            r = Rect(x, top, w, h)
            accent = t.accent(i)
            self._card(r)
            inner = r.inset(D.SPACE.md, D.SPACE.md)
            d = 38.0
            disc = Rect(inner.x, inner.y, d, d)
            self._ellipse(disc, accent)
            self._text(disc, str(i + 1), self.scale.h3, t.on(accent), bold=True,
                       align=ALIGN_CENTER, anchor=ANCHOR_MIDDLE, fit=False, line_spacing=1.0)
            y = disc.bottom + D.SPACE.sm
            y += self._text(Rect(inner.x, y, inner.w, self.scale.h3 * 2.4), st.title,
                            self.scale.h3, t.ink_on(t.surface), bold=True,
                            serif=t.heading_serif,
                            line_spacing=self.scale.line_height_tight)
            if st.detail:
                self._text(Rect(inner.x, y + D.SPACE.xs, inner.w,
                                max(14.0, inner.bottom - y - D.SPACE.xs)),
                           st.detail, self.scale.body_sm, t.muted_on(t.surface),
                           serif=t.body_serif)
            if i < n - 1:
                ax = r.right + gap / 2
                self._right_arrow(Rect(ax, top + h / 2 - arrow / 2, arrow, arrow), t.border)

    def _right_arrow(self, r: Rect, fill: str) -> None:
        c = self.c
        y = self._y(r.y, r.h)
        self._fill(fill)
        p = c.beginPath()
        p.moveTo(r.x, y)
        p.lineTo(r.x, y + r.h)
        p.lineTo(r.x + r.w, y + r.h / 2)
        p.close()
        c.drawPath(p, stroke=0, fill=1)
        c.setFillAlpha(1)

    def _comparison(self, s: Slide) -> None:
        t = self.theme
        self._bg(t.background)
        z = self._heading(s)
        sides = s.sides[:2]
        if not sides:
            return
        body = z.body
        verdict_h = 0.0
        if s.verdict:
            verdict_h = 54.0
            body = Rect(body.x, body.y, body.w, body.h - verdict_h - D.SPACE.md)

        for i, (side, r) in enumerate(zip(sides, body.split_h(len(sides), D.SPACE.lg))):
            accent = t.accent(i)
            if side.preferred:
                surface = D.mix(t.surface, accent, 0.10)
                self._rect(r, surface, radius=t.radius if t.radius else None,
                           line=accent, line_w=2.0, shadow=t.shadow)
            else:
                surface = self._card(r)
            inner = r.inset(D.SPACE.lg, D.SPACE.lg)
            y = inner.y
            if side.preferred:
                badge = Rect(inner.x, y, 108.0, 20.0)
                self._rect(badge, accent, radius=badge.h / 2 if t.radius else None)
                self._text(badge, "RECOMMENDED", self.scale.micro, t.on(accent),
                           bold=True, tracking=1.0, align=ALIGN_CENTER,
                           anchor=ANCHOR_MIDDLE, fit=False, line_spacing=1.0)
                y += badge.h + D.SPACE.sm
            y += self._text(Rect(inner.x, y, inner.w, self.scale.h2 * 2.2), side.title,
                            self.scale.h2, t.ink_on(surface), bold=True,
                            serif=t.heading_serif,
                            line_spacing=self.scale.line_height_tight)
            if side.subtitle:
                y += self._text(Rect(inner.x, y + D.SPACE.xxs, inner.w, 34), side.subtitle,
                                self.scale.body_sm, t.muted_on(surface),
                                serif=t.body_serif) + D.SPACE.xxs
            y += D.SPACE.sm
            self._rect(Rect(inner.x, y, 28, 3), accent)
            y += 3 + D.SPACE.sm
            self._render_bullets(side.points, Rect(inner.x, y, inner.w, max(20.0, inner.bottom - y)),
                                 accent=accent, on=surface, size=self.scale.body_sm, compact=True)

        if s.verdict:
            vr = Rect(z.body.x, z.body.bottom - verdict_h, z.body.w, verdict_h)
            accent = t.secondary
            bgv = D.mix(t.background, accent, 0.12)
            self._rect(vr, bgv, radius=t.radius if t.radius else None)
            self._rect(Rect(vr.x, vr.y, 4, vr.h), accent)
            self._text(Rect(vr.x + D.SPACE.lg, vr.y + 9, 74, 16), "VERDICT", self.scale.micro,
                       D.ensure_contrast(accent, bgv, 3.0), bold=True, caps=True,
                       tracking=1.2, fit=False)
            self._text(Rect(vr.x + D.SPACE.lg, vr.y + 24, vr.w - D.SPACE.lg * 2, 24),
                       s.verdict, self.scale.body, t.ink_on(bgv), bold=True, serif=t.body_serif)

    def _agenda(self, s: Slide) -> None:
        t = self.theme
        self._bg(t.background)
        z = self._heading(s if s.title else _retitled(s, "Agenda"))
        items = s.bullets
        if not items:
            return
        for i, (b, r) in enumerate(zip(items, z.body.split_v(len(items), D.SPACE.xs))):
            accent = t.accent(i)
            self._text(Rect(r.x, r.y, 54, r.h), f"{i + 1:02d}",
                       min(self.scale.h1, r.h * 0.62),
                       D.ensure_contrast(accent, t.background, 3.0), bold=True,
                       serif=t.heading_serif, anchor=ANCHOR_MIDDLE, fit=False,
                       line_spacing=1.0)
            self._text(Rect(r.x + 66, r.y, r.w - 66, r.h), b.text, self.scale.h3,
                       t.ink_on(t.background), bold=True, anchor=ANCHOR_MIDDLE,
                       serif=t.heading_serif, line_spacing=self.scale.line_height_tight)
            if i < len(items) - 1:
                self._rect(Rect(r.x, r.bottom, r.w, 0.75), t.border)

    # -- data ------------------------------------------------------------

    def _table(self, s: Slide) -> None:
        t = self.theme
        self._bg(t.background)
        z = self._heading(s)
        td = s.table
        if not td or not td.headers:
            return
        rows = len(td.rows) + 1
        cols = len(td.headers)
        row_h = min(42.0, max(26.0, z.body.h / max(1, rows)))
        col_w = z.body.w / cols
        pad = D.SPACE.sm

        header_bg = D.mix(t.background, t.secondary, 0.14)
        self._rect(Rect(z.body.x, z.body.y, z.body.w, row_h), header_bg)
        for c_i, header in enumerate(td.headers):
            cell = Rect(z.body.x + c_i * col_w + pad, z.body.y, col_w - pad * 2, row_h)
            self._text(cell, str(header), self.scale.body_sm,
                       D.ensure_contrast(t.text, header_bg, 4.5), bold=True,
                       anchor=ANCHOR_MIDDLE, align=ALIGN_RIGHT if c_i in td.numeric_columns else ALIGN_LEFT,
                       serif=t.body_serif, fit=False)

        for r_i, row in enumerate(td.rows):
            y = z.body.y + row_h * (r_i + 1)
            highlight = td.highlight_row is not None and r_i == td.highlight_row
            band = (D.mix(t.background, t.secondary, 0.10) if highlight
                    else (t.background if (r_i + 1) % 2 else t.surface))
            self._rect(Rect(z.body.x, y, z.body.w, row_h), band)
            for c_i in range(cols):
                value = row[c_i] if c_i < len(row) else ""
                cell = Rect(z.body.x + c_i * col_w + pad, y, col_w - pad * 2, row_h)
                self._text(cell, str(value), self.scale.body_sm,
                           D.ensure_contrast(t.text, band, 4.5), bold=highlight,
                           anchor=ANCHOR_MIDDLE,
                           align=ALIGN_RIGHT if c_i in td.numeric_columns else ALIGN_LEFT,
                           serif=t.body_serif, fit=False)

        self._rect(Rect(z.body.x, z.body.y + row_h - 1.5, z.body.w, 1.5), t.secondary)

    def _image(self, s: Slide) -> None:
        t = self.theme
        self._bg(t.background)
        z = self._heading(s) if s.title else None
        area = z.body if z else Rect(self.grid.left, self.grid.top,
                                     self.grid.content_w, self.grid.content_h)
        if s.caption:
            area = Rect(area.x, area.y, area.w, area.h - 28)

        placed = False
        if s.image_path and os.path.exists(s.image_path):
            try:
                from reportlab.lib.utils import ImageReader

                img = ImageReader(s.image_path)
                iw, ih = img.getSize()
                ratio = min(area.w / iw, area.h / ih)
                w, h = iw * ratio, ih * ratio
                self.c.drawImage(img, area.cx - w / 2, self._y(area.cy - h / 2, h),
                                 width=w, height=h, mask="auto")
                placed = True
            except Exception:
                placed = False

        if not placed:
            self._card(area)
            label = "Image placeholder"
            if s.image_path:
                label = f"Image not found: {os.path.basename(s.image_path)}"
            self._text(Rect(area.x, area.cy - 14, area.w, 28), label, self.scale.body,
                       t.muted_on(t.surface), align=ALIGN_CENTER)
        if s.caption:
            self._text(Rect(area.x, area.bottom + D.SPACE.xs, area.w, 22), s.caption,
                       self.scale.caption, t.muted_on(t.background),
                       align=ALIGN_CENTER, serif=t.body_serif)

    # -- charts ----------------------------------------------------------

    def _chart(self, s: Slide) -> None:
        """Draw the chart directly.

        Charts are drawn rather than delegated to ``reportlab.graphics.charts``
        because the goal is a *specific* modern look — no frame, no plot-area
        fill, hairline horizontal gridlines only, themed series, generous bar
        gaps. Fighting a charting library's defaults to reach that is more
        code than drawing the handful of marks involved.
        """
        t = self.theme
        self._bg(t.background)
        z = self._heading(s)
        cd = s.chart
        if not cd or not cd.series:
            self._text(Rect(z.body.x, z.body.y, z.body.w, 40), "(No chart data)",
                       self.scale.body, t.muted_on(t.background))
            return

        plot = z.body
        if cd.takeaway:
            th = self.fonts.height(cd.takeaway, z.body.w - 56, self.scale.body_sm,
                                   False, t.body_serif) + D.SPACE.md * 2
            plot = Rect(z.body.x, z.body.y, z.body.w, max(120.0, z.body.h - th - D.SPACE.md))

        kind = (cd.chart_type or "column").lower()
        legend_h = 22.0 if (len(cd.series) > 1 or kind in ("pie", "doughnut", "donut")) else 0.0
        plot_area = Rect(plot.x, plot.y, plot.w, plot.h - legend_h)

        if kind in ("pie", "doughnut", "donut"):
            self._draw_pie(plot_area, cd, doughnut=kind in ("doughnut", "donut"))
        elif kind == "bar":
            self._draw_bars(plot_area, cd, horizontal=True)
        elif kind in ("line", "line_plain", "area"):
            self._draw_lines(plot_area, cd, area_fill=kind == "area")
        else:
            self._draw_bars(plot_area, cd, horizontal=False,
                            stacked=kind.startswith("stacked"))

        if legend_h:
            self._draw_legend(Rect(plot.x, plot_area.bottom, plot.w, legend_h), cd, kind)

        if cd.takeaway:
            y = plot.bottom + D.SPACE.md
            accent = self._accent_for(s)
            self._rect(Rect(z.body.x, y, 3, max(20.0, z.body.bottom - y)), accent)
            self._text(Rect(z.body.x + D.SPACE.md, y, z.body.w - D.SPACE.md,
                            max(20.0, z.body.bottom - y)),
                       cd.takeaway, self.scale.body_sm, t.ink_on(t.background),
                       bold=True, serif=t.body_serif)

    def _nice_ticks(self, lo: float, hi: float, count: int = 5) -> Tuple[float, float, float]:
        """Round axis bounds to human numbers. Returns ``(lo, hi, step)``.

        Axis labels like 0/25/50/75/100 rather than 0/23.7/47.4 are most of
        what makes a chart feel finished.
        """
        if hi <= lo:
            hi = lo + 1
        raw = (hi - lo) / max(1, count)
        mag = 10 ** math.floor(math.log10(raw)) if raw > 0 else 1
        for mult in (1, 2, 2.5, 5, 10):
            step = mag * mult
            if step >= raw:
                break
        lo_t = math.floor(lo / step) * step
        hi_t = math.ceil(hi / step) * step
        return lo_t, hi_t, step

    def _axis_frame(
        self, area: Rect, cd: ChartData, stacked: bool = False
    ) -> Tuple[Rect, float, float, float]:
        """Draw gridlines and axis labels. Returns the inner plot rect + scale."""
        t = self.theme
        values = [v for s_ in cd.series for v in s_.values] or [0.0]
        if stacked:
            n = max(len(s_.values) for s_ in cd.series)
            totals = [sum((s_.values[i] if i < len(s_.values) else 0) for s_ in cd.series)
                      for i in range(n)]
            values = totals + [0.0]
        lo, hi, step = self._nice_ticks(min(min(values), 0.0), max(values))

        # Reserve space for the widest y label so labels never collide.
        label_w = max(
            self.fonts.width(_fmt_num(v), self.scale.caption)
            for v in _frange(lo, hi, step)
        ) + D.SPACE.sm
        cat_h = self.scale.caption * 2.0
        inner = Rect(area.x + label_w, area.y, area.w - label_w, area.h - cat_h)

        for v in _frange(lo, hi, step):
            frac = (v - lo) / (hi - lo) if hi > lo else 0
            y = inner.bottom - frac * inner.h
            self._rect(Rect(inner.x, y - 0.4, inner.w, 0.8), t.chart_grid)
            self._text(Rect(area.x, y - self.scale.caption * 0.72, label_w - D.SPACE.sm,
                            self.scale.caption * 1.4),
                       _fmt_num(v), self.scale.caption, t.muted_on(t.background),
                       align=ALIGN_RIGHT, fit=False, line_spacing=1.0)

        cats = cd.categories or [str(i + 1) for i in range(len(cd.series[0].values))]
        cw = inner.w / max(1, len(cats))
        for i, cat in enumerate(cats):
            self._text(Rect(inner.x + i * cw, inner.bottom + D.SPACE.xs, cw,
                            self.scale.caption * 1.5),
                       str(cat), self.scale.caption, t.muted_on(t.background),
                       align=ALIGN_CENTER, fit=False, line_spacing=1.0)
        return inner, lo, hi, step

    def _draw_bars(self, area: Rect, cd: ChartData, horizontal: bool = False,
                   stacked: bool = False) -> None:
        t = self.theme
        if horizontal:
            # A horizontal bar chart is a column chart with the roles of the
            # axes swapped; drawing it as its own case keeps labels readable.
            self._draw_hbars(area, cd)
            return
        inner, lo, hi, _ = self._axis_frame(area, cd, stacked)
        cats = cd.categories or [str(i + 1) for i in range(len(cd.series[0].values))]
        n_cat = max(1, len(cats))
        colors = t.series_colors(len(cd.series))
        span = hi - lo if hi > lo else 1
        group_w = inner.w / n_cat
        # 60% of the slot is bar, 40% gap — the ratio modern chart styles use.
        bar_total = group_w * 0.60
        bar_w = bar_total if stacked else bar_total / max(1, len(cd.series))

        for ci in range(n_cat):
            base = inner.bottom
            for si, series in enumerate(cd.series):
                v = series.values[ci] if ci < len(series.values) else 0.0
                h = (v - lo) / span * inner.h
                if stacked:
                    x = inner.x + ci * group_w + (group_w - bar_total) / 2
                    y = base - h
                    base = y
                else:
                    x = inner.x + ci * group_w + (group_w - bar_total) / 2 + si * bar_w
                    y = inner.bottom - h
                self._rect(Rect(x, y, bar_w * (1 if stacked else 0.94), max(0.0, h)),
                           colors[si % len(colors)])
                if cd.data_labels and h > 1:
                    self._text(Rect(x, y - self.scale.caption * 1.5,
                                    bar_w * (1 if stacked else 0.94), self.scale.caption * 1.4),
                               _fmt_num(v), self.scale.caption, t.ink_on(t.background),
                               bold=True, align=ALIGN_CENTER, fit=False, line_spacing=1.0)

    def _draw_hbars(self, area: Rect, cd: ChartData) -> None:
        t = self.theme
        cats = cd.categories or [str(i + 1) for i in range(len(cd.series[0].values))]
        values = [v for s_ in cd.series for v in s_.values] or [0.0]
        lo, hi, step = self._nice_ticks(min(min(values), 0.0), max(values))
        label_w = max(self.fonts.width(str(c), self.scale.caption) for c in cats) + D.SPACE.sm
        tick_h = self.scale.caption * 2.0
        inner = Rect(area.x + label_w, area.y, area.w - label_w, area.h - tick_h)
        span = hi - lo if hi > lo else 1

        for v in _frange(lo, hi, step):
            x = inner.x + (v - lo) / span * inner.w
            self._rect(Rect(x - 0.4, inner.y, 0.8, inner.h), t.chart_grid)
            self._text(Rect(x - 30, inner.bottom + D.SPACE.xs, 60, self.scale.caption * 1.4),
                       _fmt_num(v), self.scale.caption, t.muted_on(t.background),
                       align=ALIGN_CENTER, fit=False, line_spacing=1.0)

        colors = t.series_colors(len(cd.series))
        row_h = inner.h / max(1, len(cats))
        bar_total = row_h * 0.60
        bar_h = bar_total / max(1, len(cd.series))
        for ci, cat in enumerate(cats):
            self._text(Rect(area.x, inner.y + ci * row_h, label_w - D.SPACE.sm, row_h),
                       str(cat), self.scale.caption, t.muted_on(t.background),
                       align=ALIGN_RIGHT, anchor=ANCHOR_MIDDLE, fit=False, line_spacing=1.0)
            for si, series in enumerate(cd.series):
                v = series.values[ci] if ci < len(series.values) else 0.0
                w = (v - lo) / span * inner.w
                y = inner.y + ci * row_h + (row_h - bar_total) / 2 + si * bar_h
                self._rect(Rect(inner.x, y, max(0.0, w), bar_h * 0.94),
                           colors[si % len(colors)])

    def _draw_lines(self, area: Rect, cd: ChartData, area_fill: bool = False) -> None:
        """Line (or filled area) chart with round point markers."""
        t = self.theme
        c = self.c
        inner, lo, hi, _ = self._axis_frame(area, cd)
        cats = cd.categories or [str(i + 1) for i in range(len(cd.series[0].values))]
        n = max(1, len(cats))
        span = hi - lo if hi > lo else 1
        colors = t.series_colors(len(cd.series))
        # Points sit at category centres, matching where the labels are drawn.
        cw = inner.w / n

        def point(ci: int, v: float) -> Tuple[float, float]:
            x = inner.x + ci * cw + cw / 2
            y = inner.bottom - (v - lo) / span * inner.h
            return x, y

        for si, series in enumerate(cd.series):
            color = colors[si % len(colors)]
            pts = [point(ci, series.values[ci] if ci < len(series.values) else 0.0)
                   for ci in range(n)]
            if area_fill and len(pts) > 1:
                self._fill(color, 0.22)
                p = c.beginPath()
                p.moveTo(pts[0][0], self._y(inner.bottom))
                for x, y in pts:
                    p.lineTo(x, self._y(y))
                p.lineTo(pts[-1][0], self._y(inner.bottom))
                p.close()
                c.drawPath(p, stroke=0, fill=1)
                c.setFillAlpha(1)

            if len(pts) > 1:
                self._stroke(color, 2.5)
                c.setLineJoin(1)   # rounded joins read as smoother
                c.setLineCap(1)
                p = c.beginPath()
                p.moveTo(pts[0][0], self._y(pts[0][1]))
                for x, y in pts[1:]:
                    p.lineTo(x, self._y(y))
                c.drawPath(p, stroke=1, fill=0)

            # Markers: a filled dot ringed in the background colour so lines
            # crossing behind a point do not visually merge with it.
            for ci, (x, y) in enumerate(pts):
                self._ellipse(Rect(x - 4.0, y - 4.0, 8.0, 8.0), color, 1.0,
                              line=t.background, line_w=1.6)
                if cd.data_labels:
                    v = series.values[ci] if ci < len(series.values) else 0.0
                    self._text(Rect(x - 26, y - self.scale.caption * 2.1, 52,
                                    self.scale.caption * 1.4),
                               _fmt_num(v), self.scale.caption, t.ink_on(t.background),
                               bold=True, align=ALIGN_CENTER, fit=False, line_spacing=1.0)

    def _draw_pie(self, area: Rect, cd: ChartData, doughnut: bool = False) -> None:
        t = self.theme
        c = self.c
        values = [max(0.0, v) for v in (cd.series[0].values if cd.series else [])]
        total = sum(values) or 1.0
        cats = cd.categories or [str(i + 1) for i in range(len(values))]
        colors = t.series_colors(len(values))

        d = min(area.w, area.h) * 0.92
        cx, cy = area.cx, self._y(area.cy)
        start = 90.0
        for i, v in enumerate(values):
            extent = -(v / total) * 360.0
            self._fill(colors[i % len(colors)])
            self._stroke(t.background, 1.5)
            c.wedge(cx - d / 2, cy - d / 2, cx + d / 2, cy + d / 2, start, extent,
                    stroke=1, fill=1)
            if cd.data_labels and abs(extent) > 12:
                mid = math.radians(start + extent / 2)
                rr = d * (0.34 if not doughnut else 0.38)
                lx, ly = cx + math.cos(mid) * rr, cy + math.sin(mid) * rr
                pct = f"{v / total * 100:.0f}%"
                self._text(Rect(lx - 26, self._y(ly) - self.scale.caption * 0.7, 52,
                                self.scale.caption * 1.4),
                           pct, self.scale.caption, t.on(colors[i % len(colors)]),
                           bold=True, align=ALIGN_CENTER, fit=False, line_spacing=1.0)
            start += extent
        if doughnut:
            # The hole is the background colour, which is why themes must
            # never draw a doughnut over a gradient.
            self._fill(t.background)
            hole = d * 0.56
            c.setStrokeAlpha(0)
            c.circle(cx, cy, hole / 2, stroke=0, fill=1)
            c.setStrokeAlpha(1)

    def _draw_legend(self, r: Rect, cd: ChartData, kind: str) -> None:
        t = self.theme
        if kind in ("pie", "doughnut", "donut"):
            labels = cd.categories or []
            colors = t.series_colors(len(labels))
        else:
            labels = [s_.name for s_ in cd.series]
            colors = t.series_colors(len(labels))
        if not labels:
            return
        chip, gap = 9.0, D.SPACE.sm
        widths = [chip + 4 + self.fonts.width(str(l), self.scale.caption) for l in labels]
        total = sum(widths) + gap * (len(labels) - 1)
        x = r.x + max(0.0, (r.w - total) / 2)
        y = r.y + (r.h - chip) / 2
        for i, label in enumerate(labels):
            self._rect(Rect(x, y, chip, chip), colors[i % len(colors)],
                       radius=2.0)
            self._text(Rect(x + chip + 4, y - 1.5, widths[i] - chip - 4,
                            self.scale.caption * 1.4),
                       str(label), self.scale.caption, t.muted_on(t.background),
                       fit=False, line_spacing=1.0)
            x += widths[i] + gap


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------


def _frange(lo: float, hi: float, step: float) -> List[float]:
    """Inclusive float range, guarded against runaway loops."""
    if step <= 0:
        return [lo]
    out: List[float] = []
    n = int(round((hi - lo) / step))
    for i in range(min(n, 40) + 1):
        out.append(lo + i * step)
    return out


def _fmt_num(v: float) -> str:
    """Compact axis label: ``1200`` -> ``1.2k``, ``3.0`` -> ``3``."""
    a = abs(v)
    if a >= 1_000_000:
        return f"{v / 1_000_000:.1f}M".replace(".0M", "M")
    if a >= 1_000:
        return f"{v / 1_000:.1f}k".replace(".0k", "k")
    if a == int(a):
        return str(int(v))
    return f"{v:.1f}"


def _today() -> str:
    import datetime as _dt

    return _dt.date.today().strftime("%B %Y")


def _retitled(s: Slide, title: str) -> Slide:
    from copy import copy

    out = copy(s)
    out.title = title
    return out


def soffice_available() -> Optional[str]:
    """Path to a LibreOffice binary, if one is installed."""
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return None


def _render_via_soffice(deck: Deck, output_path: str, theme: Optional[Theme]) -> str:
    """Render the ``.pptx`` then convert it with LibreOffice."""
    binary = soffice_available()
    if not binary:
        raise RuntimeError("LibreOffice (soffice) not found on PATH.")
    from .renderer import render_deck as render_pptx

    out_dir = os.path.dirname(os.path.abspath(output_path)) or "."
    os.makedirs(out_dir, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        pptx_path = os.path.join(tmp, "deck.pptx")
        render_pptx(deck, pptx_path, theme=theme)
        subprocess.run(
            [binary, "--headless", "--convert-to", "pdf", "--outdir", tmp, pptx_path],
            check=True, capture_output=True, timeout=180,
        )
        produced = os.path.join(tmp, "deck.pdf")
        if not os.path.exists(produced):
            raise RuntimeError("LibreOffice did not produce a PDF.")
        shutil.copyfile(produced, output_path)
    return output_path


def render_pdf(
    deck: Deck,
    output_path: str,
    theme: Optional[Theme] = None,
    engine: str = "native",
) -> str:
    """Render ``deck`` to ``output_path`` (``.pdf``) and return the path.

    ``engine``
        ``"native"`` (default) draws the PDF directly with reportlab — no
        system dependencies, and guaranteed to match the design system.
        ``"soffice"`` renders the ``.pptx`` and converts it with LibreOffice,
        which needs LibreOffice installed but reproduces the PowerPoint file
        exactly. ``"auto"`` prefers LibreOffice when present and silently
        falls back to native.
    """
    parent = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(parent, exist_ok=True)

    engine = (engine or "native").lower()
    if engine == "soffice":
        return _render_via_soffice(deck, output_path, theme)
    if engine == "auto" and soffice_available():
        try:
            return _render_via_soffice(deck, output_path, theme)
        except Exception:
            pass  # fall through to the always-available native path
    return PDFRenderer(deck, theme=theme).render(output_path)


__all__ = ["PDFRenderer", "render_pdf", "soffice_available"]
