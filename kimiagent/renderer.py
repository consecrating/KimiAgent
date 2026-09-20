"""Render a :class:`~kimiagent.models.Deck` into a native ``.pptx``.

This is where the "native depth" happens: instead of dumping text into a
template, each slide is composed from real PowerPoint primitives — shapes,
gradients, native (editable) charts and native tables — laid out on a blank
16:9 canvas by the shared design system in :mod:`kimiagent.design` and
:mod:`kimiagent.layout`.

What makes the output look designed rather than generated:

* **Every position comes from the grid.** Nothing is placed at a hand-picked
  inch offset, so titles share a baseline and cards align across slides.
* **Text is measured before it is drawn.** Type shrinks to fit its box and
  over-full slides flow onto continuation slides, so nothing overflows and
  nothing is silently dropped.
* **Depth instead of outlines.** Rounded corners, soft shadows and translucent
  overlays replace the hard 1pt borders that make default templates look like
  spreadsheets.
* **Contrast is computed.** Foreground colours are derived from whatever sits
  behind them, so no theme can produce unreadable text.

Some of this needs OOXML that ``python-pptx`` has no API for — soft shadows,
fill transparency, letter spacing, real bullet glyphs. Those are implemented
as small, well-scoped XML helpers at the top of the module rather than by
reaching into shape internals from the layout code.

Only ``python-pptx`` is required.
"""

from __future__ import annotations

import datetime as _dt
import os
from typing import List, Optional, Sequence, Tuple

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

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
# Units
# --------------------------------------------------------------------------
#
# The design system works in points; PowerPoint works in English Metric Units.
# Conversion happens only here, at the boundary.

EMU_PER_PT = 12700


def _emu(points: float) -> int:
    return int(round(points * EMU_PER_PT))


def _rgb(hex_str: str) -> RGBColor:
    return RGBColor.from_string(hex_str.lstrip("#").upper())


# --------------------------------------------------------------------------
# OOXML helpers
# --------------------------------------------------------------------------
#
# python-pptx deliberately exposes a conservative API. The effects below are
# standard DrawingML that PowerPoint, Keynote and Google Slides all honour,
# but have no Python binding, so they are written directly. Each helper is
# defensive: a failure to apply a *visual flourish* must never fail a render.


def _solid_fill_elem(fill) -> Optional[object]:
    """Return the ``<a:srgbClr>`` element of a solid fill, if present."""
    try:
        return fill.fore_color._xFill.find(qn("a:srgbClr"))
    except Exception:
        return None


def _set_fill_alpha(shape, alpha: float) -> None:
    """Apply transparency to a shape's solid fill.

    ``alpha`` is 0..1 opacity. Translucent overlays are how the covers get
    depth without needing image assets: a wash of the accent colour over a
    gradient reads as light, not as a pasted rectangle.
    """
    srgb = _solid_fill_elem(shape.fill)
    if srgb is None:
        return
    try:
        for existing in srgb.findall(qn("a:alpha")):
            srgb.remove(existing)
        el = srgb.makeelement(qn("a:alpha"), {"val": str(int(max(0.0, min(1.0, alpha)) * 100000))})
        srgb.append(el)
    except Exception:
        pass


def _add_shadow(shape, *, blur: float = 18.0, dist: float = 5.0, alpha: float = 0.16,
                direction: float = 90.0) -> None:
    """Attach a soft outer shadow.

    python-pptx can only *disable* the inherited shadow, which is why the
    default look is either a harsh preset shadow or none at all. A wide, low
    opacity shadow directly beneath a card is what reads as elevation; the
    stock PowerPoint shadow reads as a drop shadow from 2003.
    """
    try:
        spPr = shape._element.spPr
        for existing in spPr.findall(qn("a:effectLst")):
            spPr.remove(existing)
        effect = spPr.makeelement(qn("a:effectLst"), {})
        shdw = spPr.makeelement(
            qn("a:outerShdw"),
            {
                "blurRad": str(_emu(blur)),
                "dist": str(_emu(dist)),
                "dir": str(int(direction * 60000)),
                "rotWithShape": "0",
            },
        )
        color = spPr.makeelement(qn("a:srgbClr"), {"val": "000000"})
        color.append(spPr.makeelement(qn("a:alpha"), {"val": str(int(alpha * 100000))}))
        shdw.append(color)
        effect.append(shdw)
        spPr.append(effect)
    except Exception:
        pass


def _no_shadow(shape) -> None:
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def _set_letter_spacing(run, points: float) -> None:
    """Set tracking (letter spacing) on a run.

    Wide tracking on small uppercase text is the "eyebrow"/kicker convention
    from editorial design. Without it, an uppercase label just looks like
    shouting; with it, it reads as a considered label.
    """
    try:
        run.font._rPr.set("spc", str(int(points * 100)))
    except Exception:
        pass


def _set_bullet(paragraph, char: str, color: str, size_pct: int = 90, font: str = "Arial") -> None:
    """Give a paragraph a real PowerPoint bullet glyph in ``color``.

    Using a native bullet rather than a character typed into the text is what
    produces a correct hanging indent: wrapped lines align under the text,
    not under the glyph. Elements are appended in schema order (bullet
    properties follow spacing properties in ``a:pPr``).
    """
    try:
        pPr = paragraph._pPr
        if pPr is None:
            pPr = paragraph._p.get_or_add_pPr()
        for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buFont", "a:buSzPct"):
            for existing in pPr.findall(qn(tag)):
                pPr.remove(existing)
        clr = pPr.makeelement(qn("a:buClr"), {})
        srgb = pPr.makeelement(qn("a:srgbClr"), {"val": color.lstrip("#").upper()})
        clr.append(srgb)
        pPr.append(clr)
        pPr.append(pPr.makeelement(qn("a:buSzPct"), {"val": str(int(size_pct * 1000))}))
        pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": font}))
        pPr.append(pPr.makeelement(qn("a:buChar"), {"char": char}))
    except Exception:
        pass


def _set_indent(paragraph, margin_pt: float, hang_pt: float) -> None:
    """Set the paragraph's left margin and first-line hanging indent."""
    try:
        pPr = paragraph._pPr
        if pPr is None:
            pPr = paragraph._p.get_or_add_pPr()
        pPr.set("marL", str(_emu(margin_pt)))
        pPr.set("indent", str(-_emu(hang_pt)))
    except Exception:
        pass


def _gradient(fill, start: str, end: str, angle: float) -> None:
    """Apply a two-stop linear gradient to a fill."""
    fill.gradient()
    stops = fill.gradient_stops
    stops[0].color.rgb = _rgb(start)
    stops[0].position = 0.0
    stops[1].color.rgb = _rgb(end)
    stops[1].position = 1.0
    try:
        fill.gradient_angle = angle
    except Exception:
        pass


# --------------------------------------------------------------------------
# Renderer
# --------------------------------------------------------------------------


class DeckRenderer:
    """Renders a deck to a :class:`pptx.Presentation`.

    One instance per deck keeps the theme, grid and type scale to hand so the
    slide builders stay free of bookkeeping.
    """

    def __init__(self, deck: Deck, theme: Optional[Theme] = None) -> None:
        self.deck = deck
        self.theme: Theme = theme or get_theme(deck.theme)

        # Canvas + grid. A 4:3 deck is narrower, so the grid adapts rather
        # than the layouts needing a second code path.
        if (deck.aspect or "16:9").replace(" ", "") in ("4:3", "4x3"):
            width, height = D.SLIDE_W_4_3, D.SLIDE_H_4_3
        else:
            width, height = D.SLIDE_W, D.SLIDE_H
        self.grid = Grid(width=width, height=height)
        self.scale: TypeScale = TYPE.scaled(DENSITY.get(deck.density or "normal", 1.0))

        self.prs = Presentation()
        self.prs.slide_width = _emu(width)
        self.prs.slide_height = _emu(height)
        self._blank = self.prs.slide_layouts[6]  # fully blank layout

    # -- entry point ------------------------------------------------------

    def render(self) -> Presentation:
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

        # Flow content before drawing, so nothing overflows or is dropped.
        flowed = paginate(self.deck, self.grid, self.scale, self.theme.body_serif)
        total = len(flowed.slides)

        for idx, slide in enumerate(flowed.slides):
            builder = dispatch.get(slide.type, self._bullets)
            pptx_slide = self.prs.slides.add_slide(self._blank)
            builder(pptx_slide, slide)
            self._add_notes(pptx_slide, slide.notes)
            if slide.type not in FULL_BLEED_TYPES:
                self._footer(pptx_slide, idx + 1, total)
        return self.prs

    # ------------------------------------------------------------------
    # Primitives
    # ------------------------------------------------------------------

    def _bg(self, slide, color: str) -> None:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = _rgb(color)

    def _bg_gradient(self, slide) -> Tuple[str, str]:
        """Fill the slide with the theme's text-safe hero gradient."""
        start, end, angle = self.theme.gradient
        _gradient(slide.background.fill, start, end, angle)
        return start, end

    def _shape(
        self,
        slide,
        rect: Rect,
        shape_type=MSO_SHAPE.RECTANGLE,
        fill: Optional[str] = None,
        *,
        alpha: float = 1.0,
        line: Optional[str] = None,
        line_w: float = 1.0,
        radius: Optional[float] = None,
        shadow: bool = False,
    ):
        """Draw a shape from a :class:`Rect`."""
        shp = slide.shapes.add_shape(
            shape_type, _emu(rect.x), _emu(rect.y), _emu(rect.w), _emu(rect.h)
        )
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = _rgb(fill)
            if alpha < 1.0:
                _set_fill_alpha(shp, alpha)
        else:
            shp.fill.background()

        if line:
            shp.line.color.rgb = _rgb(line)
            shp.line.width = Pt(line_w)
        else:
            shp.line.fill.background()

        # Corner radius is expressed as a fraction of the shorter side, so a
        # tall card and a wide card end up with the *same* visual radius.
        if radius is not None and shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
            short = max(1.0, min(rect.w, rect.h))
            try:
                shp.adjustments[0] = max(0.0, min(0.5, radius / short))
            except Exception:
                pass

        _no_shadow(shp)
        if shadow:
            _add_shadow(shp)
        return shp

    def _card(self, slide, rect: Rect, *, fill: Optional[str] = None, accent: Optional[str] = None):
        """A surface panel, styled per the theme's shape language.

        Themes choose elevation (shadow) *or* definition (hairline), never
        both — stacking the two is what makes a card look like a bug report.
        """
        t = self.theme
        fill = fill or t.surface
        shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if t.radius > 0 else MSO_SHAPE.RECTANGLE
        shp = self._shape(
            slide,
            rect,
            shape_type,
            fill,
            line=t.border if t.hairline else None,
            line_w=0.75,
            radius=t.radius if t.radius > 0 else None,
            shadow=t.shadow,
        )
        return shp

    def _text(
        self,
        slide,
        rect: Rect,
        text: str,
        size: float,
        color: str,
        *,
        bold: bool = False,
        italic: bool = False,
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP,
        font: Optional[str] = None,
        serif: bool = False,
        line_spacing: Optional[float] = None,
        fit: bool = True,
        min_size: float = 8.0,
        tracking: Optional[float] = None,
        caps: bool = False,
    ) -> float:
        """Draw a text box, shrinking the type if needed. Returns height used.

        Internal margins are zeroed so the measured width equals the box
        width — without that, :func:`~kimiagent.design.fit_font_size` would be
        measuring against the wrong column and text would still overflow by
        the margin amount.
        """
        if text is None:
            return 0.0
        if caps:
            text = text.upper()

        lh = line_spacing or self.scale.line_height
        if fit:
            size = D.fit_font_size(
                text, rect.w, rect.h, ideal=size, minimum=min_size,
                bold=bold, serif=serif, line_height=lh,
            )
            # If it still will not fit at the floor size, trim rather than spill.
            text = D.truncate_to_fit(text, rect.w, rect.h, size, bold, serif, lh)

        measured = D.text_height(text, rect.w, size, bold, serif, lh)

        # Top-anchored boxes are tightened to their content. Callers stack by
        # adding the returned height to a running y, so a box left at its
        # requested height would leave an invisible dead zone that the next
        # element then overlaps. Middle/bottom anchoring needs the full
        # height, because that is what the text is being centred within.
        box_h = rect.h
        if anchor == MSO_ANCHOR.TOP:
            box_h = max(size * lh, min(rect.h, measured))

        box = slide.shapes.add_textbox(_emu(rect.x), _emu(rect.y), _emu(rect.w), _emu(box_h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = anchor

        p = tf.paragraphs[0]
        p.alignment = align
        p.line_spacing = lh
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = _rgb(color)
        run.font.name = font or self.theme.body_font
        if tracking:
            _set_letter_spacing(run, tracking)

        return measured

    def _heading_font(self) -> str:
        return self.theme.heading_font

    # -- standard slide furniture ----------------------------------------

    def _zones(self, s: Slide) -> Zones:
        return zones(
            self.grid,
            self.scale,
            has_kicker=bool(s.kicker),
            has_subtitle=bool(s.subtitle),
            title_lines=title_line_count(s.title, self.grid, self.scale, self.theme.heading_serif),
        )

    def _heading(self, slide, s: Slide, *, on: Optional[str] = None) -> Zones:
        """Draw kicker + title + accent rule + subtitle. Returns the zones."""
        t = self.theme
        bg = on or t.background
        z = self._zones(s)

        if s.kicker:
            self._text(
                slide, z.kicker, s.kicker, self.scale.caption,
                t.accent_on(bg), bold=True, caps=True, tracking=1.4,
                font=t.body_font, fit=False,
            )
        if s.title:
            self._text(
                slide, z.title, s.title, self.scale.h1, t.ink_on(bg),
                bold=t.heading_bold, font=self._heading_font(), serif=t.heading_serif,
                line_spacing=self.scale.line_height_tight,
            )
        # A short rule rather than a full-width divider: it anchors the title
        # without cutting the slide in half.
        self._shape(slide, z.rule, MSO_SHAPE.RECTANGLE, self._accent_for(s))
        if s.subtitle:
            self._text(
                slide, z.subtitle, s.subtitle, self.scale.body_sm,
                t.muted_on(bg), font=t.body_font, serif=t.body_serif,
            )
        return z

    def _accent_for(self, s: Slide, index: int = 0) -> str:
        """The highlight colour for a slide, honouring a spec override."""
        if s.accent is not None:
            return self.theme.accent(s.accent + index)
        return self.theme.accent(index) if index else self.theme.secondary

    def _footer(self, slide, page: int, total: int) -> None:
        """A minimal footer plus a thin progress rule.

        The progress rule is the one piece of chrome worth its ink: it tells
        an audience how much is left, which is the question they are actually
        asking during a long deck.
        """
        t = self.theme
        g = self.grid
        z = zones(g, self.scale)
        # A long deck title would wrap onto a second line and push out of the
        # footer, so it is clipped to one line rather than shrunk: the footer's
        # small size is intentional.
        label_w = z.footer.w * 0.7
        label = D.truncate_to_width(
            self.deck.footer_note or self.deck.title, label_w, self.scale.micro
        )
        self._text(
            slide, Rect(z.footer.x, z.footer.y + 4, label_w, z.footer.h),
            label, self.scale.micro, t.muted_on(t.background), font=t.body_font, fit=False,
        )
        self._text(
            slide,
            Rect(g.right - 90, z.footer.y + 4, 90, z.footer.h),
            f"{page} / {total}", self.scale.micro, t.muted_on(t.background),
            align=PP_ALIGN.RIGHT, font=t.body_font, fit=False,
        )

        # Progress rule, flush to the bottom edge.
        y = g.height - 3.0
        self._shape(slide, Rect(0, y, g.width, 3.0), MSO_SHAPE.RECTANGLE,
                    t.border, alpha=0.55)
        if total > 0:
            self._shape(slide, Rect(0, y, g.width * (page / total), 3.0),
                        MSO_SHAPE.RECTANGLE, t.secondary)

    def _add_notes(self, slide, notes: str) -> None:
        if not notes:
            return
        try:
            slide.notes_slide.notes_text_frame.text = notes
        except Exception:
            pass

    # -- decorative ------------------------------------------------------

    def _hero_decor(self, slide, accent: str) -> None:
        """Large, soft geometry on a hero slide.

        Two overlapping translucent circles bleeding off the corner. This is
        doing real work: a flat colour field with centred text reads as a
        placeholder, whereas an off-centre asymmetric form gives the eye
        somewhere to enter the slide.
        """
        g = self.grid
        r1 = min(g.width, g.height) * 0.62
        self._shape(
            slide, Rect(g.width - r1 * 0.55, g.height - r1 * 0.62, r1, r1),
            MSO_SHAPE.OVAL, accent, alpha=0.14,
        )
        r2 = r1 * 0.58
        self._shape(
            slide, Rect(g.width - r2 * 0.30, g.height - r2 * 1.35, r2, r2),
            MSO_SHAPE.OVAL, accent, alpha=0.20,
        )

    # ------------------------------------------------------------------
    # Slide builders
    # ------------------------------------------------------------------

    def _cover(self, slide, s: Slide) -> None:
        t, g = self.theme, self.grid
        self._bg_gradient(slide)
        self._hero_decor(slide, t.secondary)
        ink = t.hero_text
        muted = D.mix(ink, D.mix(*t.gradient[:2], 0.5), 0.28)

        # A hero gets more air than a content slide: the title starts below
        # the optical centre so the eye lands on it rather than scanning.
        left = g.left
        width = g.content_w * 0.82

        y = g.height * 0.30
        kicker = s.kicker or self.deck.subtitle or ""
        if kicker:
            h = self._text(
                slide, Rect(left, y, width, 22), kicker, self.scale.caption,
                ink, bold=True, caps=True, tracking=1.8, font=t.body_font, fit=False,
            )
            y += h + D.SPACE.sm

        title = s.title or self.deck.title
        title_box = Rect(left, y, width, self.scale.display * 2.6)
        used = self._text(
            slide, title_box, title, self.scale.display, ink,
            bold=True, font=self._heading_font(), serif=t.heading_serif,
            line_spacing=self.scale.line_height_tight, min_size=self.scale.h2,
        )
        y += used + D.SPACE.md

        self._shape(slide, Rect(left, y, 62, 4), MSO_SHAPE.RECTANGLE, t.secondary)
        y += 4 + D.SPACE.md

        if s.subtitle:
            y += self._text(
                slide, Rect(left, y, width * 0.88, self.scale.h3 * 3),
                s.subtitle, self.scale.h3, muted, font=t.body_font, serif=t.body_serif,
            )

        # Byline pinned to the baseline, not floating after the subtitle.
        byline = " · ".join(x for x in (self.deck.author, _today()) if x)
        if byline:
            self._text(
                slide, Rect(left, g.bottom - 24, width, 22), byline,
                self.scale.body_sm, muted, font=t.body_font, fit=False,
            )

    def _closing(self, slide, s: Slide) -> None:
        t, g = self.theme, self.grid
        self._bg_gradient(slide)
        self._hero_decor(slide, t.secondary)
        ink = t.hero_text

        box_w = g.content_w * 0.78
        x = (g.width - box_w) / 2
        y = g.height * 0.34

        used = self._text(
            slide, Rect(x, y, box_w, self.scale.hero * 2.4),
            s.title or "Thank you", self.scale.hero, ink, bold=True,
            align=PP_ALIGN.CENTER, font=self._heading_font(), serif=t.heading_serif,
            line_spacing=self.scale.line_height_tight,
        )
        y += used + D.SPACE.md
        self._shape(slide, Rect(g.width / 2 - 31, y, 62, 4), MSO_SHAPE.RECTANGLE, t.secondary)
        y += 4 + D.SPACE.md
        if s.subtitle:
            self._text(
                slide, Rect(x, y, box_w, self.scale.h3 * 3), s.subtitle,
                self.scale.h3, D.mix(ink, D.mix(*t.gradient[:2], 0.5), 0.25),
                align=PP_ALIGN.CENTER, font=t.body_font, serif=t.body_serif,
            )

    def _section(self, slide, s: Slide) -> None:
        """A divider carried by an oversized numeral.

        The number is set very large and very low contrast so it works as
        texture rather than information — it signals "new chapter" at a glance
        without competing with the section title.
        """
        t, g = self.theme, self.grid
        self._bg(slide, t.primary if not t.dark else t.surface)
        bg = t.primary if not t.dark else t.surface
        ink = t.on(bg)
        accent = self._accent_for(s)

        number = (s.subtitle or "").strip()
        if number:
            # Drawn directly in a low-contrast tint of the ink so it reads as
            # texture. Tinting the colour is safer than overlaying a scrim,
            # which would also wash out anything else it happened to cover.
            self._text(
                slide, Rect(g.right - 340, g.height * 0.16, 340, 260),
                number, 190, D.mix(bg, ink, 0.16), bold=True, align=PP_ALIGN.RIGHT,
                font=self._heading_font(), serif=t.heading_serif, fit=False,
                line_spacing=1.0,
            )

        y = g.height * 0.40
        # The eyebrow uses ``kicker``, never ``subtitle`` — the subtitle is
        # already shown as the oversized numeral, and printing it twice was
        # both redundant and a collision waiting to happen.
        if s.kicker:
            self._text(
                slide, Rect(g.left, y - 30, g.content_w, 20),
                s.kicker, self.scale.caption, accent, bold=True, caps=True,
                tracking=1.6, font=t.body_font, fit=False,
            )
        self._shape(slide, Rect(g.left, y, 62, 4), MSO_SHAPE.RECTANGLE, accent)
        y += 4 + D.SPACE.lg
        self._text(
            slide, Rect(g.left, y, g.content_w * 0.74, self.scale.hero * 2.2),
            s.title, self.scale.hero, ink, bold=True,
            font=self._heading_font(), serif=t.heading_serif,
            line_spacing=self.scale.line_height_tight,
        )

    def _big_statement(self, slide, s: Slide) -> None:
        """One claim, given the whole slide.

        The archetype exists because the most important sentence in a deck
        should not be bullet three of seven.
        """
        t, g = self.theme, self.grid
        self._bg_gradient(slide)
        ink = t.hero_text
        accent = self._accent_for(s)

        text = s.title or s.quote
        box_w = g.content_w * 0.86
        x = (g.width - box_w) / 2

        if s.kicker:
            self._text(
                slide, Rect(x, g.height * 0.22, box_w, 22), s.kicker,
                self.scale.caption, ink, bold=True, caps=True, tracking=1.8,
                align=PP_ALIGN.CENTER, font=t.body_font, fit=False,
            )

        # Centred in the available band, with the rule and sub-line placed
        # against where the text *actually* ends. Pinning them to the bottom of
        # the band instead leaves a gap that grows as the statement gets
        # shorter — exactly backwards, since a short statement is the one that
        # most needs to feel deliberate.
        band = Rect(x, g.height * 0.28, box_w, g.height * 0.44)
        used = self._text(
            slide, band, text, self.scale.hero * 1.12, ink, bold=True,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            font=self._heading_font(), serif=t.heading_serif,
            line_spacing=self.scale.line_height_tight, min_size=self.scale.h2,
        )
        y = band.y + (band.h + used) / 2 + D.SPACE.lg
        self._shape(slide, Rect(g.width / 2 - 31, y, 62, 4), MSO_SHAPE.RECTANGLE, accent)
        if s.subtitle or s.attribution:
            self._text(
                slide, Rect(x, y + 4 + D.SPACE.md, box_w, 40),
                s.subtitle or f"— {s.attribution}",
                self.scale.body, D.mix(ink, D.mix(*t.gradient[:2], 0.5), 0.25),
                align=PP_ALIGN.CENTER, font=t.body_font, serif=t.body_serif,
            )

    # -- prose -----------------------------------------------------------

    def _bullets(self, slide, s: Slide) -> None:
        self._bg(slide, self.theme.background)
        z = self._heading(slide, s)
        self._render_bullets(slide, s.bullets, z.body, accent=self._accent_for(s))

    def _render_bullets(
        self,
        slide,
        bullets: Sequence[Bullet],
        rect: Rect,
        *,
        accent: Optional[str] = None,
        on: Optional[str] = None,
        size: Optional[float] = None,
        compact: bool = False,
    ) -> float:
        """Draw a bullet list with native glyphs and hanging indents."""
        if not bullets:
            return 0.0
        t = self.theme
        bg = on or t.background
        accent = accent or t.secondary
        base = size or self.scale.body
        gap = D.SPACE.xs if compact else D.SPACE.sm

        box = slide.shapes.add_textbox(_emu(rect.x), _emu(rect.y), _emu(rect.w), _emu(rect.h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        # One shared size for the whole list keeps the block even; sizing each
        # bullet independently makes a list look ransom-noted.
        sizes = [base if b.level == 0 else base * 0.87 for b in bullets]
        total = 0.0
        for b, sz in zip(bullets, sizes):
            indent = 16.0 + b.level * 16.0
            text = f"{b.bold_lead}: {b.text}" if b.bold_lead else b.text
            total += D.text_height(text, rect.w - indent, sz, serif=t.body_serif) + gap
        shrink = 1.0
        while total > rect.h and shrink > 0.62:
            shrink -= 0.04
            total = 0.0
            for b, sz in zip(bullets, sizes):
                indent = 16.0 + b.level * 16.0
                text = f"{b.bold_lead}: {b.text}" if b.bold_lead else b.text
                total += D.text_height(text, rect.w - indent, sz * shrink, serif=t.body_serif) + gap

        first = True
        for b, sz in zip(bullets, sizes):
            sz *= shrink
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.line_spacing = self.scale.line_height
            p.space_after = Pt(gap)
            p.alignment = PP_ALIGN.LEFT

            marker = "\u2022" if b.level == 0 else "\u2013"
            margin = 16.0 + b.level * 16.0
            _set_indent(p, margin, 16.0)
            _set_bullet(p, marker, accent if b.level == 0 else t.muted_on(bg),
                        size_pct=100 if b.level == 0 else 90)

            if b.bold_lead:
                lead = p.add_run()
                lead.text = f"{b.bold_lead}: "
                lead.font.size = Pt(sz)
                lead.font.bold = True
                lead.font.color.rgb = _rgb(t.ink_on(bg))
                lead.font.name = t.body_font
            run = p.add_run()
            run.text = b.text
            run.font.size = Pt(sz)
            run.font.color.rgb = _rgb(t.ink_on(bg) if b.level == 0 else t.muted_on(bg))
            run.font.name = t.body_font
        return total

    def _two_column(self, slide, s: Slide) -> None:
        t = self.theme
        self._bg(slide, t.background)
        z = self._heading(slide, s)
        cols = (s.columns + [[], []])[:2]
        rects = z.body.split_h(2, D.SPACE.md)
        # Size the cards to their content rather than to the whole body zone,
        # so two short columns do not leave half the slide as empty panel.
        card_h = column_card_height(cols, rects[0].w, z.body, self.scale, t.body_serif)
        for i, (col, r) in enumerate(zip(cols, rects)):
            r = r.with_h(card_h)
            self._card(slide, r)
            accent = self._accent_for(s, i)
            # A colour bar on the inner edge, not a full header block: it keys
            # the column to a colour without spending vertical space.
            self._shape(slide, Rect(r.x, r.y, r.w, 4), MSO_SHAPE.RECTANGLE, accent)
            self._render_bullets(
                slide, col,
                Rect(r.x + D.SPACE.lg, r.y + D.SPACE.lg + 4,
                     r.w - D.SPACE.lg * 2, r.h - D.SPACE.lg * 2 - 4),
                accent=accent, on=t.surface,
            )

    def _quote(self, slide, s: Slide) -> None:
        t, g = self.theme, self.grid
        self._bg(slide, t.surface)
        accent = self._accent_for(s)

        # The quote mark is decoration, so it is set huge and translucent and
        # allowed to sit behind the text rather than beside it.
        self._text(
            slide, Rect(g.left - 6, g.top + 6, 200, 220), "\u201C", 200,
            D.mix(t.surface, accent, 0.30), bold=True,
            font=self._heading_font(), serif=True, fit=False, line_spacing=1.0,
        )
        box = Rect(g.left + 54, g.height * 0.24, g.content_w - 100, g.height * 0.42)
        used = self._text(
            slide, box, s.quote or s.title, self.scale.h1 * 1.05, t.ink_on(t.surface),
            italic=False, bold=True, anchor=MSO_ANCHOR.MIDDLE,
            font=self._heading_font(), serif=t.heading_serif,
            line_spacing=self.scale.line_height_tight, min_size=self.scale.h3,
        )
        # Attribution follows the quote, not the bottom of its band.
        y = box.y + (box.h + used) / 2 + D.SPACE.lg
        self._shape(slide, Rect(box.x, y, 40, 3), MSO_SHAPE.RECTANGLE, accent)
        if s.attribution:
            self._text(
                slide, Rect(box.x, y + 3 + D.SPACE.md, box.w, 40), s.attribution,
                self.scale.body_sm, t.muted_on(t.surface), caps=True, tracking=1.2,
                bold=True, font=t.body_font, fit=False,
            )

    # -- structured ------------------------------------------------------

    def _stats(self, slide, s: Slide) -> None:
        t = self.theme
        self._bg(slide, t.background)
        z = self._heading(slide, s)
        if not s.stats:
            return
        rects = stat_grid(z.body, len(s.stats), D.SPACE.md)
        for i, (st, r) in enumerate(zip(s.stats, rects)):
            self._stat_card(slide, st, r, i, s)

    def _stat_card(self, slide, st: Stat, r: Rect, i: int, s: Slide) -> None:
        t = self.theme
        accent = self._accent_for(s, i) if s.accent is not None else t.accent(i)
        self._card(slide, r)
        self._shape(slide, Rect(r.x, r.y, r.w, 4), MSO_SHAPE.RECTANGLE, accent)

        inner = r.inset(D.SPACE.md, D.SPACE.md)
        # The number is the point of the slide, so it gets display size and
        # the accent colour, guaranteed legible as large text.
        value_h = inner.h * 0.46
        self._text(
            slide, Rect(inner.x, r.y + D.SPACE.lg, inner.w, value_h), st.value,
            self.scale.display * 0.86, D.ensure_contrast(accent, t.surface, 3.0),
            bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            font=self._heading_font(), serif=t.heading_serif, line_spacing=1.0,
            min_size=self.scale.h2,
        )

        y = r.y + D.SPACE.lg + value_h + D.SPACE.xs
        if st.trend:
            self._trend_glyph(slide, Rect(inner.cx - 6, y, 12, 10), st.trend, accent)
            y += 14

        y += self._text(
            slide, Rect(inner.x, y, inner.w, 44), st.label, self.scale.body_sm,
            t.ink_on(t.surface), bold=True, align=PP_ALIGN.CENTER, font=t.body_font,
            serif=t.body_serif,
        )
        if st.detail:
            self._text(
                slide, Rect(inner.x, y + D.SPACE.xxs, inner.w, max(14.0, r.bottom - y - D.SPACE.md)),
                st.detail, self.scale.caption, t.muted_on(t.surface),
                align=PP_ALIGN.CENTER, font=t.body_font, serif=t.body_serif,
            )

    def _trend_glyph(self, slide, r: Rect, trend: str, accent: str) -> None:
        """A small triangle/bar cueing direction.

        Drawn as a shape rather than typed as an arrow character, because
        arrow glyphs are missing from enough font fallbacks to be unreliable.
        """
        trend = (trend or "").lower()
        if trend in ("up", "rise", "increase", "+"):
            self._shape(slide, r, MSO_SHAPE.ISOSCELES_TRIANGLE, accent)
        elif trend in ("down", "fall", "decrease", "-"):
            shp = self._shape(slide, r, MSO_SHAPE.ISOSCELES_TRIANGLE, accent)
            try:
                shp.rotation = 180
            except Exception:
                pass
        else:
            self._shape(
                slide, Rect(r.x, r.cy - 1.5, r.w, 3), MSO_SHAPE.RECTANGLE,
                self.theme.muted_on(self.theme.surface),
            )

    def _cards(self, slide, s: Slide) -> None:
        t = self.theme
        self._bg(slide, t.background)
        z = self._heading(slide, s)
        if not s.cards:
            return
        rects = card_grid(z.body, len(s.cards), D.SPACE.md)
        for i, (card, r) in enumerate(zip(s.cards, rects)):
            accent = t.accent(card.accent) if card.accent is not None else t.accent(i)
            self._card(slide, r)
            inner = r.inset(D.SPACE.lg, D.SPACE.lg)
            y = inner.y

            if card.icon:
                # A tinted rounded tile behind the glyph, which reads as an
                # icon rather than as a stray character.
                tile = Rect(inner.x, y, 34, 34)
                self._shape(
                    slide, tile,
                    MSO_SHAPE.ROUNDED_RECTANGLE if t.radius else MSO_SHAPE.RECTANGLE,
                    D.mix(t.surface, accent, 0.20), radius=9 if t.radius else None,
                )
                self._text(
                    slide, tile, card.icon, 15,
                    D.ensure_contrast(accent, D.mix(t.surface, accent, 0.20), 3.0),
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                    font=t.body_font, fit=False, line_spacing=1.0,
                )
                y += tile.h + D.SPACE.sm
            else:
                self._shape(slide, Rect(inner.x, y + 4, 26, 3), MSO_SHAPE.RECTANGLE, accent)
                y += D.SPACE.md

            y += self._text(
                slide, Rect(inner.x, y, inner.w, self.scale.h3 * 2.6), card.title,
                self.scale.h3, t.ink_on(t.surface), bold=True,
                font=self._heading_font(), serif=t.heading_serif,
                line_spacing=self.scale.line_height_tight,
            )
            if card.body:
                self._text(
                    slide, Rect(inner.x, y + D.SPACE.xs, inner.w,
                                max(16.0, inner.bottom - y - D.SPACE.xs)),
                    card.body, self.scale.body_sm, t.muted_on(t.surface),
                    font=t.body_font, serif=t.body_serif,
                )

    def _timeline(self, slide, s: Slide) -> None:
        """Milestones on a horizontal axis, labels alternating above and below.

        Alternating sides is what lets five milestones carry real detail: on
        one side each label would be a narrow column of wrapped words.
        """
        t = self.theme
        self._bg(slide, t.background)
        z = self._heading(slide, s)
        items = s.timeline
        if not items:
            return

        axis_y = z.body.y + z.body.h * 0.5
        n = len(items)
        # Inset the first and last nodes so end labels are not clipped.
        usable = z.body.w * (1 - 1.0 / max(2, n))
        x0 = z.body.x + (z.body.w - usable) / 2
        step = usable / max(1, n - 1) if n > 1 else 0

        self._shape(slide, Rect(z.body.x, axis_y - 1, z.body.w, 2),
                    MSO_SHAPE.RECTANGLE, t.border)
        # Completed portion of the axis, drawn over the full rule.
        done = [i for i, it in enumerate(items) if it.done]
        if done:
            end_x = x0 + step * max(done)
            self._shape(slide, Rect(z.body.x, axis_y - 1.5, end_x - z.body.x, 3),
                        MSO_SHAPE.RECTANGLE, t.secondary)

        label_h = z.body.h * 0.40
        col_w = max(90.0, (step if n > 1 else z.body.w) * 1.5)
        for i, it in enumerate(items):
            cx = x0 + step * i if n > 1 else z.body.cx
            accent = t.accent(i)
            above = i % 2 == 0

            node = 15.0
            self._shape(slide, Rect(cx - node / 2, axis_y - node / 2, node, node),
                        MSO_SHAPE.OVAL, accent if it.done else t.background,
                        line=accent, line_w=2.0)
            if it.done:
                self._shape(slide, Rect(cx - 3, axis_y - 3, 6, 6), MSO_SHAPE.OVAL,
                            t.on(accent))

            stem = 16.0
            sy = axis_y - node / 2 - stem if above else axis_y + node / 2
            self._shape(slide, Rect(cx - 0.75, sy, 1.5, stem), MSO_SHAPE.RECTANGLE, t.border)

            bx = cx - col_w / 2
            bx = max(z.body.x, min(bx, z.body.right - col_w))
            by = sy - label_h if above else sy + stem
            block = Rect(bx, by, col_w, label_h)

            align = PP_ALIGN.CENTER
            y = block.y
            if above:
                # Bottom-align the block against the axis so the label sits
                # just above its stem regardless of how tall it is.
                y = block.bottom - self._measure_timeline(it, block.w)

            if it.period:
                y += self._text(
                    slide, Rect(block.x, y, block.w, 18), it.period, self.scale.caption,
                    D.ensure_contrast(accent, t.background, 3.0), bold=True, caps=True,
                    tracking=1.0, align=align, font=t.body_font, fit=False,
                )
                y += D.SPACE.xxs
            y += self._text(
                slide, Rect(block.x, y, block.w, self.scale.h3 * 2.4), it.title,
                self.scale.body, t.ink_on(t.background), bold=True, align=align,
                font=self._heading_font(), serif=t.heading_serif,
                line_spacing=self.scale.line_height_tight,
            )
            if it.detail:
                self._text(
                    slide, Rect(block.x, y + D.SPACE.xxs, block.w,
                                max(14.0, block.bottom - y - D.SPACE.xxs)),
                    it.detail, self.scale.caption, t.muted_on(t.background),
                    align=align, font=t.body_font, serif=t.body_serif,
                )

    def _measure_timeline(self, it: TimelineItem, w: float) -> float:
        """Height of a timeline label block, for bottom-aligning it."""
        h = 0.0
        if it.period:
            h += 18 + D.SPACE.xxs
        h += D.text_height(it.title, w, self.scale.body, bold=True,
                           serif=self.theme.heading_serif,
                           line_height=self.scale.line_height_tight)
        if it.detail:
            h += D.SPACE.xxs + D.text_height(it.detail, w, self.scale.caption,
                                             serif=self.theme.body_serif)
        return h

    def _process(self, slide, s: Slide) -> None:
        """A numbered sequence with chevron connectors."""
        t = self.theme
        self._bg(slide, t.background)
        z = self._heading(slide, s)
        steps = s.steps
        if not steps:
            return

        n = len(steps)
        gap = D.SPACE.md
        arrow = 16.0
        avail = z.body.w - (n - 1) * (gap + arrow)
        w = avail / n
        h = min(z.body.h, 210.0)
        top = z.body.y + (z.body.h - h) / 2

        for i, st in enumerate(steps):
            x = z.body.x + i * (w + gap + arrow)
            r = Rect(x, top, w, h)
            accent = t.accent(i)
            self._card(slide, r)

            inner = r.inset(D.SPACE.md, D.SPACE.md)
            # Numbered disc: the sequence has to be readable at a glance, and
            # a number does that better than position alone.
            d = 38.0
            disc = Rect(inner.x, inner.y, d, d)
            self._shape(slide, disc, MSO_SHAPE.OVAL, accent)
            # Centred within the disc itself rather than offset by a magic
            # number, so the numeral stays centred at any type scale.
            self._text(
                slide, disc, str(i + 1), self.scale.h3, t.on(accent), bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                font=self._heading_font(), fit=False, line_spacing=1.0,
            )
            y = disc.bottom + D.SPACE.sm
            y += self._text(
                slide, Rect(inner.x, y, inner.w, self.scale.h3 * 2.4), st.title,
                self.scale.h3, t.ink_on(t.surface), bold=True,
                font=self._heading_font(), serif=t.heading_serif,
                line_spacing=self.scale.line_height_tight,
            )
            if st.detail:
                self._text(
                    slide, Rect(inner.x, y + D.SPACE.xs, inner.w,
                                max(14.0, inner.bottom - y - D.SPACE.xs)),
                    st.detail, self.scale.body_sm, t.muted_on(t.surface),
                    font=t.body_font, serif=t.body_serif,
                )

            if i < n - 1:
                ax = r.right + gap / 2
                self._shape(
                    slide, Rect(ax, top + h / 2 - arrow / 2, arrow, arrow),
                    MSO_SHAPE.ISOSCELES_TRIANGLE, t.border,
                ).rotation = 90

    def _comparison(self, slide, s: Slide) -> None:
        """Two options weighed, with the recommendation made explicit."""
        t = self.theme
        self._bg(slide, t.background)
        z = self._heading(slide, s)
        sides = s.sides[:2]
        if not sides:
            return

        body = z.body
        verdict_h = 0.0
        if s.verdict:
            verdict_h = 54.0
            body = Rect(body.x, body.y, body.w, body.h - verdict_h - D.SPACE.md)

        rects = body.split_h(len(sides), D.SPACE.lg)
        for i, (side, r) in enumerate(zip(sides, rects)):
            accent = t.accent(i)
            if side.preferred:
                # The recommended option is lifted, not just labelled: a tinted
                # surface plus a full border makes the choice legible instantly.
                self._shape(
                    slide, r,
                    MSO_SHAPE.ROUNDED_RECTANGLE if t.radius else MSO_SHAPE.RECTANGLE,
                    D.mix(t.surface, accent, 0.10), line=accent, line_w=2.0,
                    radius=t.radius if t.radius else None, shadow=t.shadow,
                )
                surface = D.mix(t.surface, accent, 0.10)
            else:
                self._card(slide, r)
                surface = t.surface

            inner = r.inset(D.SPACE.lg, D.SPACE.lg)
            y = inner.y

            if side.preferred:
                badge_w, badge_h = 108.0, 20.0
                self._shape(
                    slide, Rect(inner.x, y, badge_w, badge_h),
                    MSO_SHAPE.ROUNDED_RECTANGLE if t.radius else MSO_SHAPE.RECTANGLE,
                    accent, radius=badge_h / 2 if t.radius else None,
                )
                self._text(
                    slide, Rect(inner.x, y + 4.5, badge_w, 14), "RECOMMENDED",
                    self.scale.micro, t.on(accent), bold=True, tracking=1.0,
                    align=PP_ALIGN.CENTER, font=t.body_font, fit=False,
                )
                y += badge_h + D.SPACE.sm

            y += self._text(
                slide, Rect(inner.x, y, inner.w, self.scale.h2 * 2.2), side.title,
                self.scale.h2, t.ink_on(surface), bold=True,
                font=self._heading_font(), serif=t.heading_serif,
                line_spacing=self.scale.line_height_tight,
            )
            if side.subtitle:
                y += self._text(
                    slide, Rect(inner.x, y + D.SPACE.xxs, inner.w, 34), side.subtitle,
                    self.scale.body_sm, t.muted_on(surface), font=t.body_font,
                    serif=t.body_serif,
                ) + D.SPACE.xxs
            y += D.SPACE.sm
            self._shape(slide, Rect(inner.x, y, 28, 3), MSO_SHAPE.RECTANGLE, accent)
            y += 3 + D.SPACE.sm

            self._render_bullets(
                slide, side.points,
                Rect(inner.x, y, inner.w, max(20.0, inner.bottom - y)),
                accent=accent, on=surface, size=self.scale.body_sm, compact=True,
            )

        if s.verdict:
            vr = Rect(z.body.x, z.body.bottom - verdict_h, z.body.w, verdict_h)
            accent = t.secondary
            self._shape(
                slide, vr,
                MSO_SHAPE.ROUNDED_RECTANGLE if t.radius else MSO_SHAPE.RECTANGLE,
                D.mix(t.background, accent, 0.12),
                radius=t.radius if t.radius else None,
            )
            self._shape(slide, Rect(vr.x, vr.y, 4, vr.h), MSO_SHAPE.RECTANGLE, accent)
            bgv = D.mix(t.background, accent, 0.12)
            self._text(
                slide, Rect(vr.x + D.SPACE.lg, vr.y + 9, 74, 16), "VERDICT",
                self.scale.micro, D.ensure_contrast(accent, bgv, 3.0), bold=True,
                caps=True, tracking=1.2, font=t.body_font, fit=False,
            )
            self._text(
                slide, Rect(vr.x + D.SPACE.lg, vr.y + 24, vr.w - D.SPACE.lg * 2, 24),
                s.verdict, self.scale.body, t.ink_on(bgv), bold=True,
                font=t.body_font, serif=t.body_serif,
            )

    def _agenda(self, slide, s: Slide) -> None:
        t = self.theme
        self._bg(slide, t.background)
        z = self._heading(slide, s if s.title else _retitled(s, "Agenda"))
        items = s.bullets
        if not items:
            return

        rows = z.body.split_v(len(items), D.SPACE.xs)
        for i, (b, r) in enumerate(zip(items, rows)):
            accent = t.accent(i)
            # A large outlined numeral rather than a filled chip: lighter on
            # the page, and it lets the item text be the dominant element.
            self._text(
                slide, Rect(r.x, r.cy - r.h * 0.34, 54, r.h * 0.8), f"{i + 1:02d}",
                min(self.scale.h1, r.h * 0.62), D.ensure_contrast(accent, t.background, 3.0),
                bold=True, font=self._heading_font(), serif=t.heading_serif,
                anchor=MSO_ANCHOR.MIDDLE, fit=False, line_spacing=1.0,
            )
            self._text(
                slide, Rect(r.x + 66, r.y, r.w - 66, r.h), b.text, self.scale.h3,
                t.ink_on(t.background), bold=True, anchor=MSO_ANCHOR.MIDDLE,
                font=self._heading_font(), serif=t.heading_serif,
                line_spacing=self.scale.line_height_tight,
            )
            if i < len(items) - 1:
                self._shape(slide, Rect(r.x, r.bottom, r.w, 0.75),
                            MSO_SHAPE.RECTANGLE, t.border)

    # -- data ------------------------------------------------------------

    def _chart(self, slide, s: Slide) -> None:
        t = self.theme
        self._bg(slide, t.background)
        z = self._heading(slide, s)
        cd = s.chart
        if not cd or not cd.series:
            self._text(
                slide, Rect(z.body.x, z.body.y, z.body.w, 40), "(No chart data)",
                self.scale.body, t.muted_on(t.background), font=t.body_font,
            )
            return

        plot_area = z.body
        # The takeaway gets reserved space *before* the chart is sized, so the
        # chart is never drawn on top of it.
        if cd.takeaway:
            th = D.text_height(cd.takeaway, z.body.w - 56, self.scale.body_sm,
                               serif=t.body_serif) + D.SPACE.md * 2
            plot_area = Rect(z.body.x, z.body.y, z.body.w, max(120.0, z.body.h - th - D.SPACE.md))

        chart_data = CategoryChartData()
        cats = cd.categories or [str(i + 1) for i in range(len(cd.series[0].values))]
        chart_data.categories = cats
        for series in cd.series:
            values = list(series.values)
            # Pad/trim so a malformed series cannot desynchronise the axis.
            values = (values + [0.0] * len(cats))[: len(cats)]
            chart_data.add_series(series.name, tuple(values))

        xl_type = _chart_type(cd.chart_type)
        frame = slide.shapes.add_chart(
            xl_type, _emu(plot_area.x), _emu(plot_area.y),
            _emu(plot_area.w), _emu(plot_area.h), chart_data,
        )
        chart = frame.chart
        self._style_chart(chart, cd, frame)

        if cd.takeaway:
            y = plot_area.bottom + D.SPACE.md
            accent = self._accent_for(s)
            self._shape(slide, Rect(z.body.x, y, 3, max(20.0, z.body.bottom - y)),
                        MSO_SHAPE.RECTANGLE, accent)
            self._text(
                slide, Rect(z.body.x + D.SPACE.md, y, z.body.w - D.SPACE.md,
                            max(20.0, z.body.bottom - y)),
                cd.takeaway, self.scale.body_sm, t.ink_on(t.background), bold=True,
                font=t.body_font, serif=t.body_serif,
            )

    def _style_chart(self, chart, cd: ChartData, frame) -> None:
        """Strip chart junk and apply the theme.

        Default PowerPoint charts carry a border, a grey plot area, heavy
        gridlines and an unnecessary legend. Removing them is most of what
        makes a chart look professional; the rest is using the theme's own
        colours so the chart belongs to the deck.
        """
        t = self.theme
        kind = (cd.chart_type or "").lower()
        is_pie = kind in ("pie", "doughnut", "donut")

        # The chart graphic frame itself should have no outline or fill.
        _strip_chart_frame(chart, t.background)

        chart.has_title = False
        n = len(cd.series)

        # A legend is only information when there is more than one series.
        multi = n > 1 or is_pie
        chart.has_legend = multi
        if multi:
            try:
                chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart.legend.include_in_layout = False
                chart.legend.font.size = Pt(self.scale.caption)
                chart.legend.font.color.rgb = _rgb(t.muted_on(t.background))
                chart.legend.font.name = t.body_font
            except Exception:
                pass

        if not is_pie:
            self._style_axes(chart)

        # Colours: a single series becomes a monochrome ramp across points,
        # multiple series get distinct accents.
        try:
            plot = chart.plots[0]
            if is_pie:
                colors = t.series_colors(len(cd.categories) or 1, monochrome=False)
                series = plot.series[0]
                for i, point in enumerate(series.points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = _rgb(colors[i % len(colors)])
                    point.format.line.color.rgb = _rgb(t.background)
                    point.format.line.width = Pt(1.5)
            else:
                colors = t.series_colors(n)
                for i, series in enumerate(plot.series):
                    color = colors[i % len(colors)]
                    if kind in ("line", "line_markers"):
                        series.format.line.color.rgb = _rgb(color)
                        series.format.line.width = Pt(2.5)
                        series.smooth = False
                    else:
                        series.format.fill.solid()
                        series.format.fill.fore_color.rgb = _rgb(color)
                        series.format.line.fill.background()
            # Slimmer bars with a real gap read as modern; the default is
            # almost touching.
            if kind in ("bar", "column", "stacked_bar", "stacked_column"):
                plot.gap_width = 60
                if n > 1:
                    plot.overlap = -8 if not kind.startswith("stacked") else 100
        except Exception:
            pass

        if cd.data_labels:
            try:
                plot = chart.plots[0]
                plot.has_data_labels = True
                dl = plot.data_labels
                dl.font.size = Pt(self.scale.caption)
                dl.font.bold = True
                dl.font.color.rgb = _rgb(t.ink_on(t.background))
                dl.font.name = t.body_font
                if is_pie:
                    dl.show_percentage = True
                    dl.show_value = False
                    dl.number_format = "0%"
                    dl.number_format_is_linked = False
                else:
                    dl.position = XL_LABEL_POSITION.OUTSIDE_END
            except Exception:
                pass

    def _style_axes(self, chart) -> None:
        t = self.theme
        try:
            va = chart.value_axis
            va.has_major_gridlines = True
            va.major_gridlines.format.line.color.rgb = _rgb(t.chart_grid)
            va.major_gridlines.format.line.width = Pt(0.75)
            va.format.line.fill.background()          # drop the axis spine
            va.major_tick_mark = XL_TICK_MARK.NONE
            va.minor_tick_mark = XL_TICK_MARK.NONE
            va.tick_labels.font.size = Pt(self.scale.caption)
            va.tick_labels.font.color.rgb = _rgb(t.muted_on(t.background))
            va.tick_labels.font.name = t.body_font
            va.has_minor_gridlines = False
        except Exception:
            pass
        try:
            ca = chart.category_axis
            ca.has_major_gridlines = False
            ca.has_minor_gridlines = False
            ca.format.line.color.rgb = _rgb(t.chart_grid)
            ca.format.line.width = Pt(0.75)
            ca.major_tick_mark = XL_TICK_MARK.NONE
            ca.minor_tick_mark = XL_TICK_MARK.NONE
            ca.tick_labels.font.size = Pt(self.scale.caption)
            ca.tick_labels.font.color.rgb = _rgb(t.muted_on(t.background))
            ca.tick_labels.font.name = t.body_font
        except Exception:
            pass

    def _table(self, slide, s: Slide) -> None:
        """A native table, restyled to read as data rather than as a grid.

        The default python-pptx table arrives with a blue banded style and
        borders on every cell. Modern data tables use a single rule under the
        header, subtle row banding and right-aligned numerals.
        """
        t = self.theme
        self._bg(slide, t.background)
        z = self._heading(slide, s)
        td = s.table
        if not td or not td.headers:
            return

        rows = len(td.rows) + 1
        cols = len(td.headers)
        row_h = min(42.0, max(26.0, z.body.h / max(1, rows)))
        height = min(z.body.h, row_h * rows)

        frame = slide.shapes.add_table(
            rows, cols, _emu(z.body.x), _emu(z.body.y), _emu(z.body.w), _emu(height)
        )
        table = frame.table
        _strip_table_style(table)

        try:
            table.rows[0].height = _emu(row_h)
        except Exception:
            pass

        header_bg = D.mix(t.background, t.secondary, 0.14)
        for c, header in enumerate(td.headers):
            cell = table.cell(0, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(header_bg)
            self._style_cell(
                cell, header, bold=True, color=D.ensure_contrast(t.text, header_bg, 4.5),
                size=self.scale.body_sm,
                align=PP_ALIGN.RIGHT if c in td.numeric_columns else PP_ALIGN.LEFT,
            )

        for r, row in enumerate(td.rows, start=1):
            highlight = td.highlight_row is not None and (r - 1) == td.highlight_row
            if highlight:
                band = D.mix(t.background, t.secondary, 0.10)
            else:
                band = t.background if r % 2 else t.surface
            for c in range(cols):
                cell = table.cell(r, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = _rgb(band)
                value = row[c] if c < len(row) else ""
                self._style_cell(
                    cell, value, bold=highlight,
                    color=D.ensure_contrast(t.text, band, 4.5),
                    size=self.scale.body_sm,
                    align=PP_ALIGN.RIGHT if c in td.numeric_columns else PP_ALIGN.LEFT,
                )

        # One strong rule under the header, in place of cell borders.
        self._shape(
            slide, Rect(z.body.x, z.body.y + row_h - 1.5, z.body.w, 1.5),
            MSO_SHAPE.RECTANGLE, t.secondary,
        )

    def _style_cell(self, cell, text, *, bold: bool, color: str, size: float, align=PP_ALIGN.LEFT) -> None:
        t = self.theme
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = _emu(D.SPACE.sm)
        cell.margin_top = cell.margin_bottom = _emu(D.SPACE.xxs)
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = str(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = _rgb(color)
        run.font.name = t.body_font

    def _image(self, slide, s: Slide) -> None:
        t = self.theme
        self._bg(slide, t.background)
        z = self._heading(slide, s) if s.title else None
        area = z.body if z else Rect(self.grid.left, self.grid.top,
                                     self.grid.content_w, self.grid.content_h)
        if s.caption:
            area = Rect(area.x, area.y, area.w, area.h - 28)

        placed = False
        if s.image_path and os.path.exists(s.image_path):
            try:
                # Fit inside the area preserving aspect ratio, then centre it.
                from PIL import Image  # type: ignore

                with Image.open(s.image_path) as im:
                    iw, ih = im.size
                ratio = min(area.w / iw, area.h / ih)
                w, h = iw * ratio, ih * ratio
                slide.shapes.add_picture(
                    s.image_path, _emu(area.cx - w / 2), _emu(area.cy - h / 2),
                    width=_emu(w), height=_emu(h),
                )
                placed = True
            except Exception:
                try:
                    slide.shapes.add_picture(
                        s.image_path, _emu(area.x), _emu(area.y), width=_emu(area.w)
                    )
                    placed = True
                except Exception:
                    placed = False

        if not placed:
            # The placeholder says only that an image is missing. Repeating the
            # caption here duplicates the caption drawn below the frame.
            self._card(slide, area, fill=t.surface)
            label = "Image placeholder"
            if s.image_path:
                label = f"Image not found: {os.path.basename(s.image_path)}"
            self._text(
                slide, Rect(area.x, area.cy - 14, area.w, 28), label, self.scale.body,
                t.muted_on(t.surface), align=PP_ALIGN.CENTER, font=t.body_font,
            )
        if s.caption:
            self._text(
                slide, Rect(area.x, area.bottom + D.SPACE.xs, area.w, 22), s.caption,
                self.scale.caption, t.muted_on(t.background), align=PP_ALIGN.CENTER,
                font=t.body_font, serif=t.body_serif,
            )


# --------------------------------------------------------------------------
# Chart / table XML cleanup
# --------------------------------------------------------------------------


def _strip_chart_frame(chart, bg: str) -> None:
    """Remove the chart's outline and give it the slide's background."""
    try:
        chart_space = chart._chartSpace
        spPr = chart_space.find(qn("c:spPr"))
        if spPr is None:
            spPr = chart_space.makeelement(qn("c:spPr"), {})
            # c:spPr must follow c:chart in the schema.
            chart_el = chart_space.find(qn("c:chart"))
            if chart_el is not None:
                chart_el.addnext(spPr)
            else:
                chart_space.append(spPr)
        for tag in ("a:solidFill", "a:noFill", "a:ln"):
            for existing in spPr.findall(qn(tag)):
                spPr.remove(existing)
        fill = spPr.makeelement(qn("a:solidFill"), {})
        clr = spPr.makeelement(qn("a:srgbClr"), {"val": bg.lstrip("#").upper()})
        fill.append(clr)
        spPr.append(fill)
        ln = spPr.makeelement(qn("a:ln"), {})
        ln.append(spPr.makeelement(qn("a:noFill"), {}))
        spPr.append(ln)
    except Exception:
        pass


def _strip_table_style(table) -> None:
    """Drop the banded blue default table style and all first-row emphasis.

    Without this the inherited style paints over every fill set afterwards.
    """
    try:
        tbl = table._tbl
        tblPr = tbl.find(qn("a:tblPr"))
        if tblPr is not None:
            # A nil style id means "no style", leaving explicit fills visible.
            for child in list(tblPr):
                if child.tag == qn("a:tableStyleId"):
                    tblPr.remove(child)
        table.first_row = False
        table.horz_banding = False
    except Exception:
        pass


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------


def _chart_type(name: str) -> int:
    mapping = {
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "stacked_bar": XL_CHART_TYPE.BAR_STACKED,
        "stacked_column": XL_CHART_TYPE.COLUMN_STACKED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "line_plain": XL_CHART_TYPE.LINE,
        "pie": XL_CHART_TYPE.PIE,
        "doughnut": XL_CHART_TYPE.DOUGHNUT,
        "donut": XL_CHART_TYPE.DOUGHNUT,
        "area": XL_CHART_TYPE.AREA,
        "radar": XL_CHART_TYPE.RADAR,
        "scatter": XL_CHART_TYPE.XY_SCATTER_LINES,
    }
    return mapping.get((name or "column").lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)


def _today() -> str:
    return _dt.date.today().strftime("%B %Y")


def _retitled(s: Slide, title: str) -> Slide:
    from copy import copy

    out = copy(s)
    out.title = title
    return out


def render_deck(deck: Deck, output_path: str, theme: Optional[Theme] = None) -> str:
    """Render ``deck`` to ``output_path`` (``.pptx``) and return the path.

    Creates the parent directory if it does not already exist.
    """
    parent = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(parent, exist_ok=True)
    renderer = DeckRenderer(deck, theme=theme)
    prs = renderer.render()
    prs.save(output_path)
    return output_path


__all__ = ["DeckRenderer", "render_deck"]
