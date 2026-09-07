"""Render a :class:`~kimiagent.models.Deck` into a native ``.pptx``.

This is where the "native depth" happens: instead of dumping text into a
template, each slide is composed from real PowerPoint primitives — shapes,
native (editable) charts and native tables — laid out on a blank 16:9 canvas
with a consistent, theme-driven visual system.

Only ``python-pptx`` is required.
"""

from __future__ import annotations

from typing import List, Optional

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

import struct

from .models import Deck, Slide, SlideType, Bullet, Stat, ChartData, TableData, ImageItem
from .themes import Theme, get_theme


# 16:9 canvas dimensions.
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.7)
CONTENT_W = SLIDE_W - 2 * MARGIN


def _rgb(hex_str: str) -> RGBColor:
    return RGBColor.from_string(hex_str)


def _luminance(hex_str: str) -> float:
    """Relative luminance (0=black, 1=white) of an RRGGBB colour."""
    r = int(hex_str[0:2], 16) / 255.0
    g = int(hex_str[2:4], 16) / 255.0
    b = int(hex_str[4:6], 16) / 255.0
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contrast(a: str, b: str) -> float:
    """WCAG-style contrast ratio between two colours (1 = none, 21 = max)."""
    la, lb = _luminance(a), _luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


class DeckRenderer:
    """Renders a deck. One instance per deck keeps theme state handy."""

    def __init__(self, deck: Deck) -> None:
        self.deck = deck
        self.theme: Theme = get_theme(deck.theme)
        self.prs = Presentation()
        self.prs.slide_width = SLIDE_W
        self.prs.slide_height = SLIDE_H
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
            SlideType.CHART: self._chart,
            SlideType.TABLE: self._table,
            SlideType.QUOTE: self._quote,
            SlideType.IMAGE: self._image,
            SlideType.GALLERY: self._gallery,
            SlideType.CLOSING: self._closing,
        }
        total = len(self.deck.slides)
        for idx, slide in enumerate(self.deck.slides):
            builder = dispatch.get(slide.type, self._bullets)
            pptx_slide = self.prs.slides.add_slide(self._blank)
            builder(pptx_slide, slide)
            self._add_notes(pptx_slide, slide.notes)
            # Footer + page number on everything except full-bleed opener/closer.
            if slide.type not in (SlideType.COVER, SlideType.SECTION, SlideType.CLOSING):
                self._add_footer(pptx_slide, idx + 1, total)
        return self.prs

    # -- shared primitives ------------------------------------------------

    def _fill_background(self, slide, hex_color: str) -> None:
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = _rgb(hex_color)

    def _rect(self, slide, left, top, width, height, hex_color, line_color: Optional[str] = None):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(hex_color)
        if line_color:
            shape.line.color.rgb = _rgb(line_color)
            shape.line.width = Pt(0.75)
        else:
            shape.line.fill.background()
        shape.shadow.inherit = False
        return shape

    def _text(
        self,
        slide,
        left,
        top,
        width,
        height,
        text,
        size,
        color,
        bold=False,
        italic=False,
        align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP,
        font=None,
        line_spacing=None,
    ):
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        p = tf.paragraphs[0]
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = _rgb(color)
        run.font.name = font or self.theme.body_font
        return box

    def _title_color(self, bg_hex: str) -> str:
        """A legible title colour on ``bg_hex``.

        Uses the theme's rich primary when it reads clearly against the
        background (light themes), otherwise falls back to the light body
        colour so titles stay legible on dark themes.
        """
        if _contrast(self.theme.primary, bg_hex) >= 2.8:
            return self.theme.primary
        return self.theme.text

    def _accent_bar(self, slide, top=None):
        """A short accent bar used as a heading underline motif."""
        top = top if top is not None else Inches(1.55)
        return self._rect(slide, MARGIN, top, Inches(1.1), Inches(0.09), self.theme.secondary)

    def _heading(self, slide, title: str, subtitle: str = ""):
        self._text(
            slide, MARGIN, Inches(0.55), CONTENT_W, Inches(0.9),
            title, 30, self._title_color(self.theme.background), bold=True, font=self.theme.heading_font,
        )
        self._accent_bar(slide)
        if subtitle:
            self._text(
                slide, MARGIN, Inches(1.02), CONTENT_W, Inches(0.5),
                subtitle, 15, self.theme.text_muted, italic=True,
            )

    def _add_footer(self, slide, page: int, total: int):
        self._text(
            slide, MARGIN, SLIDE_H - Inches(0.5), Inches(8), Inches(0.35),
            self.deck.title, 9, self.theme.text_muted,
        )
        self._text(
            slide, SLIDE_W - MARGIN - Inches(2), SLIDE_H - Inches(0.5), Inches(2), Inches(0.35),
            f"{page} / {total}", 9, self.theme.text_muted, align=PP_ALIGN.RIGHT,
        )

    def _add_notes(self, slide, notes: str):
        if not notes:
            return
        slide.notes_slide.notes_text_frame.text = notes

    def _body_background(self, slide):
        self._fill_background(slide, self.theme.background)

    # -- slide builders ---------------------------------------------------

    def _cover(self, slide, s: Slide):
        self._fill_background(slide, self.theme.primary)
        # Accent side band
        self._rect(slide, 0, 0, Inches(0.35), SLIDE_H, self.theme.secondary)
        # Decorative corner block
        self._rect(slide, SLIDE_W - Inches(3.2), SLIDE_H - Inches(0.55),
                   Inches(3.2), Inches(0.55), self.theme.secondary)

        self._text(
            slide, Inches(1.0), Inches(2.4), Inches(11.0), Inches(2.0),
            s.title or self.deck.title, 46, self.theme.text_on_primary,
            bold=True, font=self.theme.heading_font, line_spacing=1.05,
        )
        if s.subtitle:
            self._text(
                slide, Inches(1.05), Inches(4.5), Inches(10.5), Inches(0.8),
                s.subtitle, 20, self.theme.text_on_primary, italic=True,
            )
        if self.deck.author:
            self._text(
                slide, Inches(1.05), Inches(5.2), Inches(10.5), Inches(0.6),
                self.deck.author, 14, self.theme.text_on_primary,
            )

    def _closing(self, slide, s: Slide):
        self._fill_background(slide, self.theme.primary)
        self._rect(slide, 0, SLIDE_H / 2 - Inches(0.04), SLIDE_W, Inches(0.08), self.theme.secondary)
        self._text(
            slide, Inches(1.0), Inches(2.6), Inches(11.3), Inches(1.4),
            s.title or "Thank you", 44, self.theme.text_on_primary,
            bold=True, align=PP_ALIGN.CENTER, font=self.theme.heading_font,
        )
        if s.subtitle:
            self._text(
                slide, Inches(1.0), Inches(4.1), Inches(11.3), Inches(0.8),
                s.subtitle, 20, self.theme.text_on_primary, italic=True, align=PP_ALIGN.CENTER,
            )

    def _section(self, slide, s: Slide):
        self._fill_background(slide, self.theme.surface)
        self._rect(slide, 0, Inches(2.9), Inches(0.5), Inches(1.6), self.theme.secondary)
        if s.subtitle:
            self._text(
                slide, Inches(0.9), Inches(2.7), Inches(11), Inches(0.6),
                s.subtitle.upper(), 15, self.theme.secondary, bold=True,
            )
        self._text(
            slide, Inches(0.85), Inches(3.15), Inches(11.6), Inches(1.6),
            s.title, 40, self._title_color(self.theme.surface), bold=True, font=self.theme.heading_font,
        )

    def _bullets(self, slide, s: Slide):
        self._body_background(slide)
        self._heading(slide, s.title, s.subtitle)
        self._render_bullets(slide, s.bullets, MARGIN, Inches(1.9), CONTENT_W, SLIDE_H - Inches(2.6))

    def _render_bullets(self, slide, bullets: List[Bullet], left, top, width, height):
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.word_wrap = True
        first = True
        for b in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.level = min(max(b.level, 0), 4)
            p.space_after = Pt(10)
            p.line_spacing = 1.1
            # Bullet glyph via leading marker (kept simple + robust across viewers)
            marker = "▪  " if b.level == 0 else "–  "
            if b.bold_lead:
                lead = p.add_run()
                lead.text = f"{marker}{b.bold_lead}: "
                lead.font.bold = True
                lead.font.size = Pt(18 if b.level == 0 else 15)
                lead.font.color.rgb = _rgb(self.theme.primary)
                lead.font.name = self.theme.body_font
                rest = p.add_run()
                rest.text = b.text
                rest.font.size = Pt(18 if b.level == 0 else 15)
                rest.font.color.rgb = _rgb(self.theme.text)
                rest.font.name = self.theme.body_font
            else:
                run = p.add_run()
                run.text = f"{marker}{b.text}"
                run.font.size = Pt(18 if b.level == 0 else 15)
                run.font.color.rgb = _rgb(self.theme.text if b.level else self.theme.text)
                run.font.bold = b.level == 0
                run.font.name = self.theme.body_font
        return box

    def _agenda(self, slide, s: Slide):
        self._body_background(slide)
        self._heading(slide, s.title or "Agenda")
        top = Inches(2.0)
        row_h = Inches(0.85)
        gap = Inches(0.2)
        for i, b in enumerate(s.bullets):
            y = top + i * (row_h + gap)
            if y + row_h > SLIDE_H - Inches(0.7):
                break
            # Number chip
            chip = self._rect(slide, MARGIN, y, Inches(0.7), Inches(0.7), self.theme.accent(i))
            chip_tf = chip.text_frame
            chip_tf.word_wrap = True
            cp = chip_tf.paragraphs[0]
            cp.alignment = PP_ALIGN.CENTER
            crun = cp.add_run()
            crun.text = str(i + 1)
            crun.font.size = Pt(20)
            crun.font.bold = True
            crun.font.color.rgb = _rgb(self.theme.text_on_primary)
            # Label
            self._text(
                slide, MARGIN + Inches(1.0), y, CONTENT_W - Inches(1.0), Inches(0.7),
                b.text, 19, self.theme.text, bold=True, anchor=MSO_ANCHOR.MIDDLE,
            )

    def _two_column(self, slide, s: Slide):
        self._body_background(slide)
        self._heading(slide, s.title, s.subtitle)
        cols = s.columns[:2] if s.columns else [[], []]
        col_w = (CONTENT_W - Inches(0.5)) / 2
        top = Inches(1.95)
        card_h = SLIDE_H - Inches(2.6)
        for ci, col in enumerate(cols):
            left = MARGIN + ci * (col_w + Inches(0.5))
            self._rect(slide, left, top, col_w, card_h, self.theme.surface)
            # Colored header strip
            self._rect(slide, left, top, col_w, Inches(0.12), self.theme.accent(ci))
            self._render_bullets(
                slide, col,
                left + Inches(0.35), top + Inches(0.4),
                col_w - Inches(0.7), card_h - Inches(0.7),
            )

    def _stats(self, slide, s: Slide):
        self._body_background(slide)
        self._heading(slide, s.title, s.subtitle)
        stats = s.stats[:4]
        if not stats:
            return
        n = len(stats)
        gap = Inches(0.4)
        total_gap = gap * (n - 1)
        card_w = (CONTENT_W - total_gap) / n
        top = Inches(2.6)
        card_h = Inches(2.9)
        for i, st in enumerate(stats):
            left = MARGIN + i * (card_w + gap)
            self._rect(slide, left, top, card_w, card_h, self.theme.surface)
            self._rect(slide, left, top, card_w, Inches(0.14), self.theme.accent(i))
            self._text(
                slide, left, top + Inches(0.5), card_w, Inches(1.2),
                st.value, 46, self.theme.accent(i), bold=True,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=self.theme.heading_font,
            )
            self._text(
                slide, left + Inches(0.15), top + Inches(1.75), card_w - Inches(0.3), Inches(0.5),
                st.label, 15, self.theme.text, bold=True, align=PP_ALIGN.CENTER,
            )
            if st.detail:
                self._text(
                    slide, left + Inches(0.15), top + Inches(2.25), card_w - Inches(0.3), Inches(0.55),
                    st.detail, 11, self.theme.text_muted, align=PP_ALIGN.CENTER,
                )

    def _chart(self, slide, s: Slide):
        self._body_background(slide)
        self._heading(slide, s.title, s.subtitle)
        cd = s.chart
        if not cd or not cd.series:
            self._text(slide, MARGIN, Inches(2.2), CONTENT_W, Inches(1), "(No chart data)", 16, self.theme.text_muted)
            return

        chart_data = CategoryChartData()
        chart_data.categories = cd.categories or [str(i + 1) for i in range(len(cd.series[0].values))]
        for series in cd.series:
            chart_data.add_series(series.name, tuple(series.values))

        xl_type = _chart_type(cd.chart_type)
        x, y = MARGIN, Inches(2.0)
        cx, cy = CONTENT_W, SLIDE_H - Inches(2.8)
        gframe = slide.shapes.add_chart(xl_type, x, y, cx, cy, chart_data)
        chart = gframe.chart

        chart.has_title = False
        # Legend only when it adds information (multiple series or pie).
        multi = len(cd.series) > 1 or cd.chart_type == "pie"
        chart.has_legend = multi
        if multi:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(11)
            chart.legend.font.color.rgb = _rgb(self.theme.text)

        # Theme the series colours.
        try:
            self._color_chart(chart, cd)
        except Exception:
            pass  # colour theming is best-effort; never fail the render

    def _color_chart(self, chart, cd: ChartData):
        plot = chart.plots[0]
        if cd.chart_type == "pie":
            # Pie: colour each point.
            series = plot.series[0]
            for i, point in enumerate(series.points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = _rgb(self.theme.accent(i))
        else:
            for i, series in enumerate(plot.series):
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = _rgb(self.theme.accent(i))
                series.format.line.color.rgb = _rgb(self.theme.accent(i))

    def _table(self, slide, s: Slide):
        self._body_background(slide)
        self._heading(slide, s.title, s.subtitle)
        td = s.table
        if not td or not td.headers:
            return
        rows = len(td.rows) + 1
        cols = len(td.headers)
        top = Inches(2.1)
        height = min(SLIDE_H - Inches(2.9), Inches(0.55) * rows)
        gframe = slide.shapes.add_table(rows, cols, MARGIN, top, CONTENT_W, height)
        table = gframe.table

        # Header row
        for c, header in enumerate(td.headers):
            cell = table.cell(0, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(self.theme.primary)
            self._style_cell(cell, header, bold=True, color=self.theme.text_on_primary, size=14)
        # Body rows with zebra striping
        for r, row in enumerate(td.rows, start=1):
            band = self.theme.surface if r % 2 == 1 else self.theme.background
            for c in range(cols):
                cell = table.cell(r, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = _rgb(band)
                value = row[c] if c < len(row) else ""
                self._style_cell(cell, value, bold=False, color=self.theme.text, size=13)

    def _style_cell(self, cell, text, bold, color, size):
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.15)
        cell.margin_right = Inches(0.15)
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = str(text)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = _rgb(color)
        run.font.name = self.theme.body_font

    def _quote(self, slide, s: Slide):
        self._fill_background(slide, self.theme.surface)
        # Big decorative quotation mark
        self._text(
            slide, MARGIN, Inches(1.1), Inches(3), Inches(2),
            "\u201C", 160, self.theme.secondary, bold=True, font=self.theme.heading_font,
        )
        self._text(
            slide, Inches(1.4), Inches(2.5), Inches(10.5), Inches(2.6),
            s.quote, 28, self._title_color(self.theme.surface), italic=True, bold=True,
            anchor=MSO_ANCHOR.MIDDLE, font=self.theme.heading_font, line_spacing=1.15,
        )
        if s.attribution:
            self._text(
                slide, Inches(1.4), Inches(5.3), Inches(10.5), Inches(0.6),
                f"— {s.attribution}", 16, self.theme.text_muted,
            )

    def _image(self, slide, s: Slide):
        self._body_background(slide)
        if s.title:
            self._heading(slide, s.title, s.subtitle)
        top = Inches(1.9) if s.title else Inches(0.55)
        bottom_reserve = Inches(0.95) if s.caption else Inches(0.55)
        box_top = top
        box_h = SLIDE_H - box_top - bottom_reserve
        box_left, box_w = MARGIN, CONTENT_W

        placed = self._place_image_contained(slide, s.image_path, box_left, box_top, box_w, box_h)
        if not placed:
            self._rect(slide, box_left, box_top, box_w, box_h, self.theme.surface)
            self._text(
                slide, box_left, box_top + box_h / 2 - Inches(0.4), box_w, Inches(0.8),
                s.caption or "[ image ]", 16, self.theme.text_muted,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            )
        if s.caption:
            self._text(
                slide, MARGIN, SLIDE_H - Inches(0.9), CONTENT_W, Inches(0.4),
                s.caption, 13, self.theme.text_muted, italic=True, align=PP_ALIGN.CENTER,
            )

    def _gallery(self, slide, s: Slide):
        """A responsive grid of images, each contain-fit with a caption."""
        self._body_background(slide)
        self._heading(slide, s.title, s.subtitle)
        items = s.images
        if not items:
            return

        n = len(items)
        cols = _grid_columns(n)
        rows = (n + cols - 1) // cols

        area_left = MARGIN
        area_top = Inches(1.85) if (s.title or s.subtitle) else Inches(0.6)
        area_w = CONTENT_W
        area_h = SLIDE_H - area_top - Inches(0.65)

        gutter = Inches(0.3)
        cell_w = (area_w - gutter * (cols - 1)) / cols
        cell_h = (area_h - gutter * (rows - 1)) / rows
        caption_h = Inches(0.42)

        for i, item in enumerate(items):
            r, c = divmod(i, cols)
            cx = area_left + c * (cell_w + gutter)
            cy = area_top + r * (cell_h + gutter)
            # Soft card behind each image.
            self._rect(slide, cx, cy, cell_w, cell_h, self.theme.surface)
            img_box_h = cell_h - (caption_h if item.caption else Inches(0.0))
            placed = self._place_image_contained(
                slide, item.path,
                cx + Inches(0.08), cy + Inches(0.08),
                cell_w - Inches(0.16), img_box_h - Inches(0.16),
            )
            if not placed:
                self._text(
                    slide, cx, cy + img_box_h / 2 - Inches(0.3), cell_w, Inches(0.6),
                    "[ image ]", 12, self.theme.text_muted,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                )
            if item.caption:
                self._text(
                    slide, cx + Inches(0.1), cy + img_box_h - Inches(0.02), cell_w - Inches(0.2), caption_h,
                    item.caption, 10.5, self.theme.text, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE,
                )

    def _place_image_contained(self, slide, path, box_left, box_top, box_w, box_h):
        """Place an image inside a box, preserving aspect ratio and centering.

        Returns True if the image was placed, False otherwise (missing/unreadable).
        """
        if not path:
            return False
        size = _image_size(path)
        if size is None:
            # Try to place anyway with width only; if it fails, report failure.
            try:
                slide.shapes.add_picture(path, box_left, box_top, width=box_w)
                return True
            except Exception:
                return False
        iw, ih = size
        if iw <= 0 or ih <= 0:
            return False
        scale = min(box_w / iw, box_h / ih)
        w = int(iw * scale)
        h = int(ih * scale)
        left = int(box_left + (box_w - w) / 2)
        top = int(box_top + (box_h - h) / 2)
        try:
            slide.shapes.add_picture(path, left, top, width=w, height=h)
            return True
        except Exception:
            return False


def _grid_columns(n: int) -> int:
    """Pick a pleasant column count for ``n`` images."""
    return {1: 1, 2: 2, 3: 3, 4: 2, 5: 3, 6: 3}.get(n, 3 if n <= 9 else 4)


def _image_size(path):
    """Return (width, height) in pixels, or None if it can't be determined.

    Uses Pillow when available, otherwise falls back to reading PNG/JPEG headers
    directly so galleries render correctly with zero extra dependencies.
    """
    try:
        from PIL import Image  # type: ignore

        with Image.open(path) as im:
            return im.size
    except Exception:
        pass
    try:
        with open(path, "rb") as f:
            head = f.read(32)
        # PNG: 8-byte signature, then IHDR with width/height as big-endian uint32.
        if head[:8] == b"\x89PNG\r\n\x1a\n":
            w, h = struct.unpack(">II", head[16:24])
            return int(w), int(h)
        # JPEG: scan for a start-of-frame marker to read dimensions.
        if head[:2] == b"\xff\xd8":
            with open(path, "rb") as f:
                data = f.read()
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                    h = struct.unpack(">H", data[i + 5:i + 7])[0]
                    w = struct.unpack(">H", data[i + 7:i + 9])[0]
                    return int(w), int(h)
                seg_len = struct.unpack(">H", data[i + 2:i + 4])[0]
                i += 2 + seg_len
    except Exception:
        pass
    return None


def _chart_type(name: str) -> int:
    mapping = {
        "bar": XL_CHART_TYPE.BAR_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
        "area": XL_CHART_TYPE.AREA,
    }
    return mapping.get((name or "column").lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)


def render_deck(deck: Deck, output_path: str) -> str:
    """Render ``deck`` to ``output_path`` and return the path.

    Creates the parent directory if it does not already exist.
    """
    import os

    parent = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(parent, exist_ok=True)
    renderer = DeckRenderer(deck)
    prs = renderer.render()
    prs.save(output_path)
    return output_path
