"""Portfolio layouts for KimiAgent — showcasing visual work, not just content.

The stock :class:`~kimiagent.renderer.DeckRenderer` is built for *content* decks
(bullets, charts, tables). Its ``image`` layout forces artwork to the full
content width with no height cap, so a portrait piece overflows the slide. This
module adds layouts designed for **creative portfolios**, where the artwork is
the subject and the words are commentary on it:

``gallery``
    One piece per slide, presented like work hung on a gallery wall: artwork
    fitted (never cropped or stretched) on an offset accent block, with an
    art-directed critique column. Composition alternates left/right down the
    deck so the sequence has rhythm.
``grid``
    A contact-sheet index of every piece.
``contact``
    A closing slide for studio contact details.

Everything else — the deck spec, theme system, primitives and native chart,
table, stats and agenda layouts — is inherited from KimiAgent unchanged.

Usage::

    from kimiagent.portfolio import render_portfolio
    render_portfolio(deck, "portfolio.pptx")
"""

from __future__ import annotations

from typing import List, Optional, Sequence, Tuple

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from .models import Deck, Slide, SlideType
from .renderer import DeckRenderer, SLIDE_W, SLIDE_H, MARGIN, CONTENT_W, _rgb

try:  # Pillow is only needed to measure artwork; degrade gracefully without it.
    from PIL import Image
except Exception:  # pragma: no cover
    Image = None


def _aspect(path: str) -> float:
    """Width / height of an image, defaulting to 4:3 if it cannot be read."""
    if Image is None:
        return 4 / 3
    try:
        with Image.open(path) as im:
            return im.width / im.height
    except Exception:
        return 4 / 3


class PortfolioRenderer(DeckRenderer):
    """Renderer with portfolio-first layouts on top of KimiAgent's engine."""

    # ---------------------------------------------------------------- dispatch
    def render(self):
        dispatch = {
            SlideType.GALLERY: self._gallery,
            SlideType.GRID: self._grid,
            SlideType.CONTACT: self._contact,
            SlideType.COVER: self._portfolio_cover,
            SlideType.SECTION: self._portfolio_section,
        }
        total = len(self.deck.slides)
        gallery_total = sum(1 for s in self.deck.slides if s.type == SlideType.GALLERY)
        gallery_seen = 0

        for idx, slide in enumerate(self.deck.slides):
            pptx_slide = self.prs.slides.add_slide(self._blank)
            builder = dispatch.get(slide.type)
            if slide.type == SlideType.GALLERY:
                gallery_seen += 1
                self._gallery(pptx_slide, slide, gallery_seen, gallery_total)
            elif builder is not None:
                builder(pptx_slide, slide)
            else:
                # Fall back to KimiAgent's own native builders (stats, table,
                # quote, agenda, chart, bullets ...) — unchanged.
                super_dispatch = {
                    SlideType.AGENDA: self._agenda,
                    SlideType.BULLETS: self._bullets,
                    SlideType.TWO_COLUMN: self._two_column,
                    SlideType.STATS: self._stats,
                    SlideType.CHART: self._chart,
                    SlideType.TABLE: self._table,
                    SlideType.QUOTE: self._quote,
                    SlideType.IMAGE: self._image,
                    SlideType.CLOSING: self._closing,
                }
                super_dispatch.get(slide.type, self._bullets)(pptx_slide, slide)
            self._add_notes(pptx_slide, slide.notes)
            if slide.type in (SlideType.AGENDA, SlideType.BULLETS, SlideType.TWO_COLUMN,
                              SlideType.STATS, SlideType.CHART, SlideType.TABLE,
                              SlideType.GRID):
                self._add_footer(pptx_slide, idx + 1, total)
        return self.prs

    # ---------------------------------------------------------------- helpers
    def _spaced(self, run, pts: float) -> None:
        """Letter-spacing — python-pptx exposes no API for this."""
        run.font._rPr.set("spc", str(int(pts * 100)))

    def _label(self, slide, x, y, w, text, color, size=8.5, spc=2.0, align=PP_ALIGN.LEFT):
        """Small uppercase letterspaced label — the deck's connective tissue."""
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.3))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text.upper()
        run.font.size = Pt(size)
        run.font.bold = True
        run.font.name = self.theme.heading_font
        run.font.color.rgb = _rgb(color)
        self._spaced(run, spc)
        return box

    def _para(self, slide, x, y, w, h, text, size, color, bold=False, font=None,
              line=1.35, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
        """Multi-line text with real leading control (supports '\\n')."""
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = anchor
        for i, chunk in enumerate(str(text).split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = line
            run = p.add_run()
            run.text = chunk
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.name = font or self.theme.body_font
            run.font.color.rgb = _rgb(color)
        return box

    def _block(self, x, y, w, h, slide, color):
        return self._rect(slide, Inches(x), Inches(y), Inches(w), Inches(h), color)

    def _fit(self, slide, path: str, bx: float, by: float, bw: float, bh: float,
             accent: str, offset: float = 0.17) -> Tuple[float, float, float, float]:
        """Place artwork centred in a box, aspect preserved, on an offset block.

        Returns the placed rectangle as ``(x, y, w, h)`` in inches.
        """
        ar = _aspect(path)
        w, h = bw, bw / ar
        if h > bh:
            h, w = bh, bh * ar
        x, y = bx + (bw - w) / 2, by + (bh - h) / 2
        # offset accent block — the "designer's pop" behind the work
        self._block(x + offset, y + offset, w, h, slide, accent)
        slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
        # hairline keeps light artwork from bleeding into the dark wall
        edge = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                      Inches(w), Inches(h))
        edge.fill.background()
        edge.line.color.rgb = _rgb("4A3C5C")
        edge.line.width = Pt(0.75)
        edge.shadow.inherit = False
        edge.text_frame.text = ""
        return x, y, w, h

    def _swatch_strip(self, slide, x: float, y: float, palette: Sequence[str],
                      bar_w: float = 0.15, bar_h: float = 0.46, gap: float = 0.055):
        """Vertical bars sampled from the artwork — a designer's colour read."""
        for i, hexc in enumerate(palette):
            self._block(x + i * (bar_w + gap), y, bar_w, bar_h, slide, hexc.lstrip("#"))

    def _progress_rail(self, slide, fraction: float):
        """Thin rail showing how far through the portfolio this piece sits."""
        y, h = 7.36, 0.055
        self._block(0, y, SLIDE_W / 914400, h, slide, "2A2036")
        self._block(0, y, max(0.05, (SLIDE_W / 914400) * fraction), h, slide,
                    self.theme.secondary)

    # -- override: headings must read light on the dark gallery wall ----------
    def _heading(self, slide, title: str, subtitle: str = ""):
        self._para(slide, MARGIN / 914400, 0.55, CONTENT_W / 914400, 0.9, title, 30,
                   self.theme.text, bold=True, font=self.theme.heading_font)
        self._rect(slide, MARGIN, Inches(1.5), Inches(1.1), Inches(0.09),
                   self.theme.secondary)
        if subtitle:
            self._para(slide, MARGIN / 914400, 1.72, CONTENT_W / 914400, 0.5, subtitle,
                       13, self.theme.text_muted)

    # -- override: agenda must never silently drop items ---------------------
    def _agenda(self, slide, s: Slide):
        """Agenda that resizes to fit every item.

        KimiAgent's stock ``_agenda`` uses fixed 0.85in rows and ``break``s out
        of the loop once they run past the bottom margin — so an agenda with
        more than four entries loses the rest with no warning. Here the row
        height, gap and type size are derived from the item count, so the list
        always fits and nothing is dropped.
        """
        self._body_background(slide)
        self._heading(slide, s.title or "Agenda", s.subtitle)
        items = list(s.bullets)
        if not items:
            return

        top = 2.3 if s.subtitle else 2.1
        bottom = 7.5 - 0.8
        avail = bottom - top
        n = len(items)

        gap = 0.2 if n <= 5 else 0.14
        row_h = min(0.85, (avail - gap * (n - 1)) / n)
        block = row_h * n + gap * (n - 1)
        y0 = top + max(0.0, (avail - block) / 2)   # centre the block vertically

        chip = min(row_h, 0.7)
        num_size = 20 if chip >= 0.62 else 16
        label_size = 19 if row_h >= 0.7 else (17 if row_h >= 0.56 else 15)

        for i, b in enumerate(items):
            y = y0 + i * (row_h + gap)
            # number chip, vertically centred on the row
            cy = y + (row_h - chip) / 2
            block_shape = self._block(0.7, cy, chip, chip, slide, self.theme.accent(i))
            tf = block_shape.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(i + 1)
            run.font.size = Pt(num_size)
            run.font.bold = True
            run.font.name = self.theme.heading_font
            run.font.color.rgb = _rgb(self.theme.text_on_primary)
            # label
            self._para(slide, 0.7 + chip + 0.32, y, 11.4 - chip, row_h, b.text,
                       label_size, self.theme.text, bold=True,
                       font=self.theme.heading_font, anchor=MSO_ANCHOR.MIDDLE)

    # -- override: quote text must not be dark-on-dark -----------------------
    def _quote(self, slide, s: Slide):
        self._fill_background(slide, self.theme.surface)
        self._para(slide, 0.72, 0.85, 3.0, 2.0, "\u201C", 150, self.theme.secondary,
                   bold=True, font=self.theme.heading_font)
        self._para(slide, 1.5, 2.45, 10.3, 2.5, s.quote, 27, self.theme.text,
                   italic=True, bold=True, font=self.theme.heading_font, line=1.22,
                   anchor=MSO_ANCHOR.TOP)
        if s.attribution:
            self._para(slide, 1.5, 5.25, 10.3, 0.6, f"\u2014 {s.attribution}", 15,
                       self.theme.text_muted)

    # ---------------------------------------------------------------- layouts
    def _portfolio_cover(self, slide, s: Slide):
        """Editorial cover: brand rail, oversized title, meta row."""
        self._fill_background(slide, self.theme.background)
        # left brand rail split into the two brand accents
        self._block(0, 0, 0.32, 3.75, slide, self.theme.secondary)
        self._block(0, 3.75, 0.32, 3.75, slide, self.theme.accent(1))
        # faint oversized numeral, top-right — clear of the title block below
        self._para(slide, 9.2, 0.82, 3.5, 2.5, s.meta.get("watermark", ""), 120,
                   "2A1E38", bold=True, font=self.theme.heading_font,
                   align=PP_ALIGN.RIGHT)
        logo = s.meta.get("logo")
        if logo:
            try:
                slide.shapes.add_picture(logo, Inches(1.05), Inches(0.85), Inches(2.3))
            except Exception:
                pass
        self._label(slide, 1.08, 2.28, 8.0, s.meta.get("eyebrow", ""),
                    self.theme.accent(3), size=10, spc=3.2)
        self._para(slide, 1.02, 2.78, 10.8, 2.4, s.title, 60, self.theme.text,
                   bold=True, font=self.theme.heading_font, line=1.02)
        self._block(1.09, 5.26, 1.5, 0.055, slide, self.theme.secondary)
        if s.subtitle:
            self._para(slide, 1.05, 5.6, 9.1, 1.0, s.subtitle, 12.5,
                       self.theme.text_muted, line=1.5)
        meta = s.meta.get("footer", "")
        if meta:
            self._label(slide, 1.05, 6.82, 11.3, meta, self.theme.accent(2), size=9.5,
                        spc=1.8)

    def _portfolio_section(self, slide, s: Slide):
        """Section divider: tinted numeral, bottom-anchored title, count."""
        self._fill_background(slide, self.theme.background)
        self._block(0, 0, 0.32, 7.5, slide, self.theme.accent(0))
        num = s.meta.get("number", "")
        self._para(slide, 8.4, 1.0, 4.3, 3.4, num, 150, "251B33", bold=True,
                   font=self.theme.heading_font, align=PP_ALIGN.RIGHT)
        self._label(slide, 1.05, 2.0, 5.0, s.meta.get("eyebrow", "Section"),
                    self.theme.accent(3), size=10, spc=3.0)
        # bottom-anchored: a two-line title grows upward, never into the rule
        self._para(slide, 1.0, 2.44, 7.2, 1.95, s.title, 42, self.theme.text,
                   bold=True, font=self.theme.heading_font, line=1.06,
                   anchor=MSO_ANCHOR.BOTTOM)
        self._block(1.07, 4.6, 1.25, 0.05, slide, self.theme.secondary)
        if s.subtitle:
            self._para(slide, 1.05, 4.95, 7.0, 0.9, s.subtitle, 12.5,
                       self.theme.text_muted, line=1.5)
        if s.caption:
            self._label(slide, 1.05, 5.95, 6.0, s.caption, self.theme.accent(2),
                        size=9.5, spc=1.6)

    def _gallery(self, slide, s: Slide, n: int = 1, total: int = 1):
        """One portfolio piece: artwork on an offset block + critique column.

        Composition mirrors on alternate slides, so the deck reads with rhythm
        instead of as a repeating template.
        """
        self._fill_background(slide, self.theme.background)
        art_right = (n % 2 == 1)          # odds: artwork right, evens: artwork left
        accent = self.theme.accent(n - 1)

        if art_right:
            tx, ax = 0.88, 6.05
        else:
            tx, ax = 7.95, 0.62
        tw, aw = 4.45, 6.7

        # ---- artwork stage
        self._fit(slide, s.image_path, ax, 0.92, aw, 5.55, accent)

        # ---- critique column
        # index: big number + total
        self._para(slide, tx, 0.72, 2.2, 1.0, f"{n:02d}", 52, accent, bold=True,
                   font=self.theme.heading_font)
        self._para(slide, tx + 1.08, 1.12, 1.4, 0.4, f"/ {total:02d}", 12,
                   self.theme.text_muted, bold=True, font=self.theme.heading_font)
        # category
        self._label(slide, tx + 0.03, 1.86, tw, s.meta.get("category", ""), accent,
                    size=8.5, spc=2.2)
        self._block(tx + 0.03, 2.18, 0.8, 0.04, slide, accent)
        # title + client
        self._para(slide, tx, 2.42, tw, 1.15, s.title, 22, self.theme.text, bold=True,
                   font=self.theme.heading_font, line=1.15)
        client = s.meta.get("client", "")
        if client:
            self._label(slide, tx + 0.03, 3.62, tw, client, self.theme.text_muted,
                        size=9, spc=1.5)
        # the written description
        self._para(slide, tx, 3.98, tw, 2.15, s.caption, 10.5, self.theme.text,
                   line=1.62)
        # format line
        fmt = s.meta.get("format", "")
        if fmt:
            self._label(slide, tx + 0.03, 6.2, tw, "Format", self.theme.text_muted,
                        size=7.5, spc=1.8)
            self._para(slide, tx, 6.45, tw, 0.4, fmt, 9, self.theme.text_muted)
        # sampled palette
        palette = s.meta.get("palette") or []
        if palette:
            self._swatch_strip(slide, tx + 0.03, 6.8, palette)

        self._progress_rail(slide, n / max(total, 1))

    def _grid(self, slide, s: Slide):
        """Contact-sheet index of every piece in the portfolio.

        The header is kept compact so the tiles get the space, and a partial
        final row is centred rather than left-aligned — otherwise an item count
        that doesn't divide evenly by the column count leaves a lopsided hole.
        """
        self._fill_background(slide, self.theme.background)
        self._para(slide, 0.62, 0.36, 9.0, 0.6, s.title, 26, self.theme.text,
                   bold=True, font=self.theme.heading_font)
        self._block(0.66, 1.0, 0.95, 0.05, slide, self.theme.secondary)
        if s.subtitle:
            self._para(slide, 0.62, 1.14, 11.0, 0.35, s.subtitle, 10.5,
                       self.theme.text_muted)

        paths = list(s.images)
        if not paths:
            return
        cols = max(1, int(s.meta.get("cols", 7)))
        rows = (len(paths) + cols - 1) // cols
        x0, y0 = 0.42, 1.62
        right, bottom = 0.42, 0.62      # bottom clears the page footer
        gx, gy = 0.10, 0.12
        avail_w = 13.333 - x0 - right
        avail_h = 7.5 - y0 - bottom
        cw = (avail_w - (cols - 1) * gx) / cols
        ch = (avail_h - (rows - 1) * gy) / rows
        pad = 0.085

        for i, p in enumerate(paths):
            r, c = divmod(i, cols)
            in_row = min(cols, len(paths) - r * cols)
            row_w = in_row * cw + (in_row - 1) * gx
            rx = x0 + (avail_w - row_w) / 2      # centre a short final row
            cx = rx + c * (cw + gx)
            cy = y0 + r * (ch + gy)
            self._block(cx, cy, cw, ch, slide, self.theme.surface)
            ar = _aspect(p)
            iw, ih = cw - 2 * pad, (cw - 2 * pad) / ar
            if ih > ch - 2 * pad:
                ih, iw = ch - 2 * pad, (ch - 2 * pad) * ar
            slide.shapes.add_picture(p, Inches(cx + (cw - iw) / 2),
                                     Inches(cy + (ch - ih) / 2), Inches(iw), Inches(ih))

    def _contact(self, slide, s: Slide):
        """Closing slide: studio contact details in a clean meta row."""
        self._fill_background(slide, self.theme.background)
        self._block(0, 0, 0.32, 3.75, slide, self.theme.secondary)
        self._block(0, 3.75, 0.32, 3.75, slide, self.theme.accent(1))
        logo = s.meta.get("logo")
        if logo:
            try:
                slide.shapes.add_picture(logo, Inches(1.05), Inches(0.95), Inches(2.3))
            except Exception:
                pass
        self._para(slide, 1.02, 2.5, 10.6, 1.9, s.title, 52, self.theme.text,
                   bold=True, font=self.theme.heading_font, line=1.05)
        self._block(1.09, 4.5, 1.5, 0.055, slide, self.theme.secondary)
        items: List[Tuple[str, str]] = list(s.meta.get("contact", []))
        x = 1.05
        for label, value in items:
            self._label(slide, x, 4.88, 3.5, label, self.theme.accent(3), size=8.5,
                        spc=2.0)
            self._para(slide, x, 5.2, 3.7, 0.5, value, 15.5, self.theme.text,
                       bold=True, font=self.theme.heading_font)
            x += 3.75
        if s.subtitle:
            self._para(slide, 1.05, 6.15, 10.9, 0.6, s.subtitle, 11,
                       self.theme.text_muted, line=1.5)
        if s.caption:
            self._label(slide, 1.05, 6.85, 6.0, s.caption, self.theme.accent(2),
                        size=9.5, spc=1.8)


def render_portfolio(deck: Deck, output_path: str) -> str:
    """Render a portfolio ``deck`` to ``output_path`` and return the path."""
    import os

    parent = os.path.dirname(os.path.abspath(output_path))
    if parent:
        os.makedirs(parent, exist_ok=True)
    prs = PortfolioRenderer(deck).render()
    prs.save(output_path)
    return output_path
