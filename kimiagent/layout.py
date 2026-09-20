"""Layout and flow logic shared by the PPTX and PDF renderers.

Two concerns live here, both deliberately backend-agnostic:

**Zones.** Where the heading ends and the body begins, expressed as plain
rectangles on a :class:`~kimiagent.design.Grid`. Both renderers ask this module
for the same rectangles, so a card row sits at the same y in the deck and in
the PDF.

**Flow (pagination).** What to do when content exceeds the slide. The original
renderer silently dropped it — the agenda stopped after the eighth item, the
stats slide kept four of six, a long bullet ran off the canvas. That is the
worst available option, because the author cannot see what was lost.

:func:`paginate` instead *measures* content against the real space available
and splits an over-full slide into continuation slides. Nothing is discarded,
the deck grows by a slide instead. Because the measurement uses the same font
metrics the renderers draw with, the split points are accurate rather than
guessed at from item counts.
"""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import List, Optional, Sequence, Tuple

from .design import (
    GRID,
    SPACE,
    TYPE,
    Grid,
    TypeScale,
    max_lines,
    text_height,
    wrap_text,
)
from .models import Bullet, Deck, Slide, SlideType


# --------------------------------------------------------------------------
# Rectangles
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Rect:
    """An axis-aligned rectangle in points, origin at the top-left.

    Top-left origin matches PowerPoint. The PDF backend flips to
    reportlab's bottom-left origin at the very last moment, in one place.
    """

    x: float
    y: float
    w: float
    h: float

    @property
    def right(self) -> float:
        return self.x + self.w

    @property
    def bottom(self) -> float:
        return self.y + self.h

    @property
    def cx(self) -> float:
        return self.x + self.w / 2

    @property
    def cy(self) -> float:
        return self.y + self.h / 2

    def inset(self, dx: float, dy: Optional[float] = None) -> "Rect":
        """Shrink by ``dx`` horizontally and ``dy`` vertically (default ``dx``)."""
        dy = dx if dy is None else dy
        return Rect(self.x + dx, self.y + dy, max(0.0, self.w - 2 * dx), max(0.0, self.h - 2 * dy))

    def split_h(self, count: int, gap: float) -> List["Rect"]:
        """Split into ``count`` columns separated by ``gap``."""
        count = max(1, count)
        w = (self.w - gap * (count - 1)) / count
        return [Rect(self.x + i * (w + gap), self.y, w, self.h) for i in range(count)]

    def split_v(self, count: int, gap: float) -> List["Rect"]:
        """Split into ``count`` rows separated by ``gap``."""
        count = max(1, count)
        h = (self.h - gap * (count - 1)) / count
        return [Rect(self.x, self.y + i * (h + gap), self.w, h) for i in range(count)]

    def with_h(self, h: float) -> "Rect":
        return Rect(self.x, self.y, self.w, h)


# --------------------------------------------------------------------------
# Zones
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Zones:
    """The standard vertical structure of a content slide.

    Keeping this in one place is what makes titles land on the same baseline
    across every archetype — the single most noticeable sign of a deck built
    to a system rather than slide by slide.
    """

    kicker: Rect     # small eyebrow line above the title
    title: Rect      # the slide title
    rule: Rect       # thin accent rule under the title
    subtitle: Rect   # optional supporting line
    body: Rect       # everything else
    footer: Rect     # page number / deck title


def zones(
    grid: Grid = GRID,
    scale: TypeScale = TYPE,
    *,
    has_kicker: bool = False,
    has_subtitle: bool = False,
    title_lines: int = 1,
) -> Zones:
    """Compute the vertical zones for a content slide.

    The title zone grows when a title wraps to two lines, pushing the body
    down rather than letting the two collide — which is what happens when
    layout positions are hardcoded.
    """
    y = grid.top

    kicker_h = scale.caption * 1.5 if has_kicker else 0.0
    kicker = Rect(grid.left, y, grid.content_w, kicker_h)
    if has_kicker:
        y += kicker_h + SPACE.xxs

    title_h = scale.h1 * scale.line_height_tight * max(1, title_lines)
    title = Rect(grid.left, y, grid.content_w, title_h)
    y += title_h + SPACE.sm

    rule = Rect(grid.left, y, 46.0, 3.0)
    y += rule.h

    sub_h = scale.body_sm * scale.line_height * 1.2 if has_subtitle else 0.0
    subtitle = Rect(grid.left, y + SPACE.xs, grid.content_w, sub_h)
    if has_subtitle:
        y += SPACE.xs + sub_h

    body_top = y + SPACE.lg
    footer_h = scale.micro * 2.2
    body_bottom = grid.bottom - footer_h
    body = Rect(grid.left, body_top, grid.content_w, max(0.0, body_bottom - body_top))
    footer = Rect(grid.left, grid.bottom - footer_h, grid.content_w, footer_h)

    return Zones(kicker=kicker, title=title, rule=rule, subtitle=subtitle, body=body, footer=footer)


def title_line_count(title: str, grid: Grid = GRID, scale: TypeScale = TYPE, serif: bool = False) -> int:
    """How many lines the slide title will occupy (capped at 2).

    Beyond two lines a title is really a statement, and
    :attr:`~kimiagent.models.SlideType.BIG_STATEMENT` is the right archetype.
    """
    lines = len(wrap_text(title, grid.content_w, scale.h1, bold=True, serif=serif))
    return min(2, max(1, lines))


# --------------------------------------------------------------------------
# Capacity
# --------------------------------------------------------------------------
#
# Upper bounds per slide, chosen for *legibility* rather than to fit as much as
# possible. Six cards on a 16:9 slide already means ~4 inches each; nine means
# type so small the slide stops working from the back of a room.

MAX_BULLETS = 7
MAX_AGENDA = 7
MAX_STATS = 6
MAX_CARDS = 6
MAX_TIMELINE = 5
MAX_STEPS = 5
MAX_TABLE_ROWS = 9
MAX_COMPARISON_POINTS = 6

#: Suffix for slides created by splitting an over-full slide.
CONT_SUFFIX = " (cont.)"


def _continued(title: str) -> str:
    """Title for a continuation slide, without stacking suffixes."""
    if title.endswith(CONT_SUFFIX):
        return title
    return f"{title}{CONT_SUFFIX}" if title else ""


def bullet_capacity(
    bullets: Sequence[Bullet],
    body: Rect,
    scale: TypeScale = TYPE,
    serif: bool = False,
) -> int:
    """How many of ``bullets`` fit in ``body`` at the standard body size.

    Measures real wrapped height per bullet — a slide of one-word bullets
    holds more than a slide of full sentences, and counting items cannot
    tell the difference.
    """
    if not bullets:
        return 0
    available = body.h
    used = 0.0
    fitted = 0
    for b in bullets:
        size = scale.body if b.level == 0 else scale.body_sm
        indent = 26.0 + (b.level * 18.0)
        text = f"{b.bold_lead}: {b.text}" if b.bold_lead else b.text
        h = text_height(text, body.w - indent, size, bold=False, serif=serif)
        h += SPACE.sm  # gap between bullets
        if used + h > available and fitted > 0:
            break
        used += h
        fitted += 1
    return max(1, fitted)


# --------------------------------------------------------------------------
# Pagination
# --------------------------------------------------------------------------


def paginate(deck: Deck, grid: Grid = GRID, scale: TypeScale = TYPE, serif: bool = False) -> Deck:
    """Return a copy of ``deck`` with over-full slides split in place.

    Called by both renderers before drawing, so a deck renders identically
    to PPTX and PDF and neither can silently truncate. The returned deck is a
    deep copy: the caller's spec is never mutated, which keeps
    ``--emit-spec`` output faithful to what the agent actually produced.
    """
    out: List[Slide] = []
    for slide in deck.slides:
        out.extend(_split_slide(slide, grid, scale, serif))
    new = deepcopy(deck)
    new.slides = out
    return new


def _split_slide(slide: Slide, grid: Grid, scale: TypeScale, serif: bool) -> List[Slide]:
    """Split one slide into as many as its content needs."""
    t = slide.type

    if t == SlideType.BULLETS and slide.bullets:
        z = zones(
            grid,
            scale,
            has_kicker=bool(slide.kicker),
            has_subtitle=bool(slide.subtitle),
            title_lines=title_line_count(slide.title, grid, scale, serif),
        )
        return _split_by(
            slide,
            slide.bullets,
            lambda items: bullet_capacity(items, z.body, scale, serif),
            "bullets",
            hard_max=MAX_BULLETS,
        )

    if t == SlideType.AGENDA and slide.bullets:
        return _split_by(slide, slide.bullets, None, "bullets", hard_max=MAX_AGENDA)

    if t == SlideType.STATS and slide.stats:
        return _split_by(slide, slide.stats, None, "stats", hard_max=MAX_STATS)

    if t == SlideType.CARDS and slide.cards:
        return _split_by(slide, slide.cards, None, "cards", hard_max=MAX_CARDS)

    if t == SlideType.TIMELINE and slide.timeline:
        return _split_by(slide, slide.timeline, None, "timeline", hard_max=MAX_TIMELINE)

    if t == SlideType.PROCESS and slide.steps:
        return _split_by(slide, slide.steps, None, "steps", hard_max=MAX_STEPS)

    if t == SlideType.TABLE and slide.table and len(slide.table.rows) > MAX_TABLE_ROWS:
        return _split_table(slide)

    if t == SlideType.TWO_COLUMN and slide.columns:
        return _split_two_column(slide)

    if t == SlideType.COMPARISON and slide.sides:
        return _split_comparison(slide)

    return [slide]


def _split_by(
    slide: Slide,
    items: Sequence,
    capacity_fn,
    attr: str,
    hard_max: int,
) -> List[Slide]:
    """Generic splitter: chunk ``items`` across copies of ``slide``."""
    if len(items) <= hard_max and capacity_fn is None:
        return [slide]

    remaining = list(items)
    pages: List[Slide] = []
    first = True
    guard = 0
    while remaining and guard < 40:
        guard += 1
        limit = hard_max
        if capacity_fn is not None:
            limit = min(hard_max, capacity_fn(remaining))
        limit = max(1, limit)
        chunk, remaining = remaining[:limit], remaining[limit:]

        page = deepcopy(slide)
        setattr(page, attr, chunk)
        if not first:
            page.title = _continued(slide.title)
            page.notes = ""          # notes belong with the first page
            page.kicker = slide.kicker
        pages.append(page)
        first = False

    # Only the last page keeps a trailing element such as a chart takeaway.
    return pages or [slide]


def _split_table(slide: Slide) -> List[Slide]:
    """Chunk long table rows, repeating the header on every page."""
    assert slide.table is not None
    rows = slide.table.rows
    pages: List[Slide] = []
    for i in range(0, len(rows), MAX_TABLE_ROWS):
        page = deepcopy(slide)
        assert page.table is not None
        page.table.rows = rows[i : i + MAX_TABLE_ROWS]
        # A highlight index only applies to the page that actually holds it.
        hl = slide.table.highlight_row
        if hl is None or not (i <= hl < i + MAX_TABLE_ROWS):
            page.table.highlight_row = None
        else:
            page.table.highlight_row = hl - i
        if i:
            page.title = _continued(slide.title)
            page.notes = ""
        pages.append(page)
    return pages


def _split_two_column(slide: Slide) -> List[Slide]:
    """Split a two-column slide when either column overflows."""
    per_col = 6
    cols = slide.columns[:2]
    if all(len(c) <= per_col for c in cols):
        return [slide]
    depth = max((len(c) + per_col - 1) // per_col for c in cols)
    pages: List[Slide] = []
    for p in range(depth):
        page = deepcopy(slide)
        page.columns = [c[p * per_col : (p + 1) * per_col] for c in cols]
        if p:
            page.title = _continued(slide.title)
            page.notes = ""
        pages.append(page)
    return pages


def _split_comparison(slide: Slide) -> List[Slide]:
    """Trim comparison points, moving the overflow to a continuation slide."""
    if all(len(s.points) <= MAX_COMPARISON_POINTS for s in slide.sides):
        return [slide]
    depth = max(
        (len(s.points) + MAX_COMPARISON_POINTS - 1) // MAX_COMPARISON_POINTS
        for s in slide.sides
    )
    pages: List[Slide] = []
    for p in range(depth):
        page = deepcopy(slide)
        for side, original in zip(page.sides, slide.sides):
            side.points = original.points[p * MAX_COMPARISON_POINTS : (p + 1) * MAX_COMPARISON_POINTS]
        if p:
            page.title = _continued(slide.title)
            page.notes = ""
            page.verdict = ""           # the verdict belongs on the first page
        pages.append(page)
    return pages


# --------------------------------------------------------------------------
# Grid helpers for specific archetypes
# --------------------------------------------------------------------------


def bullets_height(
    bullets: Sequence[Bullet],
    width: float,
    scale: TypeScale = TYPE,
    serif: bool = False,
    compact: bool = False,
    size: Optional[float] = None,
) -> float:
    """Measured height of a rendered bullet block.

    Lets a container size itself to its content instead of stretching to the
    full body zone. A card with three bullets and six inches of empty space
    below them reads as an unfinished slide.
    """
    if not bullets:
        return 0.0
    base = size or scale.body
    gap = SPACE.xs if compact else SPACE.sm
    total = 0.0
    for b in bullets:
        sz = base if b.level == 0 else base * 0.87
        indent = 16.0 + b.level * 16.0
        text = f"{b.bold_lead}: {b.text}" if b.bold_lead else b.text
        total += text_height(text, max(1.0, width - indent), sz, False, serif) + gap
    return total


def column_card_height(
    columns: Sequence[Sequence[Bullet]],
    card_w: float,
    body: Rect,
    scale: TypeScale = TYPE,
    serif: bool = False,
    padding: float = SPACE.lg,
    minimum: float = 130.0,
) -> float:
    """Common height for side-by-side bullet cards.

    All columns share one height so their tops and bottoms align, but that
    height is driven by the tallest column's content rather than by the slide.
    """
    inner_w = max(1.0, card_w - padding * 2)
    needed = max(
        (bullets_height(col, inner_w, scale, serif) for col in columns),
        default=0.0,
    )
    return max(minimum, min(body.h, needed + padding * 2 + 8.0))


def card_grid(body: Rect, count: int, gap: float = SPACE.md, max_per_row: int = 3) -> List[Rect]:
    """Lay ``count`` cards into a balanced grid inside ``body``.

    Balance matters: four cards read better as 2x2 than as 3+1, because a
    lone trailing card on its own row looks like a mistake. Row counts are
    chosen to keep rows even.
    """
    count = max(0, count)
    if count == 0:
        return []
    if count <= max_per_row:
        rows = [count]
    elif count == 4:
        rows = [2, 2]
    elif count <= 6:
        rows = [(count + 1) // 2, count // 2]
    else:
        per = max_per_row
        rows = []
        left = count
        while left > 0:
            rows.append(min(per, left))
            left -= min(per, left)

    row_rects = body.split_v(len(rows), gap)
    out: List[Rect] = []
    for n, rr in zip(rows, row_rects):
        # Keep a short final row the same card width as full rows, centred,
        # so card sizes stay consistent across the slide.
        full = max(rows)
        card_w = (rr.w - gap * (full - 1)) / full
        total = n * card_w + gap * (n - 1)
        x0 = rr.x + (rr.w - total) / 2
        for i in range(n):
            out.append(Rect(x0 + i * (card_w + gap), rr.y, card_w, rr.h))
    return out


def stat_grid(body: Rect, count: int, gap: float = SPACE.md) -> List[Rect]:
    """Lay KPI cards out, wrapping to a second row beyond four.

    Five or six metrics on one row produce columns too narrow for the
    oversized numerals that make a KPI slide work, so they wrap instead.
    """
    if count <= 0:
        return []
    if count <= 4:
        # A single centred row, vertically inset so the numbers sit optically
        # centred rather than floating at the top of the body zone.
        h = min(body.h, 200.0)
        row = Rect(body.x, body.y + (body.h - h) / 2, body.w, h)
        return row.split_h(count, gap)
    return card_grid(body, count, gap=gap, max_per_row=3)


__all__ = [
    "Rect",
    "Zones",
    "zones",
    "title_line_count",
    "bullet_capacity",
    "bullets_height",
    "column_card_height",
    "paginate",
    "card_grid",
    "stat_grid",
    "MAX_BULLETS",
    "MAX_AGENDA",
    "MAX_STATS",
    "MAX_CARDS",
    "MAX_TIMELINE",
    "MAX_STEPS",
    "MAX_TABLE_ROWS",
    "MAX_COMPARISON_POINTS",
    "CONT_SUFFIX",
]
