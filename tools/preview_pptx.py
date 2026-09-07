"""Dev-only: rasterize a .pptx to PNG thumbnails for visual QA.

This is NOT part of the KimiAgent package — it's an internal helper used to
eyeball decks in environments without PowerPoint/LibreOffice. It reads the real
shapes from the generated file (geometry, solid + gradient fills, ovals,
rounded rectangles, embedded pictures and text) and approximates their
appearance with Pillow. Gradients/shadows/rounded corners are approximated.

Usage:
    python tools/preview_pptx.py deck.pptx out_dir [--scale 120]
"""

import io
import os
import sys

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

NOTO = "/usr/share/fonts/google-noto"
FONTS = {
    "regular": f"{NOTO}/NotoSans-Regular.ttf",
    "bold": f"{NOTO}/NotoSans-Bold.ttf",
    "medium": f"{NOTO}/NotoSans-Medium.ttf",
}
SCALE = 120  # px per inch


def font(bold=False, size=18):
    path = FONTS["bold"] if bold else FONTS["regular"]
    try:
        return ImageFont.truetype(path, max(6, int(size)))
    except Exception:
        return ImageFont.load_default()


def emu_px(v):
    return int(v / 914400.0 * SCALE)


def rgb_of(color):
    try:
        c = color.rgb
        return (c[0], c[1], c[2])
    except Exception:
        return None


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def draw_gradient(img, box, c1, c2, horizontal=False):
    x0, y0, x1, y1 = box
    w, h = max(1, x1 - x0), max(1, y1 - y0)
    grad = Image.new("RGB", (w, h))
    dr = ImageDraw.Draw(grad)
    n = w if horizontal else h
    for i in range(n):
        t = i / max(1, n - 1)
        col = lerp(c1, c2, t)
        if horizontal:
            dr.line([(i, 0), (i, h)], fill=col)
        else:
            dr.line([(0, i), (w, i)], fill=col)
    img.paste(grad, (x0, y0))


def shape_fill_colors(shape):
    """Return ('solid', color) or ('gradient', c1, c2, horizontal) or None."""
    try:
        fill = shape.fill
        ftype = fill.type
    except Exception:
        return None
    # 1 == solid, 3 == gradient (MSO_FILL)
    if ftype == 1:
        c = rgb_of(fill.fore_color)
        return ("solid", c) if c else None
    if ftype == 3:
        try:
            stops = list(fill.gradient_stops)
            c1 = rgb_of(stops[0].color) or (200, 200, 200)
            c2 = rgb_of(stops[-1].color) or (120, 120, 120)
            horizontal = False
            try:
                ang = fill.gradient_angle
                horizontal = (45 <= (ang % 180) <= 135) is False
            except Exception:
                pass
            return ("gradient", c1, c2, horizontal)
        except Exception:
            return None
    return None


def draw_text(draw, shape, sx, sy):
    try:
        tf = shape.text_frame
    except Exception:
        return
    left = emu_px(shape.left) + sx
    top = emu_px(shape.top) + sy
    width = emu_px(shape.width)
    height = emu_px(shape.height)
    # collect lines
    lines = []
    for para in tf.paragraphs:
        txt = "".join(r.text for r in para.runs) or para.text
        if not txt:
            continue
        size = 18
        bold = False
        color = (30, 30, 30)
        align = para.alignment
        runs = para.runs or []
        if runs:
            r0 = runs[0]
            if r0.font.size:
                size = r0.font.size.pt * SCALE / 72.0
            bold = bool(r0.font.bold)
            c = rgb_of(r0.font.color)
            if c:
                color = c
        lines.append((txt, size, bold, color, align))
    if not lines:
        return
    total_h = sum(s for _, s, _, _, _ in lines) * 1.25
    try:
        anchor = tf.vertical_anchor
    except Exception:
        anchor = None
    y = top
    if anchor == MSO_ANCHOR.MIDDLE:
        y = top + max(0, (height - total_h) / 2)
    for txt, size, bold, color, align in lines:
        f = font(bold, size)
        try:
            tw = draw.textlength(txt, font=f)
        except Exception:
            tw = len(txt) * size * 0.5
        x = left
        if align == PP_ALIGN.CENTER:
            x = left + max(0, (width - tw) / 2)
        elif align == PP_ALIGN.RIGHT:
            x = left + max(0, width - tw)
        draw.text((x, y), txt, fill=color, font=f)
        y += size * 1.3


def render_slide(slide, sw, sh):
    img = Image.new("RGB", (sw, sh), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    # slide background (best effort)
    try:
        bg = shape_fill_colors_bg(slide)
        if bg:
            draw.rectangle([0, 0, sw, sh], fill=bg)
    except Exception:
        pass
    for shape in slide.shapes:
        try:
            x0, y0 = emu_px(shape.left), emu_px(shape.top)
            x1, y1 = x0 + emu_px(shape.width), y0 + emu_px(shape.height)
        except Exception:
            continue
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                blob = shape.image.blob
                im = Image.open(io.BytesIO(blob)).convert("RGB")
                im = im.resize((max(1, x1 - x0), max(1, y1 - y0)))
                img.paste(im, (x0, y0))
            except Exception:
                pass
            continue
        fill = shape_fill_colors(shape)
        is_oval = False
        is_round = False
        try:
            st = shape.auto_shape_type
            is_oval = (st is not None and "OVAL" in str(st))
            is_round = (st is not None and "ROUNDED" in str(st))
        except Exception:
            pass
        if fill:
            if fill[0] == "solid":
                if is_oval:
                    draw.ellipse([x0, y0, x1, y1], fill=fill[1])
                elif is_round:
                    r = max(2, int(min(x1 - x0, y1 - y0) * 0.12))
                    draw.rounded_rectangle([x0, y0, x1, y1], radius=r, fill=fill[1])
                else:
                    draw.rectangle([x0, y0, x1, y1], fill=fill[1])
            elif fill[0] == "gradient":
                draw_gradient(img, (x0, y0, x1, y1), fill[1], fill[2], fill[3])
        # text on top
        draw_text(draw, shape, 0, 0)
    # tables
    for shape in slide.shapes:
        if shape.has_table:
            draw_table(draw, shape)
    return img


def shape_fill_colors_bg(slide):
    try:
        fill = slide.background.fill
        if fill.type == 1:
            return rgb_of(fill.fore_color)
    except Exception:
        return None
    return None


def draw_table(draw, shape):
    try:
        table = shape.table
    except Exception:
        return
    x0, y0 = emu_px(shape.left), emu_px(shape.top)
    rows = list(table.rows)
    cols = list(table.columns)
    widths = [emu_px(c.width) for c in cols]
    heights = [emu_px(r.height) for r in rows]
    yy = y0
    for ri, row in enumerate(rows):
        xx = x0
        for ci in range(len(cols)):
            cell = table.cell(ri, ci)
            cw, ch = widths[ci], heights[ri]
            col = rgb_of(cell.fill.fore_color) if cell.fill.type == 1 else (240, 240, 240)
            draw.rectangle([xx, yy, xx + cw, yy + ch], fill=col)
            txt = cell.text
            if txt:
                r0 = cell.text_frame.paragraphs[0].runs
                bold = bool(r0 and r0[0].font.bold)
                color = (rgb_of(r0[0].font.color) if r0 else None) or (30, 30, 30)
                f = font(bold, 15 * SCALE / 72.0)
                draw.text((xx + 10, yy + ch / 2 - 10), txt, fill=color, font=f)
            xx += cw
        yy += heights[ri]


def main():
    path = sys.argv[1]
    out_dir = sys.argv[2] if len(sys.argv) > 2 else "preview"
    os.makedirs(out_dir, exist_ok=True)
    prs = Presentation(path)
    sw = emu_px(prs.slide_width)
    sh = emu_px(prs.slide_height)
    thumbs = []
    for i, slide in enumerate(prs.slides):
        img = render_slide(slide, sw, sh)
        p = os.path.join(out_dir, f"slide_{i+1:02d}.png")
        img.save(p)
        thumbs.append(img)
    # contact sheet
    cols = 3
    rows = (len(thumbs) + cols - 1) // cols
    tw, th = sw // 3, sh // 3
    sheet = Image.new("RGB", (cols * tw + (cols + 1) * 12, rows * th + (rows + 1) * 12), (238, 238, 238))
    for i, im in enumerate(thumbs):
        r, c = divmod(i, cols)
        t = im.resize((tw, th))
        sheet.paste(t, (12 + c * (tw + 12), 12 + r * (th + 12)))
    sheet.save(os.path.join(out_dir, "contact_sheet.png"))
    print(f"wrote {len(thumbs)} slides + contact_sheet.png to {out_dir}/")


if __name__ == "__main__":
    main()
