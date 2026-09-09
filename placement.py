"""Collision-aware placement: given a SKU's anchor position and the other
elements on its page/slide, find a clear spot for the price label — right,
then below, then left. Returns None (caller flags for manual review) if none
of the three are clear.
"""
import fitz

_HELV = fitz.Font("helv")


def rects_overlap(a, b):
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    return not (ax1 <= bx0 or bx1 <= ax0 or ay1 <= by0 or by1 <= ay0)


def pdf_label_size(text, fontsize):
    width = _HELV.text_length(text, fontsize=fontsize)
    height = fontsize * 1.2
    return width, height


def pptx_label_size_emu(text, font_size_emu):
    font_size_pt = (font_size_emu / 12700) if font_size_emu else 18
    width_pt = len(text) * font_size_pt * 0.55
    height_pt = font_size_pt * 1.3
    return width_pt * 12700, height_pt * 12700


def find_clear_position(anchor_bbox, label_width, label_height, obstacles, bounds, gap):
    """anchor_bbox/obstacles/bounds/gap must all be in the same unit
    (points for PDF, EMU for PPTX). Returns a (x0,y0,x1,y1) rect or None."""
    x0, y0, x1, y1 = anchor_bbox
    bound_w, bound_h = bounds

    candidates = [
        (x1 + gap, y0, x1 + gap + label_width, y0 + label_height),  # right
        (x0, y1 + gap, x0 + label_width, y1 + gap + label_height),  # below
        (x0 - gap - label_width, y0, x0 - gap, y0 + label_height),  # left
    ]

    other_obstacles = [o for o in obstacles if o != anchor_bbox]

    for cand in candidates:
        cx0, cy0, cx1, cy1 = cand
        if cx0 < 0 or cy0 < 0 or cx1 > bound_w or cy1 > bound_h:
            continue
        if any(rects_overlap(cand, obs) for obs in other_obstacles):
            continue
        return cand
    return None
