"""Extract text tokens and obstacle bounding boxes from PPTX slides and PDF
pages. Deliberately does not filter by "looks like a SKU" here — every token
is collected, and price_engine.PriceList decides what's a real SKU.
"""
import fitz
from pptx import Presentation

EMU_PER_POINT = 12700


def _walk_shapes(shapes):
    for shp in shapes:
        if shp.shape_type == 6:  # GROUP — recurse (titles/SKUs are sometimes grouped)
            yield from _walk_shapes(shp.shapes)
        else:
            yield shp


def extract_pptx_tokens(pptx_path):
    """Return (tokens, obstacles_by_slide).

    tokens: list of dicts {slide_idx, text, left, top, width, height, font_size, font_name}
            (left/top/width/height in EMU)
    obstacles_by_slide: {slide_idx: [(left, top, width, height), ...]} for every
            shape on the slide (text or not), used for collision checks.
    """
    prs = Presentation(pptx_path)
    tokens = []
    obstacles_by_slide = {}

    for slide_idx, slide in enumerate(prs.slides):
        obstacles = []
        for shp in _walk_shapes(slide.shapes):
            try:
                bbox = (shp.left, shp.top, shp.width, shp.height)
            except Exception:
                bbox = None
            if bbox and None not in bbox:
                obstacles.append(bbox)

            if not shp.has_text_frame:
                continue
            text = shp.text_frame.text.strip()
            if not text or bbox is None or None in bbox:
                continue
            font_size = None
            font_name = None
            paras = shp.text_frame.paragraphs
            if paras and paras[0].runs:
                run = paras[0].runs[0]
                font_size = run.font.size
                font_name = run.font.name
            tokens.append(
                {
                    "slide_idx": slide_idx,
                    "text": text,
                    "left": bbox[0],
                    "top": bbox[1],
                    "width": bbox[2],
                    "height": bbox[3],
                    "font_size": font_size,
                    "font_name": font_name,
                }
            )
        obstacles_by_slide[slide_idx] = obstacles

    return tokens, obstacles_by_slide, prs.slide_width, prs.slide_height


def extract_pdf_tokens(pdf_path):
    """Return (tokens, obstacles_by_page, page_sizes).

    tokens: list of dicts {page_idx, text, bbox=(x0,y0,x1,y1), font_size, font_name}
            (points)
    obstacles_by_page: {page_idx: [(x0,y0,x1,y1), ...]} for every text span and
            image block on the page.
    page_sizes: {page_idx: (width, height)} in points.
    """
    doc = fitz.open(pdf_path)
    tokens = []
    obstacles_by_page = {}
    page_sizes = {}

    for page_idx, page in enumerate(doc):
        page_sizes[page_idx] = (page.rect.width, page.rect.height)
        obstacles = []
        d = page.get_text("dict")
        for block in d["blocks"]:
            if block.get("type") == 1:  # image block
                obstacles.append(tuple(block["bbox"]))
                continue
            for line in block.get("lines", []):
                for span in line["spans"]:
                    bbox = tuple(span["bbox"])
                    obstacles.append(bbox)
                    text = span["text"].strip()
                    if not text:
                        continue
                    tokens.append(
                        {
                            "page_idx": page_idx,
                            "text": text,
                            "bbox": bbox,
                            "font_size": span["size"],
                            "font_name": span["font"],
                        }
                    )
        obstacles_by_page[page_idx] = obstacles

    doc.close()
    return tokens, obstacles_by_page, page_sizes


def emu_bbox_to_points(left, top, width, height):
    x0 = left / EMU_PER_POINT
    y0 = top / EMU_PER_POINT
    x1 = (left + width) / EMU_PER_POINT
    y1 = (top + height) / EMU_PER_POINT
    return (x0, y0, x1, y1)
