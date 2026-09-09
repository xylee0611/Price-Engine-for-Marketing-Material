"""Cross-reference PPTX/PDF tokens against the price list to build confirmed
SKU occurrences, following the PPTX-first / PDF-position-preferred strategy.
"""
from price_engine import SKU_TOKEN_RE
from sku_extractor import emu_bbox_to_points

ASPECT_RATIO_TOLERANCE = 0.02  # 2% — sanity check before trusting EMU->pt mapping


def build_sku_occurrences(pptx_data, pdf_data, price_list):
    """
    pptx_data: (tokens, obstacles_by_slide, slide_width, slide_height) or None
    pdf_data: (tokens, obstacles_by_page, page_sizes) or None

    Returns (occurrences, unmatched, warnings):
      occurrences: list of dicts, one per confirmed SKU found anywhere:
        {
          sku, matched_via, record,
          pptx: {slide_idx, left, top, width, height, font_size, font_name} or None,
          pdf: {page_idx, bbox, font_size, position_source} or None,
              position_source is "direct" (found as real PDF text) or
              "mapped" (converted from the PPTX shape's position)
        }
      unmatched: list of {"source": "pptx"/"pdf", "location": slide/page idx, "text": str}
      warnings: list of str (e.g. slide/page count mismatch notices)
    """
    occurrences = {}
    unmatched = []
    warnings = []

    pptx_tokens = pptx_data[0] if pptx_data else []
    pdf_tokens = pdf_data[0] if pdf_data else []

    # Index PDF tokens by text for fast direct lookup regardless of page numbering.
    pdf_tokens_by_text = {}
    for tok in pdf_tokens:
        pdf_tokens_by_text.setdefault(tok["text"], []).append(tok)

    # 1) Confirm SKUs from the PPTX (primary source when available).
    for tok in pptx_tokens:
        text = tok["text"]
        if not SKU_TOKEN_RE.match(text):
            continue
        record, matched_via = price_list.lookup(text)
        if record is None:
            unmatched.append({"source": "pptx", "location": tok["slide_idx"], "text": text})
            continue
        occ = occurrences.setdefault(text, {"sku": text, "matched_via": matched_via, "record": record, "pptx": None, "pdf": None})
        occ["pptx"] = tok

    # 2) Confirm SKUs from the PDF directly (covers PDF-only mode, and gives
    #    pixel-exact position even when a PPTX is also present).
    for tok in pdf_tokens:
        text = tok["text"]
        if not SKU_TOKEN_RE.match(text):
            continue
        record, matched_via = price_list.lookup(text)
        if record is None:
            unmatched.append({"source": "pdf", "location": tok["page_idx"], "text": text})
            continue
        occ = occurrences.setdefault(text, {"sku": text, "matched_via": matched_via, "record": record, "pptx": None, "pdf": None})
        if occ["pdf"] is None:  # keep first occurrence if a SKU repeats
            occ["pdf"] = {
                "page_idx": tok["page_idx"],
                "bbox": tok["bbox"],
                "font_size": tok["font_size"],
                "position_source": "direct",
            }

    # 3) For SKUs confirmed via PPTX but with no direct PDF text match, fall
    #    back to mapping the PPTX shape's position onto the matching PDF page.
    if pptx_data and pdf_data:
        _, _, slide_w, slide_h = pptx_data
        _, _, page_sizes = pdf_data
        slide_count = len({t["slide_idx"] for t in pptx_tokens}) if pptx_tokens else None

        for sku, occ in occurrences.items():
            if occ["pdf"] is not None or occ["pptx"] is None:
                continue
            slide_idx = occ["pptx"]["slide_idx"]
            page_size = page_sizes.get(slide_idx)
            if page_size is None:
                warnings.append(f"SKU {sku}: slide {slide_idx + 1} has no matching PDF page — cannot map position.")
                continue

            slide_ratio = slide_w / slide_h
            page_ratio = page_size[0] / page_size[1]
            if abs(slide_ratio - page_ratio) / slide_ratio > ASPECT_RATIO_TOLERANCE:
                warnings.append(
                    f"SKU {sku}: PDF page {slide_idx + 1} aspect ratio doesn't match the PPTX slide — "
                    "skipping position mapping (PDF likely isn't a direct export of this PPTX)."
                )
                continue

            bbox = emu_bbox_to_points(
                occ["pptx"]["left"], occ["pptx"]["top"], occ["pptx"]["width"], occ["pptx"]["height"]
            )
            occ["pdf"] = {
                "page_idx": slide_idx,
                "bbox": bbox,
                "font_size": (occ["pptx"]["font_size"] / 12700) if occ["pptx"]["font_size"] else 18,
                "position_source": "mapped",
            }

    return list(occurrences.values()), unmatched, warnings
