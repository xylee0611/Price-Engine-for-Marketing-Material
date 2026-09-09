import fitz

from placement import find_clear_position, pdf_label_size

GAP_PT = 12


def overlay_pdf(original_pdf_path, output_pdf_path, occurrences, prices, currency, obstacles_by_page, page_sizes):
    doc = fitz.open(original_pdf_path)
    placed = []
    skipped = []

    for occ in occurrences:
        sku = occ["sku"]
        if sku not in prices:
            continue
        pdf_info = occ.get("pdf")
        if pdf_info is None:
            skipped.append({"sku": sku, "reason": "No position available in the PDF (and no PPTX to map from)."})
            continue

        page_idx = pdf_info["page_idx"]
        anchor = pdf_info["bbox"]
        fontsize = pdf_info["font_size"] or 18
        label = f"{currency} {prices[sku]:.1f}"

        label_w, label_h = pdf_label_size(label, fontsize)
        bounds = page_sizes[page_idx]
        obstacles = obstacles_by_page.get(page_idx, [])

        rect = find_clear_position(anchor, label_w, label_h, obstacles, bounds, GAP_PT)
        if rect is None:
            skipped.append({"sku": sku, "reason": f"No clear space found near SKU on PDF page {page_idx + 1}."})
            continue

        page = doc[page_idx]
        baseline = fitz.Point(rect[0], rect[1] + fontsize * 0.9)
        page.insert_text(baseline, label, fontsize=fontsize, fontname="helv", color=(0, 0, 0))
        placed.append({"sku": sku, "page": page_idx + 1, "label": label})

    doc.save(output_pdf_path)
    doc.close()
    return placed, skipped
