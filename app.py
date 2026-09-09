import os
import tempfile
import uuid

import pandas as pd
import streamlit as st

from matcher import build_sku_occurrences
from pdf_writer import overlay_pdf
from pptx_writer import overlay_pptx
from price_engine import (
    PriceList,
    compute_custom_price,
    compute_price,
    currency_for_tier,
    currency_from_header,
    discount_factor,
)
from sku_extractor import extract_pdf_tokens, extract_pptx_tokens

CUSTOM_TIER_LABEL = "Custom tier..."

st.set_page_config(page_title="Price Engine for Marketing Material", layout="wide")
st.title("Price Engine for Marketing Material")
st.caption(
    "Upload a price list plus a PDF and/or PPTX deck. Every SKU code found in the "
    "deck is matched against the price list and priced at the tier/currency you choose."
)

ROUNDING_OPTIONS = {
    "Round up to nearest 0.5": "up_half",
    "Round to nearest whole number": "nearest_int",
    "No rounding": "none",
}


def get_session_dir():
    if "session_dir" not in st.session_state:
        st.session_state.session_dir = tempfile.mkdtemp(prefix=f"apply_price_{uuid.uuid4().hex[:8]}_")
    return st.session_state.session_dir


def save_upload(uploaded_file, session_dir):
    path = os.path.join(session_dir, uploaded_file.name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return path


session_dir = get_session_dir()

st.header("1. Input files")
col1, col2, col3 = st.columns(3)
with col1:
    price_list_file = st.file_uploader("Price list (.xlsx)", type=["xlsx"])
with col2:
    pptx_file = st.file_uploader("Material PPTX (optional)", type=["pptx"])
with col3:
    pdf_file = st.file_uploader("Material PDF (optional)", type=["pdf"])

if not price_list_file:
    st.info("Upload a price list to continue.")
    st.stop()

if not pptx_file and not pdf_file:
    st.warning("Upload at least one of PPTX or PDF. Providing both is recommended — "
               "it's more robust to unusual layouts and lets us produce both deliverables.")
    st.stop()

price_list_path = save_upload(price_list_file, session_dir)
try:
    price_list = PriceList(price_list_path)
except ValueError as e:
    st.error(str(e))
    st.stop()

st.success(f"Loaded **{price_list_file.name}** — {price_list.row_count} SKUs found on sheet \"{price_list.sheet_name}\".")

pptx_path = save_upload(pptx_file, session_dir) if pptx_file else None
pdf_path = save_upload(pdf_file, session_dir) if pdf_file else None

st.header("2. Pricing settings")
tiers = price_list.available_tiers()
if not tiers:
    st.error("This price list doesn't have any recognized price columns (RRP/RLP/P1-P3 for USD/MYR/SGD/RMB).")
    st.stop()

colA, colB = st.columns(2)
with colA:
    tier = st.selectbox("Pricing tier", tiers + [CUSTOM_TIER_LABEL])
with colB:
    rounding_label = st.selectbox("Rounding", list(ROUNDING_OPTIONS.keys()))
rounding_mode = ROUNDING_OPTIONS[rounding_label]

custom_base_header = custom_factor = custom_label = None
if tier == CUSTOM_TIER_LABEL:
    st.info(
        "Not every tier your price list might need is pre-defined — this is for those cases. "
        "State exactly what you want computed; nothing here is guessed from the data."
    )
    base_options = price_list.available_base_headers()
    if not base_options:
        st.error("No known price columns found to build a custom tier from.")
        st.stop()

    custom_mode = st.radio(
        "Type", ["Discount (cascading %)", "Markup (× factor)"], horizontal=True
    )

    if custom_mode == "Discount (cascading %)":
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            custom_base_header = st.selectbox("Base column", base_options)
        with c2:
            custom_pct = st.number_input("Discount % per step", min_value=0.0, max_value=100.0, value=5.0, step=0.5)
        with c3:
            custom_steps = st.number_input("Cascading steps", min_value=1, max_value=20, value=1, step=1)
        with c4:
            default_label = f"Custom {currency_from_header(custom_base_header)}"
            custom_label = st.text_input("Label for this tier", value=default_label)
        custom_factor = discount_factor(custom_pct, custom_steps)
        st.caption(f"Formula: **{custom_base_header} × (1 − {custom_pct}%)^{int(custom_steps)}** = ×{custom_factor:.4f}")
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            custom_base_header = st.selectbox("Base column", base_options)
        with c2:
            custom_factor = st.number_input(
                "Markup multiplier (e.g. 1.2 = base × 1.2)", min_value=0.0, value=1.2, step=0.05
            )
        with c3:
            default_label = f"Markup {custom_factor}x {currency_from_header(custom_base_header)}"
            custom_label = st.text_input("Label for this tier", value=default_label)
        st.caption(f"Formula: **{custom_base_header} × {custom_factor}**")

    currency = currency_from_header(custom_base_header)
    tier_display = custom_label
else:
    currency = currency_for_tier(tier)
    tier_display = tier

if st.button("Process", type="primary"):
    pptx_data = None
    pdf_data = None
    if pptx_path:
        tokens, obstacles, slide_w, slide_h = extract_pptx_tokens(pptx_path)
        pptx_data = (tokens, obstacles, slide_w, slide_h)
    if pdf_path:
        tokens, obstacles, page_sizes = extract_pdf_tokens(pdf_path)
        pdf_data = (tokens, obstacles, page_sizes)

    occurrences, unmatched, warnings = build_sku_occurrences(pptx_data, pdf_data, price_list)

    prices = {}
    unpriceable = []
    rows = []
    for occ in occurrences:
        if tier == CUSTOM_TIER_LABEL:
            raw, final = compute_custom_price(occ["record"], custom_base_header, custom_factor, rounding_mode)
        else:
            raw, final = compute_price(occ["record"], tier, rounding_mode)
        if final is None:
            unpriceable.append(occ["sku"])
            continue
        prices[occ["sku"]] = final
        rows.append(
            {
                "SKU": occ["sku"],
                "Matched via": occ["matched_via"],
                "Description": occ["record"].get("SKU Description", ""),
                "Price": f"{currency} {final:.1f}",
                "PPTX slide": (occ["pptx"]["slide_idx"] + 1) if occ["pptx"] else "—",
                "PDF page": (occ["pdf"]["page_idx"] + 1) if occ["pdf"] else "—",
            }
        )

    st.session_state.result = {
        "rows": rows,
        "unmatched": unmatched,
        "unpriceable": unpriceable,
        "warnings": warnings,
        "occurrences": occurrences,
        "prices": prices,
        "pptx_data": pptx_data,
        "pdf_data": pdf_data,
        "pptx_path": pptx_path,
        "pdf_path": pdf_path,
        "currency": currency,
        "tier": tier_display,
    }

result = st.session_state.get("result")
if result:
    st.header("3. Review")
    st.write(
        f"**{len(result['rows'])}** SKUs matched and priced at **{result['tier']}** "
        f"out of {len(result['occurrences']) + len(result['unpriceable'])} SKUs found in the material."
    )

    if result["rows"]:
        st.dataframe(pd.DataFrame(result["rows"]), width='stretch', hide_index=True)
    else:
        st.warning("No SKUs from the material matched the price list.")

    if result["unmatched"]:
        st.error(f"{len(result['unmatched'])} number(s) found in the material don't match any SKU in the price list — "
                 "excluded from output, not guessed:")
        st.dataframe(pd.DataFrame(result["unmatched"]), width='stretch', hide_index=True)

    if result["unpriceable"]:
        st.error(f"{len(result['unpriceable'])} matched SKU(s) have no value for tier \"{result['tier']}\" in this "
                 f"price list — excluded from output: {', '.join(result['unpriceable'])}")

    if result["warnings"]:
        for w in result["warnings"]:
            st.warning(w)

    if result["rows"]:
        st.header("4. Generate & download")
        if st.button("Generate output files"):
            placed_pdf, skipped_pdf = [], []
            placed_pptx, skipped_pptx = [], []

            out_pdf_path = out_pptx_path = None
            if result["pdf_path"]:
                out_pdf_path = os.path.join(session_dir, "priced_output.pdf")
                placed_pdf, skipped_pdf = overlay_pdf(
                    result["pdf_path"], out_pdf_path, result["occurrences"], result["prices"],
                    result["currency"], result["pdf_data"][1], result["pdf_data"][2],
                )
            if result["pptx_path"]:
                out_pptx_path = os.path.join(session_dir, "priced_output.pptx")
                placed_pptx, skipped_pptx = overlay_pptx(
                    result["pptx_path"], out_pptx_path, result["occurrences"], result["prices"],
                    result["currency"], result["pptx_data"][1], result["pptx_data"][2], result["pptx_data"][3],
                )

            needs_review = {s["sku"]: s["reason"] for s in skipped_pdf}
            for s in skipped_pptx:
                needs_review.setdefault(s["sku"], s["reason"])

            if needs_review:
                st.warning(
                    f"{len(needs_review)} SKU(s) couldn't be safely auto-placed and were skipped "
                    "(no price stamped over a photo/other text) — needs manual placement:"
                )
                st.dataframe(
                    pd.DataFrame([{"SKU": k, "Reason": v} for k, v in needs_review.items()]),
                    width='stretch', hide_index=True,
                )

            dl_col1, dl_col2 = st.columns(2)
            if out_pdf_path:
                with dl_col1:
                    with open(out_pdf_path, "rb") as f:
                        st.download_button("Download priced PDF", f, file_name=f"priced - {result['tier']}.pdf")
                    st.caption(f"{len(placed_pdf)} price(s) placed on the PDF.")
            if out_pptx_path:
                with dl_col2:
                    with open(out_pptx_path, "rb") as f:
                        st.download_button("Download priced PPTX", f, file_name=f"priced - {result['tier']}.pptx")
                    st.caption(f"{len(placed_pptx)} price(s) placed on the PPTX.")
