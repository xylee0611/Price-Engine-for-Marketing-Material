# Price Engine for Marketing Material

A Streamlit app that applies price-list prices onto SKU marketing decks (PDF and/or PPTX). Upload a price list and a deck, pick a pricing tier/currency, review every matched SKU and its computed price, then download the priced PDF/PPTX with the price stamped next to each SKU code.

## How it works

- **SKU detection** cross-references every plausible text token found in the deck against the price list's SKU Code / Combine SKU Code columns — layout-independent, no hardcoded SKU list. When both a PPTX and PDF are provided, the PPTX (almost always real text) drives detection, and the PDF's own text is preferred for placement precision.
- **Pricing** supports direct price-list columns (RRP/RLP per currency, P1–P3 MYR), documented cascading-discount tiers (R1–R6 USD, R4–R6 MYR), and a **Custom tier** for anything else — either a cascading discount (base column × (1−%)ⁿ) or a direct markup multiplier (e.g. RRP × 1.2) — always user-specified, never guessed.
- **Placement** is collision-aware: it tries a clear spot to the right, then below, then left of each SKU code, and skips (flagging for manual review) rather than stamping a price over a photo or other text.
- Every final price rounds up to the nearest 0.5 by default (configurable per run).

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (default `http://localhost:8501`).

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI |
| `price_engine.py` | Price list loading, tier/currency formulas, rounding |
| `sku_extractor.py` | Pulls text tokens + obstacle boxes from PPTX/PDF |
| `matcher.py` | Cross-references tokens against the price list |
| `placement.py` | Collision-aware placement algorithm |
| `pdf_writer.py` / `pptx_writer.py` | Generate the priced output files |

## Adding a new fixed tier

Edit `price_engine.py`:
- A tier that's a direct column in your price list → add a line to `DIRECT_TIER_HEADERS`.
- A tier that's a confirmed cascading formula (base column × (1−%)ⁿ) → add a line to `CALCULATED_TIER_DEFS`.

Never add a tier here without an explicit, confirmed formula — use the in-app "Custom tier" option for anything not yet confirmed as a standing rule.
