# VC Automation Dashboard

A Streamlit app for automating Lazada & Zalora Voucher Code (VC) eligibility processing.

---

## Setup & Run

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the app
```bash
streamlit run app.py
```

The app opens at **http://localhost:8501**

---

## Files Required

| File | Format | Purpose |
|------|--------|---------|
| Lazada / Zalora SC Report | CSV or XLSX | Main working sheet |
| Ecom Tracker | XLSX (multi-tab) | Source of truth for product data |
| Content File | CSV or XLSX | ALU (Color No) reference |
| AM Exclusion Sheet | XLSX (multi-tab) | Optional voucher exclusions |
| Price Tier Reference | Text input | e.g. `AX - BA` |

---

## Column Mapping Reference

| Concept | SC Report | Ecom Tracker | Content File |
|---------|-----------|--------------|--------------|
| ALU | (derived via EAN) | STYLE# | Color No |
| EAN / Barcode | SellerSKU | EAN / SellerSKU | — |
| Eligible flag | ShopSKU | — | — |

---

## Filter Logic (in order)

1. **Launch Date** — Exclude: blank, 00-Jan-1900, "past season", "TBC", text-only values, dates after cutoff → Future Launch
2. **Ecom Status** — Keep only `YES`
3. **RRP** — Exclude ≤ 16 and 0
4. **SRP** — Exclude ≤ 16 (0 is allowed)
5. **Exclusion keyword** — Strict exact match against inclusion keywords per voucher
6. **AM Exclusion** — Per-ALU remarks applied per voucher percentage rules

---

## AM Exclusion Rules

| Remark contains | Effect |
|----------------|--------|
| "Exclude from all Voucher, Markdown, Campaign" | Excluded from every VC |
| "Exclude from platform voucher" | Excluded from all platform VCs |
| "Exclude from 10% VC" / "below" | Excluded from ≤ 10% VCs |
| "Exclude 20% and above" | Excluded from ≥ 20% VCs |

---

## Output

Downloaded as `VC_Output_<Marketplace>_<Region>_<Date>.xlsx`  
Contains the original SC report columns + added columns + one column per voucher name with `YES` for eligible SKUs.
