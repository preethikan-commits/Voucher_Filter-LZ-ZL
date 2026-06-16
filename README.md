# VC Automation Dashboard — Streamlit App

## Setup

### 1. Install Python dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the app
```bash
streamlit run app.py
```
The browser will open at `http://localhost:8501`.

---

## File Requirements

| Upload Field | Format | Key Columns Needed |
|---|---|---|
| **Lazada / Zalora SC Report** | CSV or XLSX | `SellerSKU` (= EAN) |
| **Ecom Tracker** | XLSX | Sheet named with Region (`SG`/`MY`); `STYLE#`, `Launch Dates`, `Ecom Status`, RRP/SRP/DISC/Exclusion columns |
| **Ecom Tracker Price Tier Reference** | CSV or XLSX | `Reference` column (e.g. `AX - BA`), then column names that map to Ecom Tracker |
| **Content File** | CSV or XLSX | `Color No` (= ALU) + `SellerSKU` / `EAN` |
| **AM Exclusion Sheet** *(optional)* | XLSX | Sheet: `SG VC Exclusions` or `MY VC Exclusions`; `Article` (= ALU), `Exclusion Type` / `Remarks` |

---

## Logic Summary

1. **EAN → ALU mapping**: Content file `SellerSKU` → `Color No`.
2. **ALU → Ecom data**: Ecom Tracker `STYLE#` → Launch date, Ecom Status, RRP, SRP, DISC%, Exclusion.
3. **Price Tier**: Looks up the provided reference (e.g. `AX - BA`) in the Price Tier Reference file to resolve which columns to read from the Ecom Tracker.
4. **Enrichment**: Adds `ALU`, `Launch date`, `Ecom Status`, `RRP`, `SRP`, `DISC %`, `Exclusion`, `AM exclude` columns to the SC report.
5. **Filtration** (per voucher):
   - Launch date: excludes blank, `00-Jan-1900`, `past season`, `TBC`, text-only, and future launches.
   - Ecom Status: keep only `YES`.
   - RRP: exclude ≤ 16 or = 0.
   - SRP: exclude ≤ 16 (but 0 is allowed).
   - Exclusion: must exactly match the inclusion keyword for the voucher.
   - AM Exclude: blocks based on remark type and voucher percentage.
6. Eligible SKUs receive `YES` in the corresponding voucher column.
7. Output is downloadable as `.xlsx`.

---

## AM Exclusion Remark Logic

| Remark contains | Effect |
|---|---|
| "Exclude from all Voucher / Markdown / Campaign" | Excluded from ALL vouchers |
| "Exclude from platform voucher" | Excluded from ALL vouchers |
| "Exclude from 10% VC" (and below) | Excluded from vouchers ≤ 10% |
| "Exclude 20% and above" | Excluded from vouchers ≥ 20% |
