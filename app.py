import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, date
import re
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="VC Automation Dashboard",
    page_icon="🏷️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* global */
  html, body, [class*="css"] {
    background-color: #ffffff !important;
    color: #000000 !important;
    font-family: 'Segoe UI', sans-serif;
  }
  .stApp { background-color: #ffffff; }

  /* header banner */
  .dashboard-header {
    background: #000000;
    color: #ffffff;
    padding: 20px 32px;
    border-radius: 8px;
    margin-bottom: 24px;
  }
  .dashboard-header h1 { color: #ffffff; margin: 0; font-size: 1.8rem; letter-spacing: 0.5px; }
  .dashboard-header p  { color: #cccccc; margin: 4px 0 0; font-size: 0.9rem; }

  /* section cards */
  .section-card {
    background: #f8f8f8;
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 20px 24px;
    margin-bottom: 18px;
  }
  .section-title {
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    color: #555555;
    border-bottom: 2px solid #000000;
    padding-bottom: 6px;
    margin-bottom: 16px;
  }

  /* form labels */
  label, .stTextInput label, .stSelectbox label,
  .stDateInput label, .stTextArea label,
  .stFileUploader label { color: #000000 !important; font-weight: 600; }

  /* buttons */
  .stButton > button {
    background-color: #000000 !important;
    color: #ffffff !important;
    border: none;
    border-radius: 6px;
    padding: 10px 28px;
    font-weight: 600;
    font-size: 0.95rem;
    cursor: pointer;
    transition: background 0.2s;
  }
  .stButton > button:hover { background-color: #333333 !important; }

  /* download button */
  .stDownloadButton > button {
    background-color: #1a6b3c !important;
    color: #ffffff !important;
    border-radius: 6px;
    font-weight: 600;
  }
  .stDownloadButton > button:hover { background-color: #145230 !important; }

  /* inputs */
  .stTextInput input, .stTextArea textarea,
  .stDateInput input {
    background: #ffffff !important;
    color: #000000 !important;
    border: 1px solid #cccccc !important;
    border-radius: 6px !important;
  }
  .stSelectbox div[data-baseweb="select"] > div {
    background: #ffffff !important;
    color: #000000 !important;
    border: 1px solid #cccccc !important;
    border-radius: 6px !important;
  }

  /* file uploader */
  .stFileUploader > div {
    background: #f0f0f0 !important;
    border: 1.5px dashed #999999 !important;
    border-radius: 8px !important;
  }

  /* info / success / error boxes */
  .stSuccess { background: #e8f5e9 !important; color: #000000 !important; }
  .stError   { background: #ffebee !important; color: #000000 !important; }
  .stWarning { background: #fff8e1 !important; color: #000000 !important; }
  .stInfo    { background: #e3f2fd !important; color: #000000 !important; }

  /* divider */
  hr { border-color: #e0e0e0; }

  /* metric */
  [data-testid="metric-container"] {
    background: #f8f8f8;
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 12px 16px;
  }
  [data-testid="metric-container"] label { color: #555555 !important; font-size: 0.8rem; }
  [data-testid="metric-container"] [data-testid="stMetricValue"] { color: #000000 !important; }

  /* expander */
  .streamlit-expanderHeader { color: #000000 !important; font-weight: 600; }

  /* optional badge */
  .optional-badge {
    display: inline-block;
    background: #eeeeee;
    color: #555555;
    font-size: 0.7rem;
    font-weight: 600;
    padding: 1px 7px;
    border-radius: 4px;
    margin-left: 8px;
    vertical-align: middle;
    letter-spacing: 0.5px;
  }
</style>
""", unsafe_allow_html=True)

# ── helpers ──────────────────────────────────────────────────────────────────

def read_file(uploaded_file, sheet_name=None):
    if uploaded_file is None:
        return None
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(uploaded_file, dtype=str)
        else:
            if sheet_name:
                return pd.read_excel(uploaded_file, sheet_name=sheet_name, dtype=str)
            return pd.read_excel(uploaded_file, dtype=str)
    except Exception as e:
        st.error(f"Error reading {uploaded_file.name}: {e}")
        return None


def get_excel_sheets(uploaded_file):
    if uploaded_file is None:
        return []
    try:
        xf = pd.ExcelFile(uploaded_file)
        return xf.sheet_names
    except:
        return []


def read_excel_smart(uploaded_file, sheet_name=None, header_keywords=None, max_scan_rows=20):
    """
    Read an Excel sheet, auto-detecting the true header row.
    Scans up to max_scan_rows rows looking for a row that contains
    any of the header_keywords. Falls back to row 0 if none found.
    """
    if uploaded_file is None:
        return None
    if header_keywords is None:
        header_keywords = []
    keywords_lower = [k.lower() for k in header_keywords]
    try:
        uploaded_file.seek(0)
        # Read raw without headers to scan for the real header row
        raw = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name,
            header=None,
            dtype=str,
            nrows=max_scan_rows
        )
        header_row = 0
        for i, row in raw.iterrows():
            row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v) and str(v).strip()]
            if any(any(kw in cell for kw in keywords_lower) for cell in row_vals):
                header_row = i
                break
        # Now re-read with the correct header row
        uploaded_file.seek(0)
        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name,
            header=header_row,
            dtype=str
        )
        # Drop fully-empty rows and columns
        df.dropna(how="all", inplace=True)
        df.dropna(axis=1, how="all", inplace=True)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None


def parse_launch_date(val):
    """Convert various date formats to DD-MMM-YYYY. Returns (display_str, date_obj) or None for exclusions."""
    if pd.isna(val) or str(val).strip() == "":
        return None
    val = str(val).strip()

    # Strip time component if present (e.g. "13/05/2025 00:00:00")
    val = re.sub(r'\s+\d{1,2}:\d{2}(:\d{2})?(\s*(AM|PM))?$', '', val, flags=re.IGNORECASE).strip()

    # Default Excel date origin
    if val in ("0", "00-Jan-1900", "00/01/1900"):
        return None
    # Past season / TBC / text exclusions
    lower = val.lower()
    if any(x in lower for x in ["past season", "tbc", "n/a", "na"]):
        return None
    # Try common date formats
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d %b %Y", "%d-%b-%Y", "%d/%m/%y", "%m/%d/%Y"):
        try:
            dt = datetime.strptime(val, fmt)
            return dt.strftime("%d-%b-%Y"), dt.date()
        except:
            pass
    # Purely text (no digits) → exclude
    if not any(c.isdigit() for c in val):
        return None
    return val, None


def to_excel_bytes(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Working Sheet")
    return buf.getvalue()


def normalize_alu(val):
    return str(val).strip().upper() if pd.notna(val) else ""


# ── main logic ───────────────────────────────────────────────────────────────

def run_automation(config):
    marketplace   = config["marketplace"]
    region        = config["region"]
    price_tier    = config["price_tier"].strip().upper()
    cutoff_date   = config["cutoff_date"]
    voucher_rows  = config["voucher_rows"]   # list of {name, keywords}
    sc_file       = config["sc_file"]
    ecom_file     = config["ecom_file"]
    content_file  = config["content_file"]
    am_file       = config["am_file"]

    logs = []

    # ── 1. Read SC report ────────────────────────────────────────────────────
    sc_df = read_file(sc_file)
    if sc_df is None:
        return None, ["❌ Could not read SC report."]
    logs.append(f"✅ SC report loaded: {len(sc_df)} rows")

    # normalise SellerSKU / EAN column
    ean_col = next((c for c in sc_df.columns if "sellersku" in c.lower()), None)
    if ean_col is None:
        ean_col = next((c for c in sc_df.columns if "ean" in c.lower()), None)
    if ean_col is None:
        return None, ["❌ SC report must contain a SellerSKU or EAN column."]
    sc_df["__EAN__"] = sc_df[ean_col].apply(normalize_alu)

    # ── 2. Read Content file ─────────────────────────────────────────────────
    content_df = read_file(content_file)
    if content_df is None:
        return None, ["❌ Could not read Content file."]
    color_col = next((c for c in content_df.columns if "color no" in c.lower() or c.lower() == "color_no"), None)
    if color_col is None:
        # try generic "alu" column
        color_col = next((c for c in content_df.columns if "alu" in c.lower()), None)
    sku_col_content = next((c for c in content_df.columns if "sellersku" in c.lower() or "ean" in c.lower()), None)
    if color_col is None or sku_col_content is None:
        return None, [f"❌ Content file needs a 'Color No'/'ALU' column and a SellerSKU/EAN column. Found: {list(content_df.columns)}"]
    content_df["__ALU__"]  = content_df[color_col].apply(normalize_alu)
    content_df["__EAN__"]  = content_df[sku_col_content].apply(normalize_alu)
    # EAN → ALU map
    ean_to_alu = dict(zip(content_df["__EAN__"], content_df["__ALU__"]))
    logs.append(f"✅ Content file loaded: {len(content_df)} rows, {len(ean_to_alu)} EAN→ALU mappings")

    # ── 3. Read Ecom Tracker ─────────────────────────────────────────────────
    ecom_sheets = get_excel_sheets(ecom_file)
    ecom_tab = next((s for s in ecom_sheets if region.upper() in s.upper()), None)
    if ecom_tab is None:
        ecom_tab = ecom_sheets[0] if ecom_sheets else None
    if ecom_tab is None:
        return None, ["❌ Ecom Tracker: no matching sheet found."]

    ecom_df = read_excel_smart(
        ecom_file, sheet_name=ecom_tab,
        header_keywords=["style#", "style #", "alu", "rrp", "srp", "ecom status", "launch"]
    )
    if ecom_df is None:
        return None, [f"❌ Could not read Ecom Tracker tab '{ecom_tab}'."]
    logs.append(f"✅ Ecom Tracker tab '{ecom_tab}' loaded: {len(ecom_df)} rows, {len(ecom_df.columns)} columns")

    # find STYLE# column in ecom — handles "STYLE#", "STYLE #", "Style No", "ALU", etc.
    style_col = next(
        (c for c in ecom_df.columns if
         "style#" in c.lower().replace(" ", "") or
         "style #" in c.lower() or
         c.lower().strip() in ("style", "style no", "style number") or
         "alu" in c.lower()),
        None
    )
    if style_col is None:
        return None, [f"❌ Ecom Tracker needs a STYLE# / ALU column. Found: {list(ecom_df.columns)}"]
    logs.append(f"✅ ALU/STYLE# column identified as: '{style_col}'")
    ecom_df["__ALU__"] = ecom_df[style_col].apply(normalize_alu)

    # ── 4. Resolve Ecom Tracker columns using Price Tier text input ──────────
    # The price_tier value (e.g. "AX - BA") is used to find matching columns
    # directly inside the Ecom Tracker sheet. Columns whose header contains the
    # tier reference (or starts with it) are selected; otherwise we fall back to
    # keyword-based auto-detection.

    def find_tier_col(cols, tier, keyword):
        """Try to find a column whose name contains the tier reference AND keyword."""
        tier_clean = tier.replace(" ", "").lower()
        for c in cols:
            c_clean = c.replace(" ", "").lower()
            if tier_clean in c_clean and keyword in c_clean:
                return c
        # fallback: keyword only
        for c in cols:
            if keyword in c.lower():
                return c
        return None

    ecom_cols = list(ecom_df.columns)
    rrp_col_name  = find_tier_col(ecom_cols, price_tier, "rrp")
    srp_col_name  = find_tier_col(ecom_cols, price_tier, "srp")
    disc_col_name = find_tier_col(ecom_cols, price_tier, "disc")

    # Exclusion column = the column immediately after DISC % in the Ecom Tracker
    if disc_col_name and disc_col_name in ecom_cols:
        disc_idx = ecom_cols.index(disc_col_name)
        excl_col_name = ecom_cols[disc_idx + 1] if disc_idx + 1 < len(ecom_cols) else None
    else:
        # fallback: keyword search
        excl_col_name = find_tier_col(ecom_cols, price_tier, "exclusion")

    logs.append(
        f"✅ Price tier '{price_tier}' column resolution — "
        f"RRP: '{rrp_col_name}', SRP: '{srp_col_name}', "
        f"DISC: '{disc_col_name}', Exclusion (next col): '{excl_col_name}'"
    )

    # find launch date column
    launch_col = next((c for c in ecom_df.columns if "launch date" in c.lower() or "launch_date" in c.lower()), None)
    ecom_status_col = next((c for c in ecom_df.columns if "ecom status" in c.lower() or "ecom_status" in c.lower()), None)

    # build ALU → ecom data map
    ecom_map = {}
    for _, row in ecom_df.iterrows():
        alu = row["__ALU__"]
        if not alu:
            continue
        ecom_map[alu] = {
            "launch_date_raw": row.get(launch_col, "") if launch_col else "",
            "ecom_status":     row.get(ecom_status_col, "") if ecom_status_col else "",
            "rrp":             row.get(rrp_col_name, "")  if rrp_col_name  and rrp_col_name  in row.index else "",
            "srp":             row.get(srp_col_name, "")  if srp_col_name  and srp_col_name  in row.index else "",
            "disc":            row.get(disc_col_name, "") if disc_col_name and disc_col_name in row.index else "",
            "exclusion":       row.get(excl_col_name, "") if excl_col_name and excl_col_name in row.index else "",
        }
    logs.append(f"✅ Ecom map built: {len(ecom_map)} ALUs")

    # ── 5. Read AM Exclusion Sheet ───────────────────────────────────────────
    am_excl_map = {}   # {alu: [remark, ...]}
    if am_file is not None:
        am_sheets = get_excel_sheets(am_file)
        am_tab = next((s for s in am_sheets if ("sg" in s.lower() and region.upper() == "SG" and "vc" in s.lower() and "excl" in s.lower())), None)
        if am_tab is None:
            am_tab = next((s for s in am_sheets if ("my" in s.lower() and region.upper() == "MY" and "vc" in s.lower() and "excl" in s.lower())), None)
        if am_tab is None:
            am_tab = next((s for s in am_sheets if region.upper() in s.upper()), am_sheets[0] if am_sheets else None)
        if am_tab:
            am_file.seek(0)
            am_df = read_file(am_file, sheet_name=am_tab)
            if am_df is not None:
                art_col  = next((c for c in am_df.columns if c.lower() == "article" or "article" in c.lower()), None)
                rem_col  = next((c for c in am_df.columns if "exclusion type" in c.lower() or "remarks" in c.lower()), None)
                if art_col and rem_col:
                    for _, row in am_df.iterrows():
                        alu = normalize_alu(row[art_col])
                        remark = str(row[rem_col]).strip().lower() if pd.notna(row[rem_col]) else ""
                        if alu:
                            am_excl_map.setdefault(alu, []).append(remark)
                    logs.append(f"✅ AM Exclusion loaded: {len(am_excl_map)} ALUs from tab '{am_tab}'")
                else:
                    logs.append(f"⚠️  AM Exclusion sheet columns not matched (Article/Exclusion Type). Found: {list(am_df.columns)}")

    # ── 6. Enrich SC report ──────────────────────────────────────────────────
    sc_df["ALU"]         = sc_df["__EAN__"].map(ean_to_alu).fillna("")
    sc_df["Launch date"] = ""
    sc_df["Ecom Status"] = ""
    sc_df["RRP"]         = ""
    sc_df["SRP"]         = ""
    sc_df["DISC %"]      = ""
    sc_df["Exclusion"]   = ""
    sc_df["AM exclude"]  = ""

    # add voucher columns
    voucher_names = [v["name"] for v in voucher_rows]
    for vn in voucher_names:
        sc_df[vn] = ""

    # ── 7. Fill enrichment columns ───────────────────────────────────────────
    for idx, row in sc_df.iterrows():
        alu = row["ALU"]
        if not alu:
            continue
        ed = ecom_map.get(alu, {})

        # launch date
        raw_ld = str(ed.get("launch_date_raw", "")).strip()
        parsed = parse_launch_date(raw_ld)
        if parsed is None:
            ld_display = ""
            ld_date    = None
        elif isinstance(parsed, tuple):
            ld_display, ld_date = parsed
        else:
            ld_display = parsed
            ld_date    = None

        # future launch check
        if ld_date and cutoff_date and ld_date > cutoff_date:
            ld_display = f"Future Launch ({ld_display})"

        sc_df.at[idx, "Launch date"] = ld_display
        sc_df.at[idx, "Ecom Status"] = ed.get("ecom_status", "")
        sc_df.at[idx, "RRP"]         = ed.get("rrp", "")
        sc_df.at[idx, "SRP"]         = ed.get("srp", "")
        sc_df.at[idx, "DISC %"]      = ed.get("disc", "")
        sc_df.at[idx, "Exclusion"]   = ed.get("exclusion", "")

        # AM exclusion
        remarks = am_excl_map.get(alu, [])
        sc_df.at[idx, "AM exclude"] = "; ".join(remarks) if remarks else ""

    logs.append("✅ SC report enriched with Ecom Tracker data")

    # ── 8. Filtration and YES assignment ────────────────────────────────────
    def is_eligible_base(row):
        """Returns True if row passes base filters (launch date, ecom status, RRP, SRP)."""
        # Launch date
        ld = str(row.get("Launch date") or "").strip().lower()
        if ld == "" or "future launch" in ld:
            return False

        # Ecom Status — must be YES
        es = str(row.get("Ecom Status") or "").strip().upper()
        if es != "YES":
            return False

        # RRP — exclude if 0 or ≤ 16
        try:
            rrp_val = float(str(row.get("RRP") or "0").replace(",", "").strip() or 0)
        except:
            rrp_val = 0
        if rrp_val == 0 or rrp_val <= 16:
            return False

        # SRP — exclude only if ≤ 16 AND not 0 (0 is allowed)
        try:
            srp_val = float(str(row.get("SRP") or "0").replace(",", "").strip() or 0)
        except:
            srp_val = 0
        if srp_val != 0 and srp_val <= 16:
            return False

        return True

    def am_exclusion_applies(alu, remark_text, voucher_pct):
        """Check if AM exclusion remark blocks this voucher percentage."""
        remark_text = remark_text.lower()
        remarks = [r.strip() for r in remark_text.split(";")]
        for r in remarks:
            if "exclude from all" in r or "all voucher" in r or "markdown" in r or "campaign" in r:
                return True
            if "platform voucher" in r:
                return True
            # "exclude from 10% vc" → exclude 10% and below
            m10 = re.search(r"exclude from (\d+)%", r)
            if m10 and "and below" in r:
                threshold = int(m10.group(1))
                if voucher_pct is not None and voucher_pct <= threshold:
                    return True
            # "exclude 20% and above"
            m20 = re.search(r"exclude (\d+)% and above", r)
            if m20:
                threshold = int(m20.group(1))
                if voucher_pct is not None and voucher_pct >= threshold:
                    return True
        return False

    def extract_pct(voucher_name):
        m = re.search(r"(\d+)\s*%", voucher_name)
        return int(m.group(1)) if m else None

    yes_counts = {}
    # Convert to list of dicts for faster iteration
    sc_records = sc_df.to_dict("records")

    for v in voucher_rows:
        vname    = v["name"]
        keywords = [kw.strip().lower() for kw in v["keywords"].split(",") if kw.strip()]
        vpct     = extract_pct(vname)
        count    = 0

        for i, row in enumerate(sc_records):
            if not is_eligible_base(row):
                continue

            # Inclusion keyword filter — keyword must be contained within the Exclusion column value
            excl_val = str(row.get("Exclusion", "") or "").strip().lower()
            if keywords:
                if not any(kw in excl_val for kw in keywords):
                    continue

            # AM Exclusion
            am_val = str(row.get("AM exclude", "") or "").strip()
            alu    = str(row.get("ALU", "") or "").strip()
            if am_val and am_exclusion_applies(alu, am_val, vpct):
                continue

            sc_records[i][vname] = "YES"
            count += 1

        yes_counts[vname] = count
        logs.append(f"✅ Voucher '{vname}': {count} eligible SKUs marked YES")

    # Write results back to dataframe
    sc_df = pd.DataFrame(sc_records)

    # drop internal columns
    sc_df.drop(columns=["__EAN__"], inplace=True, errors="ignore")
    logs.append(f"✅ Done! Output: {len(sc_df)} rows, {len(sc_df.columns)} columns")

    return sc_df, logs, yes_counts


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="dashboard-header">
  <h1>🏷️ VC Automation Dashboard</h1>
  <p>Voucher Code Eligibility Automation · Lazada & Zalora</p>
</div>
""", unsafe_allow_html=True)

# ── Row 1 : Marketplace & Region ─────────────────────────────────────────────
st.markdown('<div class="section-card"><div class="section-title">Configuration</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    marketplace = st.selectbox("Select Marketplace", ["Lazada", "Zalora"])
with col2:
    region = st.selectbox("Region", ["SG", "MY"])
st.markdown('</div>', unsafe_allow_html=True)

# ── Row 2 : File uploads ──────────────────────────────────────────────────────
st.markdown('<div class="section-card"><div class="section-title">File Uploads</div>', unsafe_allow_html=True)

col_a, col_b, col_c = st.columns(3)
with col_a:
    sc_label = "Lazada SC Report" if marketplace == "Lazada" else "Zalora SC Report"
    sc_file  = st.file_uploader(sc_label, type=["csv","xlsx"])
with col_b:
    ecom_file = st.file_uploader("Ecom Tracker", type=["csv","xlsx"])
with col_c:
    content_file = st.file_uploader("Content File", type=["csv","xlsx"])

col_d, _ = st.columns([1, 2])
with col_d:
    st.markdown('<span style="font-weight:600;color:#000;">AM Exclusion Sheet <span class="optional-badge">OPTIONAL</span></span>', unsafe_allow_html=True)
    am_file = st.file_uploader("", type=["csv","xlsx"], key="am_excl", label_visibility="collapsed")

st.markdown('</div>', unsafe_allow_html=True)

# ── Row 3 : Parameters ───────────────────────────────────────────────────────
st.markdown('<div class="section-card"><div class="section-title">Parameters</div>', unsafe_allow_html=True)

col_p1, col_p2 = st.columns(2)
with col_p1:
    price_tier = st.text_input(
        f"Ecom Tracker Price Tier Reference ({region})",
        placeholder="e.g. AX - BA",
        help="Enter the reference range exactly as it appears in the Price Tier Reference file."
    )
with col_p2:
    cutoff_date = st.date_input(
        "Launch Date Cutoff",
        value=date.today(),
        help="Dates up to and including this date are considered. Dates after are 'Future Launch'."
    )

st.markdown('</div>', unsafe_allow_html=True)

# ── Row 4 : Vouchers ─────────────────────────────────────────────────────────
st.markdown('<div class="section-card"><div class="section-title">Voucher Names & Inclusion Keywords</div>', unsafe_allow_html=True)
st.markdown('<p style="font-size:0.85rem;color:#555;margin-bottom:12px;">Add one or more vouchers. For each voucher, provide its name and the inclusion keyword(s) — multiple keywords comma-separated (exact match per SKU).</p>', unsafe_allow_html=True)

if "voucher_count" not in st.session_state:
    st.session_state.voucher_count = 1

btn_col1, btn_col2, _ = st.columns([1, 1, 6])
with btn_col1:
    if st.button("＋ Add Voucher"):
        st.session_state.voucher_count += 1
with btn_col2:
    if st.button("－ Remove Last") and st.session_state.voucher_count > 1:
        st.session_state.voucher_count -= 1

voucher_rows = []
for i in range(st.session_state.voucher_count):
    vc1, vc2 = st.columns(2)
    with vc1:
        vname = st.text_input(f"Voucher Name #{i+1}", key=f"vname_{i}", placeholder="e.g. 10% OFF Voucher")
    with vc2:
        vkw   = st.text_input(f"Inclusion Keyword(s) #{i+1}", key=f"vkw_{i}", placeholder="e.g. beauty, health, personal care")
    if vname.strip():
        voucher_rows.append({"name": vname.strip(), "keywords": vkw.strip()})

st.markdown('</div>', unsafe_allow_html=True)

# ── Run button ────────────────────────────────────────────────────────────────
st.markdown("---")
run_col, _ = st.columns([2, 8])
with run_col:
    run_btn = st.button("▶ Run Automation", use_container_width=True)

if run_btn:
    # validation
    errors = []
    if sc_file is None:        errors.append("SC Report file is required.")
    if ecom_file is None:      errors.append("Ecom Tracker file is required.")
    if content_file is None:   errors.append("Content File is required.")
    if not price_tier.strip(): errors.append("Price Tier Reference value is required.")
    if not voucher_rows:       errors.append("At least one Voucher Name is required.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        with st.spinner("Processing… this may take a moment for large files."):
            config = {
                "marketplace":   marketplace,
                "region":        region,
                "price_tier":    price_tier,
                "cutoff_date":   cutoff_date,
                "voucher_rows":  voucher_rows,
                "sc_file":       sc_file,
                "ecom_file":     ecom_file,
                "content_file":  content_file,
                "am_file":       am_file,
            }
            result = run_automation(config)

        if result[0] is None:
            logs = result[1]
            for log in logs:
                if log.startswith("❌"):
                    st.error(log)
                else:
                    st.warning(log)
        else:
            output_df, logs, yes_counts = result

            st.markdown("### Results")
            # metrics
            mcols = st.columns(len(voucher_rows) + 1)
            with mcols[0]:
                st.metric("Total SKUs", len(output_df))
            for i, v in enumerate(voucher_rows):
                with mcols[i + 1]:
                    st.metric(f"{v['name'][:20]}…" if len(v['name']) > 20 else v['name'],
                              yes_counts.get(v['name'], 0))

            # log expander
            with st.expander("Processing Log", expanded=False):
                for log in logs:
                    if log.startswith("❌"):
                        st.error(log)
                    elif log.startswith("⚠️"):
                        st.warning(log)
                    else:
                        st.success(log)

            # preview
            st.markdown("**Preview (first 50 rows)**")
            st.dataframe(output_df.head(50), use_container_width=True, height=300)

            # download
            excel_bytes = to_excel_bytes(output_df)
            st.download_button(
                label="⬇ Download Working Sheet (.xlsx)",
                data=excel_bytes,
                file_name=f"VC_Working_Sheet_{marketplace}_{region}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=False
            )
