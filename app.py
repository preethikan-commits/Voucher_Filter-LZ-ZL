import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
from datetime import datetime, date
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="VC Automation Dashboard",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── CSS: Black on White, clean ──────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

  html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    color: #000 !important;
    background-color: #fff !important;
  }

  .main { background: #fff; }
  .block-container { padding: 2rem 2.5rem 3rem; }

  /* Header banner */
  .dashboard-header {
    background: #000;
    color: #fff !important;
    padding: 1.2rem 2rem;
    border-radius: 8px;
    margin-bottom: 1.8rem;
    display: flex;
    align-items: center;
    gap: 1rem;
  }
  .dashboard-header h1 { color:#fff !important; margin:0; font-size:1.5rem; font-weight:700; }
  .dashboard-header p  { color:#ccc !important; margin:0; font-size:0.85rem; }

  /* Section cards */
  .section-card {
    border: 1.5px solid #e0e0e0;
    border-radius: 8px;
    padding: 1.2rem 1.5rem;
    margin-bottom: 1.2rem;
    background: #fafafa;
  }
  .section-title {
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #555 !important;
    margin-bottom: 0.9rem;
    border-bottom: 1px solid #ddd;
    padding-bottom: 0.4rem;
  }

  /* Labels */
  label, .stSelectbox label, .stTextInput label, .stDateInput label,
  .stFileUploader label, .stTextArea label {
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    color: #000 !important;
  }

  /* Inputs */
  input, textarea, select,
  .stTextInput input, .stTextArea textarea {
    color: #000 !important;
    background: #fff !important;
    border: 1.5px solid #ccc !important;
    border-radius: 6px !important;
  }

  /* Buttons */
  .stButton > button {
    background: #000 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    padding: 0.55rem 1.6rem !important;
    font-size: 0.9rem !important;
    cursor: pointer;
  }
  .stButton > button:hover { background: #333 !important; }

  /* Download button */
  .stDownloadButton > button {
    background: #fff !important;
    color: #000 !important;
    border: 2px solid #000 !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    padding: 0.55rem 1.6rem !important;
  }
  .stDownloadButton > button:hover { background: #f0f0f0 !important; }

  /* Voucher rows */
  .voucher-row {
    background: #fff;
    border: 1px solid #ddd;
    border-radius: 6px;
    padding: 0.7rem 1rem;
    margin-bottom: 0.6rem;
  }

  /* Status boxes */
  .status-success {
    background:#e8f5e9; color:#1b5e20 !important;
    border:1px solid #a5d6a7; border-radius:6px; padding:0.8rem 1rem;
    font-weight:600; margin-top:0.5rem;
  }
  .status-error {
    background:#fce4ec; color:#b71c1c !important;
    border:1px solid #f48fb1; border-radius:6px; padding:0.8rem 1rem;
    font-weight:600; margin-top:0.5rem;
  }
  .status-info {
    background:#e3f2fd; color:#0d47a1 !important;
    border:1px solid #90caf9; border-radius:6px; padding:0.8rem 1rem;
    font-weight:600; margin-top:0.5rem;
  }

  /* File uploader */
  .uploadedFile { color:#000 !important; }
  [data-testid="stFileUploader"] { border:1.5px dashed #bbb; border-radius:8px; padding:0.5rem; }

  /* Divider */
  hr { border-color: #e0e0e0 !important; }

  /* Selectbox */
  .stSelectbox [data-baseweb="select"] { background:#fff !important; }

  /* Expander */
  .streamlit-expanderHeader { font-weight:600 !important; color:#000 !important; }

  /* Hide Streamlit branding */
  #MainMenu {visibility:hidden;}
  footer {visibility:hidden;}
</style>
""", unsafe_allow_html=True)

# ── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="dashboard-header">
  <div>
    <h1>🛒 VC Automation Dashboard</h1>
    <p>Lazada &amp; Zalora Voucher Code Eligibility Automation</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────
def read_file(uploaded):
    if uploaded is None:
        return None
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded)
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded)
    return None

def read_excel_sheets(uploaded):
    if uploaded is None:
        return None
    name = uploaded.name.lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded, sheet_name=None)
    elif name.endswith(".csv"):
        df = pd.read_csv(uploaded)
        return {"Sheet1": df}
    return None

def find_col(df, candidates):
    """Case-insensitive column finder."""
    for c in candidates:
        for col in df.columns:
            if col.strip().lower() == c.lower():
                return col
    return None

def parse_launch_date(val):
    if pd.isna(val) or val == "" or val is None:
        return None
    s = str(val).strip()
    if s in ("0", "0.0", "00-Jan-1900", "00/01/1900"):
        return None
    if s.lower() in ("tbc", "past season", "future", "n/a", "na", "-", ""):
        return None
    # Try numeric Excel serial
    try:
        serial = float(s)
        if serial < 2:
            return None
        base = datetime(1899, 12, 30)
        return (base + pd.Timedelta(days=int(serial))).date()
    except:
        pass
    # Try to detect text content (not a date)
    if re.search(r'[a-zA-Z]', s) and not re.search(r'\d', s):
        return None
    # Try various formats
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%m/%d/%Y", "%d %b %Y", "%d-%b-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except:
            pass
    return None

def format_launch_date(val):
    d = parse_launch_date(val)
    if d is None:
        return str(val).strip() if val not in (None, np.nan, "") else ""
    return d.strftime("%d-%m-%Y")

def is_excluded_launch(val):
    if pd.isna(val) or val == "" or val is None:
        return True
    s = str(val).strip().lower()
    if s in ("0", "0.0"):
        return True
    if s in ("past season", "tbc"):
        return True
    if "past season" in s or "tbc" in s:
        return True
    # Excel date serial 0/1
    try:
        if float(s) <= 1:
            return True
    except:
        pass
    # If purely text (no digits)
    if re.search(r'[a-zA-Z]', str(val)) and not re.search(r'\d', str(val)):
        return True
    return False

# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR / CONFIG
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("### Configuration")
    st.caption("Set global options here.")

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN LAYOUT
# ══════════════════════════════════════════════════════════════════════════════

col_left, col_right = st.columns([1, 1], gap="large")

# ── LEFT COLUMN ───────────────────────────────────────────────────────────────
with col_left:

    # ── Marketplace & Region ─────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Platform Selection</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        marketplace = st.selectbox("Select Marketplace", ["Lazada", "Zalora"])
    with c2:
        region = st.selectbox("Region", ["SG", "MY"])
    st.markdown('</div>', unsafe_allow_html=True)

    # ── File Uploads ─────────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">File Uploads</div>', unsafe_allow_html=True)

    if marketplace == "Lazada":
        sc_file = st.file_uploader("Lazada SC Report", type=["csv", "xlsx"])
    else:
        sc_file = st.file_uploader("Zalora SC Report", type=["csv", "xlsx"])

    ecom_file    = st.file_uploader("Ecom Tracker", type=["csv", "xlsx"])
    content_file = st.file_uploader("Content File", type=["csv", "xlsx"])
    am_file      = st.file_uploader("AM Exclusion Sheet (Optional)", type=["csv", "xlsx"])
    st.markdown('</div>', unsafe_allow_html=True)

# ── RIGHT COLUMN ──────────────────────────────────────────────────────────────
with col_right:

    # ── Price Tier & Launch Cutoff ────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Reference & Cutoff</div>', unsafe_allow_html=True)

    price_tier_ref = st.text_input(
        f"Ecom Tracker Price Tier Reference ({region})",
        placeholder="e.g. AX - BA",
        help="Refers to the 'Reference' column in the Ecom Tracker. Depends on Region."
    )

    ecom_ref_file = st.file_uploader("Ecom Tracker Price Tier Reference File", type=["csv","xlsx"])

    launch_cutoff = st.date_input(
        "Launch Date Cutoff",
        value=date.today(),
        help="Dates up to this date are eligible. Dates after this are Future Launch."
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Voucher Names & Inclusion Keywords ───────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Vouchers & Keywords</div>', unsafe_allow_html=True)

    if "voucher_rows" not in st.session_state:
        st.session_state.voucher_rows = [{"name": "", "keywords": ""}]

    def add_voucher():
        st.session_state.voucher_rows.append({"name": "", "keywords": ""})

    def remove_voucher(i):
        st.session_state.voucher_rows.pop(i)

    for idx, row in enumerate(st.session_state.voucher_rows):
        st.markdown(f'<div class="voucher-row">', unsafe_allow_html=True)
        vc1, vc2, vc3 = st.columns([2, 3, 0.5])
        with vc1:
            st.session_state.voucher_rows[idx]["name"] = st.text_input(
                f"Voucher Name #{idx+1}", value=row["name"],
                key=f"vname_{idx}", placeholder="e.g. 10% VC"
            )
        with vc2:
            st.session_state.voucher_rows[idx]["keywords"] = st.text_input(
                f"Inclusion Keywords #{idx+1}", value=row["keywords"],
                key=f"vkw_{idx}", placeholder="e.g. keyword1, keyword2"
            )
        with vc3:
            st.markdown("<br>", unsafe_allow_html=True)
            if idx > 0 and st.button("✕", key=f"del_{idx}"):
                remove_voucher(idx)
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    if st.button("＋ Add Voucher"):
        add_voucher()
        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  RUN AUTOMATION
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
run_col, _ = st.columns([1, 3])
with run_col:
    run_btn = st.button("▶  Run Automation", use_container_width=True)

if run_btn:

    errors = []
    if sc_file is None:    errors.append(f"{'Lazada' if marketplace=='Lazada' else 'Zalora'} SC Report")
    if ecom_file is None:  errors.append("Ecom Tracker")
    if content_file is None: errors.append("Content File")
    if not price_tier_ref.strip(): errors.append("Price Tier Reference")

    if errors:
        st.markdown(f'<div class="status-error">⚠️ Missing required inputs: {", ".join(errors)}</div>', unsafe_allow_html=True)
        st.stop()

    voucher_list = [
        r for r in st.session_state.voucher_rows
        if r["name"].strip()
    ]
    if not voucher_list:
        st.markdown('<div class="status-error">⚠️ Please add at least one Voucher Name.</div>', unsafe_allow_html=True)
        st.stop()

    with st.spinner("Processing… please wait."):

        # ── 1. Load files ─────────────────────────────────────────────────────
        try:
            sc_df = read_file(sc_file)
            if sc_df is None: raise ValueError("Could not read SC Report")
        except Exception as e:
            st.markdown(f'<div class="status-error">❌ SC Report read error: {e}</div>', unsafe_allow_html=True)
            st.stop()

        try:
            ecom_sheets = read_excel_sheets(ecom_file)
            if ecom_sheets is None: raise ValueError("Could not read Ecom Tracker")
            # Pick the tab matching region
            ecom_tab = None
            for tab_name, tab_df in ecom_sheets.items():
                if region.upper() in tab_name.upper():
                    ecom_tab = tab_df
                    break
            if ecom_tab is None:
                # fallback: first sheet
                ecom_tab = list(ecom_sheets.values())[0]
            ecom_df = ecom_tab.copy()
        except Exception as e:
            st.markdown(f'<div class="status-error">❌ Ecom Tracker read error: {e}</div>', unsafe_allow_html=True)
            st.stop()

        try:
            content_df = read_file(content_file)
            if content_df is None: raise ValueError("Could not read Content File")
        except Exception as e:
            st.markdown(f'<div class="status-error">❌ Content File read error: {e}</div>', unsafe_allow_html=True)
            st.stop()

        # AM Exclusion (optional)
        am_df = None
        am_tab_df = None
        if am_file is not None:
            try:
                am_sheets = read_excel_sheets(am_file)
                if am_sheets:
                    tab_key = "SG VC exclusions" if region == "SG" else "MY VC exclusions"
                    # fuzzy match
                    matched = None
                    for k in am_sheets.keys():
                        if region.upper() in k.upper():
                            matched = k; break
                    if matched:
                        am_tab_df = am_sheets[matched].copy()
                    else:
                        am_tab_df = list(am_sheets.values())[0].copy()
            except Exception as e:
                st.warning(f"AM Exclusion sheet could not be read: {e}. Proceeding without it.")

        # ── 2. Build ALU → SellerSKU mapping from Content File ──────────────
        # Content file: "Color No" = ALU
        alu_col_content = find_col(content_df, ["Color No", "ColorNo", "ALU", "Color_No"])
        if alu_col_content is None:
            # fallback: first col
            alu_col_content = content_df.columns[0]

        # Build a set of valid ALUs
        valid_alus = set(content_df[alu_col_content].dropna().astype(str).str.strip().str.upper())

        # ── 3. Map SellerSKU → ALU in SC Report ──────────────────────────────
        # SellerSKU (EAN) in SC report
        seller_sku_col = find_col(sc_df, ["SellerSKU", "Seller SKU", "SKU", "EAN", "Barcode"])
        if seller_sku_col is None:
            seller_sku_col = sc_df.columns[0]
            st.warning(f"Could not find SellerSKU column; using '{seller_sku_col}'")

        # ShopSKU col in SC
        shop_sku_col = find_col(sc_df, ["ShopSKU", "Shop SKU", "ItemID", "Item ID"])

        # From Ecom tracker: STYLE# = ALU, find SKU/EAN col
        ecom_style_col  = find_col(ecom_df, ["STYLE#", "Style#", "Style No", "StyleNo", "ALU", "Style"])
        ecom_sku_col    = find_col(ecom_df, ["SellerSKU", "EAN", "Barcode", "SKU", "Seller SKU"])

        # Price Tier filter: find 'Reference' column in ecom_df
        ecom_ref_col = find_col(ecom_df, ["Reference", "Ref", "Price Tier", "Tier"])

        # Build SellerSKU → ALU from Ecom Tracker
        sku_to_alu = {}
        if ecom_style_col and ecom_sku_col:
            for _, row in ecom_df.iterrows():
                alu_val = str(row[ecom_style_col]).strip().upper() if pd.notna(row[ecom_style_col]) else ""
                sku_val = str(row[ecom_sku_col]).strip() if pd.notna(row[ecom_sku_col]) else ""
                if sku_val and alu_val:
                    sku_to_alu[sku_val] = alu_val

        # Map ALU into SC report
        sc_df["ALU"] = sc_df[seller_sku_col].astype(str).str.strip().map(sku_to_alu).fillna("")

        # ── 4. Build ALU-level lookup from Ecom Tracker (price-tier filtered) ─
        tier_ref = price_tier_ref.strip().upper()

        # Filter ecom by price tier reference
        ecom_filtered = ecom_df.copy()
        if ecom_ref_col and tier_ref:
            ecom_filtered = ecom_df[
                ecom_df[ecom_ref_col].astype(str).str.strip().str.upper() == tier_ref
            ].copy()
            if ecom_filtered.empty:
                st.warning(f"Price Tier Reference '{tier_ref}' returned 0 rows. Using all Ecom rows.")
                ecom_filtered = ecom_df.copy()

        # Columns to fetch
        launch_col  = find_col(ecom_filtered, ["Launch Dates", "Launch Date", "LaunchDate", "Launch"])
        status_col  = find_col(ecom_filtered, ["Ecom Status", "EcomStatus", "Status"])
        rrp_col     = find_col(ecom_filtered, ["RRP"])
        srp_col     = find_col(ecom_filtered, ["SRP"])
        disc_col    = find_col(ecom_filtered, ["DISC %", "DISC%", "Disc", "Discount"])
        excl_col    = find_col(ecom_filtered, ["Exclusion", "Exclude"])

        # Build ALU → row mapping (take first match)
        alu_lookup = {}
        if ecom_style_col:
            for _, row in ecom_filtered.iterrows():
                alu_val = str(row[ecom_style_col]).strip().upper()
                if alu_val and alu_val not in alu_lookup:
                    alu_lookup[alu_val] = row

        # ── 5. Add columns to SC Report ───────────────────────────────────────
        def get_alu_field(alu, col):
            if not alu or alu not in alu_lookup or col is None:
                return np.nan
            return alu_lookup[alu].get(col, np.nan)

        sc_df["Launch date"]  = sc_df["ALU"].apply(lambda a: format_launch_date(get_alu_field(a, launch_col)))
        sc_df["Ecom Status"]  = sc_df["ALU"].apply(lambda a: get_alu_field(a, status_col))
        sc_df["RRP"]          = sc_df["ALU"].apply(lambda a: get_alu_field(a, rrp_col))
        sc_df["SRP"]          = sc_df["ALU"].apply(lambda a: get_alu_field(a, srp_col))
        sc_df["DISC %"]       = sc_df["ALU"].apply(lambda a: get_alu_field(a, disc_col))
        sc_df["Exclusion"]    = sc_df["ALU"].apply(lambda a: get_alu_field(a, excl_col))

        # ── 6. Build AM Exclusion lookup ──────────────────────────────────────
        # article_col => ALU, exclusion_type_col => Remarks
        am_exclusions = {}  # alu -> list of remark strings
        if am_tab_df is not None:
            art_col  = find_col(am_tab_df, ["Article", "ALU", "article"])
            rem_col  = find_col(am_tab_df, ["Exclusion Type", "Remarks", "Remark", "Type"])
            if art_col:
                for _, row in am_tab_df.iterrows():
                    alu_val = str(row[art_col]).strip().upper() if pd.notna(row[art_col]) else ""
                    rem_val = str(row[rem_col]).strip() if rem_col and pd.notna(row[rem_col]) else ""
                    if alu_val:
                        am_exclusions.setdefault(alu_val, []).append(rem_val.lower())

        def get_am_exclude_summary(alu):
            if not alu or alu not in am_exclusions:
                return ""
            return "; ".join(am_exclusions[alu])

        sc_df["AM exclude"] = sc_df["ALU"].apply(get_am_exclude_summary)

        # ── 7. Eligibility filter + Voucher columns ───────────────────────────
        cutoff = launch_cutoff  # datetime.date

        def is_eligible_launch(raw_val):
            """True if launch date is valid and on/before cutoff."""
            if is_excluded_launch(raw_val):
                return False
            d = parse_launch_date(raw_val)
            if d is None:
                return False
            if d > cutoff:
                return False   # Future launch
            return True

        def is_eligible_ecom_status(val):
            if pd.isna(val):
                return False
            return str(val).strip().upper() == "YES"

        def is_eligible_rrp(val):
            try:
                v = float(val)
                return v > 16
            except:
                return False

        def is_eligible_srp(val):
            try:
                v = float(val)
                return v > 16 or v == 0
            except:
                return True   # include unknowns

        def am_excludes_voucher(alu, voucher_name):
            """Check AM exclusion rules for a given voucher name."""
            if not alu or alu not in am_exclusions:
                return False
            remarks = am_exclusions[alu]
            vname_lower = voucher_name.lower()

            for rem in remarks:
                # Rule 1: Exclude from all Voucher, Markdown, Campaign
                if "exclude from all" in rem:
                    return True
                # Rule 2: Exclude from platform voucher
                if "platform voucher" in rem:
                    return True
                # Rule 3: Exclude from 10% VC (and below)
                m3 = re.search(r'(\d+)%', rem)
                if m3 and "10%" in rem and "below" in rem:
                    # Extract % from voucher name
                    vm = re.search(r'(\d+)%', vname_lower)
                    if vm and int(vm.group(1)) <= 10:
                        return True
                # Rule 4: Exclude 20% and above
                if m3 and "20%" in rem and "above" in rem:
                    vm = re.search(r'(\d+)%', vname_lower)
                    if vm and int(vm.group(1)) >= 20:
                        return True
                # Generic percentage match
                if m3:
                    threshold = int(m3.group(1))
                    vm = re.search(r'(\d+)%', vname_lower)
                    if vm:
                        vc_pct = int(vm.group(1))
                        if "below" in rem and vc_pct <= threshold:
                            return True
                        if "above" in rem and vc_pct >= threshold:
                            return True
            return False

        for v_row in voucher_list:
            vname    = v_row["name"].strip()
            keywords_raw = v_row["keywords"].strip()
            keywords = [k.strip().lower() for k in keywords_raw.split(",") if k.strip()]

            col_vals = []
            for _, sc_row in sc_df.iterrows():
                alu = str(sc_row.get("ALU", "")).strip().upper()

                # Launch date
                raw_launch = sc_row.get("Launch date", "")
                # Reconstruct original for parsing
                orig_launch = get_alu_field(alu, launch_col)
                ok_launch = is_eligible_launch(orig_launch)

                # Ecom Status
                ok_status = is_eligible_ecom_status(sc_row.get("Ecom Status"))

                # RRP
                ok_rrp = is_eligible_rrp(sc_row.get("RRP"))

                # SRP
                ok_srp = is_eligible_srp(sc_row.get("SRP"))

                # Exclusion keyword match (strict exact match in the Exclusion column)
                excl_val = str(sc_row.get("Exclusion", "")).strip().lower()
                if keywords:
                    ok_keyword = excl_val in keywords
                else:
                    ok_keyword = True

                # AM exclusion
                ok_am = not am_excludes_voucher(alu, vname)

                eligible = ok_launch and ok_status and ok_rrp and ok_srp and ok_keyword and ok_am
                col_vals.append("YES" if eligible else "")

            sc_df[vname] = col_vals

        # ── 8. Reorder columns ────────────────────────────────────────────────
        added_cols = ["ALU", "Launch date", "Ecom Status", "RRP", "SRP", "DISC %", "Exclusion", "AM exclude"]
        voucher_cols = [v["name"].strip() for v in voucher_list if v["name"].strip()]
        original_cols = [c for c in sc_df.columns if c not in added_cols + voucher_cols]
        final_cols = original_cols + added_cols + voucher_cols
        sc_df = sc_df[[c for c in final_cols if c in sc_df.columns]]

    # ── Done ──────────────────────────────────────────────────────────────────
    yes_counts = {v: (sc_df.get(v, pd.Series()) == "YES").sum() for v in voucher_cols}
    total = len(sc_df)

    st.markdown(f'<div class="status-success">✅ Automation complete — {total} rows processed.</div>', unsafe_allow_html=True)

    # Summary
    st.markdown("### Summary")
    sum_data = {"Voucher": list(yes_counts.keys()), "Eligible SKUs": list(yes_counts.values())}
    st.dataframe(pd.DataFrame(sum_data), use_container_width=True, hide_index=True)

    # Preview
    with st.expander("Preview Output (first 100 rows)"):
        st.dataframe(sc_df.head(100), use_container_width=True)

    # Download
    st.markdown("### Download")
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sc_df.to_excel(writer, index=False, sheet_name="Working Sheet")
    output.seek(0)

    fname = f"VC_Output_{marketplace}_{region}_{date.today().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        label="⬇  Download Working Sheet (.xlsx)",
        data=output.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
