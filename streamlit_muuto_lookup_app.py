import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re
from typing import Dict, List
import time 

# --- File-dependent Constants ---
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")
except NameError:
    LOGO_PATH = "muuto_logo.png"

# -----------------------------
# Constants
# -----------------------------
# The Google Sheet URL remains here, but is hidden from the customer UI.
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1S50it_q1BahpZCPW8dbuN7DyOMnyDgFIg76xIDSoXEk/edit?gid=1056617222#gid=1056617222"

OUTPUT_HEADERS = [
    "New Item No.",
    "OLD Item-variant",
    "Ean no.",
    "Description",
    "Family",
    "Category",
]

# -----------------------------
# Page configuration (Favicon and title)
# -----------------------------
st.set_page_config(
    layout="wide",
    page_title="Muuto Item Number Converter",
    page_icon="favicon.png", 
)

# -----------------------------
# Styling (Ensures download button text is WHITE and all main text is dark)
# -----------------------------
st.markdown(
    """
<style>
    .stApp, body { background-color: #EFEEEB !important; }
    .main .block-container { background-color: #EFEEEB !important; padding-top: 2rem; }
    
    /* Headings for a structured look and better branding */
    h1 { color: #5B4A14; font-size: 2.5em; margin-top: 0; }
    h2 { color: #333 !important; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; border-bottom: 1px solid #CCC; } 
    h3 { color: #5B4A14; font-size: 1.5em; padding-bottom: 3px; margin-top: 20px; margin-bottom: 10px; }
    h4 { color: #333 !important; font-size: 1.1em; margin-top: 15px; margin-bottom: 5px; }

    /* Ensure all general text is dark/black for readability */
    .stMarkdown, label, p, .stCaption, div[data-testid="stText"] { color: #333 !important; }
    
    /* Info/Alert boxes */
    div[data-testid="stAlert"] { background-color: #f7f6f4 !important; border: 1px solid #dcd4c3 !important; border-radius: 0.25rem !important; }
    div[data-testid="stAlert"] > div:first-child { background-color: transparent !important; }
    div[data-testid="stAlert"] div[data-testid="stMarkdownContainer"],
    div[data-testid="stAlert"] div[data-testid="stMarkdownContainer"] p { color: #31333F !important; }
    div[data-testid="stAlert"] svg { fill: #5B4A14 !important; }

    /* Inputs - Focus state */
    div[data-testid="stTextArea"] textarea:focus,
    div[data-testid="stTextInput"] input:focus,
    div[data-testid="stSelectbox"] div[data-baseweb="select"][aria-expanded="true"] > div:first-child,
    div[data-testid="stMultiSelect"] div[data-baseweb="input"]:focus-within,
    div[data-testid="stMultiSelect"] div[aria-expanded="true"] {
        border-color: #5B4A14 !important; box-shadow: 0 0 0 1px #5B4A14 !important;
    }

    /* Buttons (Muuto Brand Accent) */
    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"],
    div[data-testid="stButton"] button[data-testid^="stBaseButton"] {
        border: 1px solid #5B4A14 !important; 
        background-color: #5B4A14 !important; 
        padding: 0.5rem 1rem !important; 
        font-size: 1rem !important; 
        line-height: 1.5 !important; 
        border-radius: 0.25rem !important;
        font-weight: 600 !important; 
        text-transform: uppercase !important;
    }
    
    /* NEW: SPECIFIC CSS TO FORCE WHITE FONT COLOR ON THE DOWNLOAD BUTTON TEXT (inside the <p> tag) */
    div[data-testid="stDownloadButton"] p {
        color: #FFFFFF !important;
    }

    div[data-testid="stDownloadButton"] button[data-testid^="stBaseButton"]:hover,
    div[data-testid="stButton"] button[data-testid^="stBaseButton"]:hover {
        background-color: #4A3D10 !important; 
        color: #FFFFFF !important; 
        border-color: #4A3D10 !important;
    }
    
    .stDataFrame {
        border: 1px solid #CCC;
        border-radius: 0.25rem;
    }
</style>
""",
    unsafe_allow_html=True,
)

# -----------------------------
# Helpers
# -----------------------------
def parse_pasted_ids(raw: str) -> List[str]:
    """Extracts unique item IDs from a raw text block."""
    if not raw:
        return []
    tokens = re.split(r"[\s,;]+", raw.strip())
    cleaned = [t.strip().strip('"').strip("'") for t in tokens if t.strip()]
    seen, out = set(), []
    for t in tokens:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out


def to_csv_export_url(url: str) -> str:
    """
    Accepts a Google Sheets URL and returns a direct CSV export URL.
    """
    if not url:
        return ""
    url = url.strip()
    
    m_edit = re.search(r"https://docs.google.com/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if m_edit:
        sheet_id = m_edit.group(1)
        gid_match = re.search(r"[?&#]gid=(\d+)", url)
        gid = gid_match.group(1) if gid_match else "0"
        return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        
    m_pub = re.search(r"https://docs.google.com/spreadsheets/d/e/([a-zA-Z0-9-_]+)", url)
    if m_pub:
        doc_id_e = m_pub.group(1)
        gid_match = re.search(r"[?&#]gid=(\d+)", url)
        gid = gid_match.group(1) if gid_match else "0"
        return f"https://docs.google.com/spreadsheets/d/e/{doc_id_e}/pub?gid={gid}&single=true&output=csv"
        
    return url


@st.cache_data(show_spinner="Loading and preparing mapping database...")
def read_mapping_from_gsheets(csv_url: str) -> pd.DataFrame:
    """Loads mapping data from Google Sheets (internal function)."""
    if not csv_url:
        return pd.DataFrame()
    try:
        df = pd.read_csv(csv_url, dtype=str, keep_default_na=False, on_bad_lines='skip')
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(
            "❌ **Conversion Tool Error.** The mapping database could not be loaded. "
            "Please try refreshing the page. If the issue persists, contact Muuto support."
        )
        return pd.DataFrame()


def map_case_insensitive(df: pd.DataFrame, required: list) -> Dict[str, str]:
    """Maps required header names (case-insensitive) to actual column names."""
    lower_map = {c.lower(): c for c in df.columns}
    return {name: lower_map.get(name.lower()) for name in required}


def select_order_and_rename(df: pd.DataFrame, colmap: Dict[str, str]) -> pd.DataFrame:
    """Selects columns, ensures order, and renames them to canonical headers."""
    cols = []
    for h in OUTPUT_HEADERS:
        actual = colmap.get(h)
        if actual and actual in df.columns:
            cols.append(actual)
        else:
            df[h] = None
            cols.append(h)
            
    out = df[cols].copy()
    
    rename_map = {colmap[h]: h for h in OUTPUT_HEADERS if colmap.get(h) and colmap[h] != h}
    if rename_map:
        out = out.rename(columns=rename_map)
        
    return out


def to_xlsx_bytes_with_spinner(df: pd.DataFrame, sheet_name: str = "Conversion Output") -> bytes:
    """Converts DataFrame to Excel bytes, showing a spinner during the process."""
    with st.spinner("Generating Excel file... Please wait."):
        time.sleep(0.5) 
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        return buf.getvalue()

# -----------------------------
# Internal Setup (Hidden from UI)
# -----------------------------
csv_url = to_csv_export_url(DEFAULT_SHEET_URL)
mapping_df = read_mapping_from_gsheets(csv_url) if csv_url else pd.DataFrame()

# Check for successful internal load
if mapping_df.empty:
    st.stop() 

# Internal column validation and preparation
required = OUTPUT_HEADERS + ["OLD Item-variant", "Ean no."]
colmap = map_case_insensitive(mapping_df, required)

if not colmap.get("OLD Item-variant") or not colmap.get("Ean no."):
    st.error(
        "❌ Internal Error: Required mapping columns are missing. Please contact Muuto support."
    )
    st.stop()

old_col = colmap["OLD Item-variant"]
ean_col = colmap["Ean no."]
work = mapping_df.copy()
work[old_col] = work[old_col].astype(str).str.strip()
work[ean_col] = work[ean_col].astype(str).str.strip()


# -----------------------------
# App Main Content (Customer-facing flow)
# -----------------------------

# --- Header and Introduction ---
left, right = st.columns([6, 1])
with left:
    st.title("Muuto Item Number Converter")
    st.markdown("---")
    st.markdown(
        """
        **This tool is the simplest way to map your legacy Item-Variants and EANs to the new Muuto Item Numbers.**
        
        **How It Works:**
        * **1. Paste IDs:** Enter your old item codes below.
        * **2. View & Download:** Get instant results showing the new Item Number, Description, Family, and Category.
        """
    )
with right:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)

st.markdown("---")

# -----------------------------
# Step 1: Paste Item IDs (Input collection)
# -----------------------------
st.header("1. Paste Item IDs")

raw_input = st.text_area(
    "Paste your Old Item-Variants or EAN Numbers here.",
    height=200,
    placeholder="Enter one or more IDs, separated by new lines, commas, or spaces.\n"
                "Example:\n"
                "5710562801234\n"
                "MTO-CHAIR-001-01\n"
                "MTO-SOFA-CHAIS-LEFT-22",
)

ids = parse_pasted_ids(raw_input)


# -----------------------------
# Conditional Logic: Load data only if IDs are present
# -----------------------------
if ids:
    # --- Internal Setup (TRIGGERED HERE) ---
    # The setup logic is defined outside the main blocks, but the data is only loaded here.
    
    # -----------------------------
    # Step 2: Results and Export
    # -----------------------------
    st.header("2. Results and Export")

    # --- Lookup Logic ---
    mask = work[old_col].isin(ids) | work[ean_col].isin(ids)
    matches = work.loc[mask].copy()

    matched_keys = set(matches[old_col].dropna().astype(str)) | set(matches[ean_col].dropna().astype(str))
    not_found = [x for x in ids if x not in matched_keys]

    ordered = select_order_and_rename(matches, colmap)

    # --- Metrics and Feedback ---
    c1, c2, c3 = st.columns([1, 1, 4])
    with c1:
        st.metric("IDs Provided", len(ids))
    with c2:
        st.metric("Matches Found", len(ordered))
    with c3:
        if not_found:
            st.warning(f"⚠️ **{len(not_found)} IDs** were not matched. See the list below.")

    # --- Display Not Found ---
    if not_found:
        st.caption("The following IDs could not be found in the database (check for typos):")
        st.code("\n".join(not_found), language=None)
        st.markdown("---")
        
    if ordered.empty:
        st.error("None of the entered IDs were matched. Please check your inputs.")
        st.stop()

    # --- Result Table and Download ---
    st.subheader("Conversion Result Table")
    st.dataframe(
        ordered, 
        use_container_width=True, 
        hide_index=True,
    )
    
    # GENERATE FILE WITH SPINNER
    xlsx = to_xlsx_bytes_with_spinner(ordered, sheet_name="Muuto Item Conversion")
    
    st.download_button(
        label="Download Results as Excel File (.xlsx)",
        data=xlsx,
        file_name="muuto_item_conversion_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_button"
    )

else:
    # Message shown when app loads, before IDs are pasted
    st.info("Paste your Item IDs in Step 1 to run the lookup.")


# --- Customer Support Note ---
st.markdown("---")
st.markdown(
    """
<div style="text-align: center;">
<small>
For support, please contact your Muuto sales representative.
</small>
</div>
""",
    unsafe_allow_html=True,
)
