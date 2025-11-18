import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re
from typing import List
import time
import zipfile

# --- File-dependent Constants ---
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")
except NameError:
    BASE_DIR = "."
    LOGO_PATH = "muuto_logo.png"

# Mapping file (CSV inside ZIP only)
MAPPING_CSV_ZIP = os.path.join(BASE_DIR, "mapping.csv.zip")

# Output headers
OUTPUT_HEADERS = [
    "New Item No.",
    "Old Item no.",
    "Ean No.",
    "Description",
    "Family",
    "Category",
]

OLD_COL_NAME = "Old Item no."
EAN_COL_NAME = "Ean No."

# Page configuration
st.set_page_config(
    layout="wide",
    page_title="Muuto Item Number Converter",
    page_icon="favicon.png",
)

# ---------------------------------------------------------
# Safe CSS
# ---------------------------------------------------------
st.markdown(
    '''
    <style>
        .stApp, body { background-color: #EFEEEB !important; }
        .main .block-container { background-color: #EFEEEB !important; padding-top: 2rem; }

        h1 { color: #5B4A14; font-size: 2.5em; margin-top: 0; }
        h2 { color: #333 !important; padding-bottom: 5px; margin-top: 30px; margin-bottom: 15px; border-bottom: 1px solid #CCC; }
        h3 { color: #5B4A14; font-size: 1.5em; padding-bottom: 3px; margin-top: 20px; margin-bottom: 10px; }
        h4 { color: #333 !important; font-size: 1.1em; margin-top: 15px; margin-bottom: 5px; }

        .stMarkdown, label, p, .stCaption, div[data-testid="stText"] { color: #333 !important; }

        div[data-testid="stAlert"] { background-color: #f7f6f4 !important; border: 1px solid #dcd4c3 !important; border-radius: 0.25rem !important; }
        div[data-testid="stAlert"] > div:first-child { background-color: transparent !important; }

        div[data-testid="stDownloadButton"] p { color: white !important; }

        div[data-testid="stDownloadButton"] button,
        div[data-testid="stButton"] button {
            border: 1px solid #5B4A14 !important;
            background-color: #5B4A14 !important;
            padding: 0.5rem 1rem !important;
            font-size: 1rem !important;
            line-height: 1.5 !important;
            border-radius: 0.25rem !important;
            font-weight: 600 !important;
            text-transform: uppercase !important;
        }
    </style>
    ''',
    unsafe_allow_html=True
)

# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------
def parse_pasted_ids(raw: str) -> List[str]:
    if not raw:
        return []
    tokens = re.split(r"[\\s,;]+", raw.strip())
    cleaned = [t.strip().strip('"').strip("'") for t in tokens if t.strip()]
    seen = set()
    out = []
    for t in cleaned:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out


@st.cache_data(show_spinner=False)
def read_mapping_from_csv_zip(zip_path: str) -> pd.DataFrame:
    """Load mapping.csv from mapping.csv.zip (fast + safe)."""
    if os.path.exists(zip_path):
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                if "mapping.csv" not in zf.namelist():
                    st.error("ZIP file does not contain mapping.csv")
                    return pd.DataFrame()

                with zf.open("mapping.csv") as f:
                    df = pd.read_csv(f, dtype=str, encoding="utf-8", sep=",")
                    df.columns = [c.strip() for c in df.columns]
                    for c in df.columns:
                        df[c] = df[c].astype(str).str.strip()
                    return df

        except Exception as e:
            st.error(f"Failed to read mapping.csv.zip: {e}")
            return pd.DataFrame()

    st.error("mapping.csv.zip not found in repository.")
    return pd.DataFrame()


def select_order_and_rename(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame()
    for h in OUTPUT_HEADERS:
        if h in df.columns:
            out[h] = df[h]
        else:
            out[h] = None
    return out


def to_xlsx_bytes_with_spinner(df: pd.DataFrame) -> bytes:
    with st.spinner("Generating Excel file..."):
        time.sleep(0.3)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        return buf.getvalue()


# ---------------------------------------------------------
# UI — unchanged
# ---------------------------------------------------------
left, right = st.columns([6, 1])
with left:
    st.title("Muuto Item Number Converter")
    st.markdown("---")
    st.markdown(
        """
        **This tool is the simplest way to map your legacy Item-Variants and EANs to the new Muuto Item Numbers.**

        **How It Works:**
        * **1. Paste IDs:** Enter your old item codes or EAN numbers below.
        * **2. View & Download:** Get instant results showing the new Item Number, Description, Family, and Category.
        """
    )
with right:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)

st.markdown("---")

st.header("1. Paste Item IDs")
raw_input = st.text_area(
    "Paste your Old Item Numbers or EANs here:",
    height=200
)

ids = parse_pasted_ids(raw_input)

# ---------------------------------------------------------
# Load mapping only IF IDs have been entered
# ---------------------------------------------------------
if ids:
    st.header("2. Results and Export")

    with st.spinner("Loading mapping file..."):
        mapping_df = read_mapping_from_csv_zip(MAPPING_CSV_ZIP)

        if mapping_df.empty:
            st.error("Mapping file is empty or unreadable.")
            st.stop()

        # Required columns check
        required_cols = [OLD_COL_NAME, EAN_COL_NAME]
        for col in required_cols:
            if col not in mapping_df.columns:
                st.error(
                    f"Required column '{col}' not found.\nAvailable: {list(mapping_df.columns)}"
                )
                st.stop()

    # Clean
    mapping_df[OLD_COL_NAME] = mapping_df[OLD_COL_NAME].astype(str).str.strip()
    mapping_df[EAN_COL_NAME] = mapping_df[EAN_COL_NAME].astype(str).str.strip()

    # Lookup
    mask = mapping_df[OLD_COL_NAME].isin(ids) | mapping_df[EAN_COL_NAME].isin(ids)
    matches = mapping_df.loc[mask].copy()

    ordered = select_order_and_rename(matches)

    # Feedback
    c1, c2, c3 = st.columns([1, 1, 4])
    with c1:
        st.metric("IDs Provided", len(ids))
    with c2:
        st.metric("Matches Found", len(ordered))
    with c3:
        pass

    # Results
    st.subheader("Conversion Result Table")
    st.dataframe(ordered, use_container_width=True, hide_index=True)

    # Download
    xlsx = to_xlsx_bytes_with_spinner(ordered)
    st.download_button(
        "Download Excel",
        data=xlsx,
        file_name="muuto_item_conversion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Paste your IDs to begin.")

# Footer
st.markdown("<hr><div style='text-align:center'>For support contact your Muuto representative.</div>", unsafe_allow_html=True)
