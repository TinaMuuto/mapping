import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re
from typing import List
import time
import zipfile
from rapidfuzz import fuzz, process    # FAST fuzzy matching


# =====================================================================================
# CONFIG
# =====================================================================================
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")
except NameError:
    BASE_DIR = "."
    LOGO_PATH = "muuto_logo.png"

MAPPING_ZIP_PATH = os.path.join(BASE_DIR, "mapping.csv.zip")
MAPPING_FILENAME = "mapping.csv"

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


# =====================================================================================
# PAGE SETTINGS + CSS
# =====================================================================================
st.set_page_config(
    layout="wide",
    page_title="Muuto Item Number Converter",
    page_icon="favicon.png",
)

st.markdown(
    """
    <style>
        .stApp, body { background-color: #EFEEEB !important; }
        .main .block-container { background-color: #EFEEEB !important; padding-top: 2rem; }

        h1 { color: #5B4A14; font-size: 2.5em; margin-top: 0; }
        h2 { color: #333 !important; border-bottom: 1px solid #CCC; }

        div[data-testid="stDownloadButton"] p { color: white !important; }
        div[data-testid="stDownloadButton"] button,
        div[data-testid="stButton"] button {
            border: 1px solid #5B4A14 !important;
            background-color: #5B4A14 !important;
            padding: 0.5rem 1rem !important;
            font-size: 1rem !important;
            border-radius: 0.25rem !important;
            font-weight: 600 !important;
            text-transform: uppercase !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =====================================================================================
# HELPERS
# =====================================================================================
def parse_pasted_ids(raw: str) -> List[str]:
    if not raw:
        return []
    tokens = re.split(r"[\s,;]+", raw.strip())
    out = []
    seen = set()
    for t in tokens:
        t = t.strip().strip('"').strip("'")
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def autodetect_separator(first_10_lines: str) -> str:
    """Auto-detect separator (comma, semicolon, tab)."""
    if ";" in first_10_lines:
        return ";"
    if "," in first_10_lines:
        return ","
    if "\t" in first_10_lines:
        return "\t"
    return ","


@st.cache_data(show_spinner=False)
def read_mapping_from_zip(zip_path: str, filename: str) -> pd.DataFrame:
    """Load mapping.csv inside mapping.csv.zip with autodetected delimiter."""
    if not os.path.exists(zip_path):
        st.error("mapping.csv.zip not found in repository.")
        return pd.DataFrame()

    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            if filename not in zf.namelist():
                st.error("ZIP file does not contain mapping.csv")
                return pd.DataFrame()

            # Peek at first lines for separator detection
            with zf.open(filename) as f:
                head = f.read(5000).decode("utf-8", errors="ignore")
                sep = autodetect_separator(head)

            # Load for real
            with zf.open(filename) as f:
                df = pd.read_csv(
                    f, dtype=str, encoding="utf-8",
                    sep=sep, engine="python"
                )

            df.columns = [c.strip() for c in df.columns]
            for c in df.columns:
                df[c] = df[c].astype(str).str.strip()

            return df

    except Exception as e:
        st.error(f"Failed to read mapping.csv.zip: {e}")
        return pd.DataFrame()


def to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def fuzzy_lookup(ids: List[str], df: pd.DataFrame) -> pd.DataFrame:
    """
    For each ID:
      1) Try exact match on OLD or EAN
      2) If no match → fuzzy match (threshold 70)
    """

    rows = []

    for id_ in ids:
        id_clean = id_.strip()

        # Exact match
        exact = df[
            (df[OLD_COL_NAME] == id_clean) |
            (df[EAN_COL_NAME] == id_clean)
        ]

        if not exact.empty:
            tmp = exact.copy()
            tmp["Match Score"] = 100
            tmp["Query"] = id_
            rows.append(tmp)
            continue

        # Fuzzy match
        candidates = df[OLD_COL_NAME].dropna().unique().tolist() + \
                     df[EAN_COL_NAME].dropna().unique().tolist()

        best = process.extractOne(
            id_clean,
            candidates,
            scorer=fuzz.WRatio
        )

        if best and best[1] >= 70:
            match_value, score = best[0], best[1]
            fuzzy = df[
                (df[OLD_COL_NAME] == match_value) |
                (df[EAN_COL_NAME] == match_value)
            ].copy()
            fuzzy["Match Score"] = score
            fuzzy["Query"] = id_
            rows.append(fuzzy)
        else:
            # No match → create empty result row
            empty = pd.DataFrame({
                "Query": [id_],
                "Match Score": [0],
                **{h: None for h in OUTPUT_HEADERS}
            })
            rows.append(empty)

    return pd.concat(rows, ignore_index=True)


# =====================================================================================
# UI
# =====================================================================================
left, right = st.columns([6, 1])
with left:
    st.title("Muuto Item Number Converter")
    st.markdown("---")
    st.markdown(
        """
        **Convert legacy Item-Variants or EAN numbers to new Muuto item numbers.**

        **How It Works**
        1. Paste IDs  
        2. Click *Convert IDs*  
        3. View exact + fuzzy matches  
        4. Download results  
        """
    )
with right:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)

st.markdown("---")

st.header("1. Paste Item IDs")
raw_input = st.text_area("Paste Old Item Numbers or EANs:", height=200)
ids = parse_pasted_ids(raw_input)

submitted = st.button("Convert IDs")


# =====================================================================================
# EXECUTE ONLY WHEN USER CLICKS
# =====================================================================================
if submitted:

    if not ids:
        st.error("You must paste at least one ID before converting.")
        st.stop()

    with st.spinner("Loading mapping…"):
        df = read_mapping_from_zip(MAPPING_ZIP_PATH, MAPPING_FILENAME)

    if df.empty:
        st.error("Mapping file is empty or unreadable.")
        st.stop()

    # Validate required columns
    missing_cols = [c for c in [OLD_COL_NAME, EAN_COL_NAME] if c not in df.columns]
    if missing_cols:
        st.error(f"Missing required columns: {missing_cols}")
        st.stop()

    with st.spinner("Matching IDs… this may take a moment"):
        results = fuzzy_lookup(ids, df)

    st.header("2. Results")
    st.metric("IDs Provided", len(ids))
    st.metric("Matches Found", results["Match Score"].gt(0).sum())

    # Show ordered table
    ordered_cols = ["Query", "Match Score"] + OUTPUT_HEADERS
    display_df = results[ordered_cols]

    st.dataframe(display_df, use_container_width=True, hide_index=True)

    # Download
    xlsx = to_xlsx_bytes(display_df)
    st.download_button(
        "Download Excel File",
        data=xlsx,
        file_name="muuto_item_conversion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# Footer
st.markdown("<hr><div style='text-align:center'>For support contact your Muuto representative.</div>", unsafe_allow_html=True)
