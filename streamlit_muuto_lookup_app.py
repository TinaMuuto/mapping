import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re
from typing import List
import time
import zipfile
from rapidfuzz import fuzz, process  # til fuzzy matching

# --- Paths og konstanter ---
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

# ---------------------------------------------------------
# Page configuration
# ---------------------------------------------------------
st.set_page_config(
    layout="wide",
    page_title="Muuto Item Number Converter",
    page_icon="favicon.png",
)

# ---------------------------------------------------------
# CSS
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# Hjælpefunktioner
# ---------------------------------------------------------
def parse_pasted_ids(raw: str) -> List[str]:
    """Split input på whitespace/komma/semikolon og returnér unikke IDs."""
    if not raw:
        return []
    tokens = re.split(r"[\s,;]+", raw.strip())
    seen = set()
    out = []
    for t in tokens:
        t = t.strip().strip('"').strip("'")
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def autodetect_separator(first_chunk: str) -> str:
    """Auto-detekter separator i CSV (semikolon, komma eller tab)."""
    if ";" in first_chunk:
        return ";"
    if "\t" in first_chunk:
        return "\t"
    if "," in first_chunk:
        return ","
    return ","


def normalize_id(s: str) -> str:
    """
    Normaliser ID til match:
    - upper case
    - fjern mellemrum og bindestreger
    """
    return re.sub(r"[\s\-]+", "", str(s)).upper()


@st.cache_data(show_spinner=False)
def read_mapping_from_zip(zip_path: str, filename: str) -> pd.DataFrame:
    """Loader mapping.csv fra mapping.csv.zip med auto-separator."""
    if not os.path.exists(zip_path):
        st.error("mapping.csv.zip not found in repository.")
        return pd.DataFrame()

    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            if filename not in zf.namelist():
                st.error(f"ZIP file does not contain {filename}")
                return pd.DataFrame()

            # Læs et lille chunk for at gætte separator
            with zf.open(filename) as f:
                head = f.read(5000).decode("utf-8", errors="ignore")
                sep = autodetect_separator(head)

            # Læs hele filen med funden separator
            with zf.open(filename) as f:
                df = pd.read_csv(
                    f,
                    dtype=str,
                    encoding="utf-8",
                    sep=sep,
                    engine="python",
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
    For hvert ID:
    1) Forsøg normaliseret exact match (case-insensitivt, ignorerer mellemrum og '-')
    2) Hvis ingen: fuzzy match mod normaliserede værdier
    3) Hvis stadig ingen: tom række med Match Score = 0
    """
    rows = []

    # Kandidatliste til fuzzy (normaliserede værdier)
    candidates = df["_OLD_norm"].dropna().unique().tolist() + df["_EAN_norm"].dropna().unique().tolist()

    for raw_id in ids:
        norm_id = normalize_id(raw_id)

        # Exact på normaliseret værdi
        exact = df[(df["_OLD_norm"] == norm_id) | (df["_EAN_norm"] == norm_id)]

        if not exact.empty:
            tmp = exact.copy()
            tmp["Query"] = raw_id
            tmp["Match Score"] = 100
            rows.append(tmp)
            continue

        # Fuzzy fallback
        best = process.extractOne(
            norm_id,
            candidates,
            scorer=fuzz.WRatio,
        )

        if best and best[1] >= 60:  # threshold kan justeres
            best_norm_value, score = best[0], best[1]
            fuzzy_rows = df[(df["_OLD_norm"] == best_norm_value) | (df["_EAN_norm"] == best_norm_value)].copy()
            fuzzy_rows["Query"] = raw_id
            fuzzy_rows["Match Score"] = score
            rows.append(fuzzy_rows)
        else:
            # Intet match
            empty_row = {h: None for h in OUTPUT_HEADERS}
            empty_row["Query"] = raw_id
            empty_row["Match Score"] = 0
            rows.append(pd.DataFrame([empty_row]))

    result = pd.concat(rows, ignore_index=True)

    # Sørg for at alle OUTPUT_HEADERS eksisterer
    for h in OUTPUT_HEADERS:
        if h not in result.columns:
            result[h] = None

    return result


# ---------------------------------------------------------
# UI
# ---------------------------------------------------------
left, right = st.columns([6, 1])
with left:
    st.title("Muuto Item Number Converter")
    st.markdown("---")
    st.markdown(
        """
        **Convert legacy Item-Variants or EANs to new Muuto item numbers.**

        **How It Works**
        1. Paste IDs  
        2. Click **Convert IDs**  
        3. See exact and fuzzy matches (with score)  
        4. Download results  
        """
    )
with right:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)

st.markdown("---")

st.header("1. Paste Item IDs")
raw_input = st.text_area("Paste Old Item Numbers or EANs here:", height=200)
ids = parse_pasted_ids(raw_input)

submitted = st.button("Convert IDs")

# ---------------------------------------------------------
# EXECUTE WHEN USER CLICKS
# ---------------------------------------------------------
if submitted:
    if not ids:
        st.error("You must paste at least one ID before converting.")
        st.stop()

    with st.spinner("Loading mapping file..."):
        mapping_df = read_mapping_from_zip(MAPPING_ZIP_PATH, MAPPING_FILENAME)

    if mapping_df.empty:
        st.error("Mapping file is empty or unreadable.")
        st.stop()

    # Tjek nødvendige kolonner
    missing = [c for c in [OLD_COL_NAME, EAN_COL_NAME] if c not in mapping_df.columns]
    if missing:
        st.error(f"Required column(s) missing in mapping.csv: {missing}\nActual columns: {list(mapping_df.columns)}")
        st.stop()

    # Normaliser kolonner til match
    mapping_df["_OLD_norm"] = mapping_df[OLD_COL_NAME].apply(normalize_id)
    mapping_df["_EAN_norm"] = mapping_df[EAN_COL_NAME].apply(normalize_id)

    # Sørg for at alle outputkolonner findes (Family/Category kan mangle i CSV)
    for h in OUTPUT_HEADERS:
        if h not in mapping_df.columns:
            mapping_df[h] = None

    with st.spinner("Matching IDs..."):
        results = fuzzy_lookup(ids, mapping_df)

    st.header("2. Results")

    matches_with_score = results["Match Score"].gt(0).sum()
    st.metric("IDs Provided", len(ids))
    st.metric("IDs with a match", matches_with_score)

    # Sortér så højeste matchscore står øverst
    results_sorted = results.sort_values(by=["Match Score"], ascending=False)

    display_cols = ["Query", "Match Score"] + OUTPUT_HEADERS
    display_df = results_sorted[display_cols]

    st.dataframe(display_df, use_container_width=True, hide_index=True)

    # Download
    xlsx = to_xlsx_bytes(display_df)
    st.download_button(
        "Download Excel File",
        data=xlsx,
        file_name="muuto_item_conversion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Paste your IDs above and click **Convert IDs** to run the lookup.")
