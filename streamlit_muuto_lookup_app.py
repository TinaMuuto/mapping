import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re
from typing import List
import zipfile

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

        /* Hvid tekst på alle knapper */
        div[data-testid="stDownloadButton"] button,
        div[data-testid="stButton"] button {
            border: 1px solid #5B4A14 !important;
            background-color: #5B4A14 !important;
            padding: 0.5rem 1rem !important;
            font-size: 1rem !important;
            border-radius: 0.25rem !important;
            font-weight: 600 !important;
            text-transform: uppercase !important;
            color: #FFFFFF !important;
        }

        div[data-testid="stDownloadButton"] p {
            color: #FFFFFF !important;
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
    - fjern ALT der ikke er A-Z eller 0-9 (inkl. mellemrum, '-', '.', '/')
    """
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


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

            # Læs hele filen med den fundne separator
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


def deterministic_lookup(ids: List[str], df: pd.DataFrame) -> pd.DataFrame:
    """
    For hvert ID:
    1) Exact match på normaliseret værdi (OLD / EAN)
    2) Hvis ingen: prefix-match begge veje (ID starter med mapping, mapping starter med ID)
    3) Hvis ingen: substring-match
    4) Hvis stadig ingen: tom række
    """

    rows = []

    for raw_id in ids:
        norm_id = normalize_id(raw_id)

        # Exact match på normaliserede kolonner
        exact = df[(df["_OLD_norm"] == norm_id) | (df["_EAN_norm"] == norm_id)]

        if not exact.empty:
            tmp = exact.copy()
            tmp["Query"] = raw_id
            tmp["Match Type"] = "Exact"
            rows.append(tmp)
            continue

        # Prefix: mapping starter med ID
        prefix_old = df[df["_OLD_norm"].str.startswith(norm_id)]
        prefix_ean = df[df["_EAN_norm"].str.startswith(norm_id)]

        if not prefix_old.empty or not prefix_ean.empty:
            tmp = pd.concat([prefix_old, prefix_ean]).drop_duplicates()
            tmp["Query"] = raw_id
            tmp["Match Type"] = "Prefix (ID shorter)"
            rows.append(tmp)
            continue

        # Prefix: ID starter med mapping (mapping er "basis", ID mere detaljeret)
        prefix_old_rev = df[df["_OLD_norm"].apply(lambda x: norm_id.startswith(x) if isinstance(x, str) else False)]
        prefix_ean_rev = df[df["_EAN_norm"].apply(lambda x: norm_id.startswith(x) if isinstance(x, str) else False)]

        if not prefix_old_rev.empty or not prefix_ean_rev.empty:
            tmp = pd.concat([prefix_old_rev, prefix_ean_rev]).drop_duplicates()
            tmp["Query"] = raw_id
            tmp["Match Type"] = "Prefix (mapping shorter)"
            rows.append(tmp)
            continue

        # Substring-match som sidste forsøg
        substr = df[
            df["_OLD_norm"].str.contains(norm_id, na=False)
            | df["_EAN_norm"].str.contains(norm_id, na=False)
        ]
        if not substr.empty:
            tmp = substr.copy()
            tmp["Query"] = raw_id
            tmp["Match Type"] = "Substring"
            rows.append(tmp)
            continue

        # Ingen match
        empty_row = {h: None for h in OUTPUT_HEADERS}
        empty_row["Query"] = raw_id
        empty_row["Match Type"] = "No match"
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
        3. See matches (Exact / Prefix / Substring)  
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

    # Én samlet spinner, så det er tydeligt at appen arbejder
    with st.spinner("Converting IDs..."):
        mapping_df = read_mapping_from_zip(MAPPING_ZIP_PATH, MAPPING_FILENAME)

        if mapping_df.empty:
            st.error("Mapping file is empty or unreadable.")
            st.stop()

        # Tjek nødvendige kolonner
        missing = [c for c in [OLD_COL_NAME, EAN_COL_NAME] if c not in mapping_df.columns]
        if missing:
            st.error(
                f"Required column(s) missing in mapping.csv: {missing}\n"
                f"Actual columns: {list(mapping_df.columns)}"
            )
            st.stop()

        # Normaliser kolonner til match
        mapping_df["_OLD_norm"] = mapping_df[OLD_COL_NAME].apply(normalize_id)
        mapping_df["_EAN_norm"] = mapping_df[EAN_COL_NAME].apply(normalize_id)

        # Sørg for at alle outputkolonner findes (Family/Category kan mangle i CSV)
        for h in OUTPUT_HEADERS:
            if h not in mapping_df.columns:
                mapping_df[h] = None

        # Deterministisk match (exact + prefix + substring)
        results = deterministic_lookup(ids, mapping_df)

        matches_count = (results["Match Type"] != "No match").sum()
        results_sorted = results.sort_values(
            by=["Match Type", "Query"],
            ascending=[True, True],
        )

        display_cols = ["Query", "Match Type"] + OUTPUT_HEADERS
        display_df = results_sorted[display_cols]

    # Når vi kommer herud, er spinneren væk
    st.header("2. Results")
    st.metric("IDs Provided", len(ids))
    st.metric("IDs with a match", matches_count)

    st.dataframe(display_df, use_container_width=True, hide_index=True)

    xlsx = to_xlsx_bytes(display_df)
    st.download_button(
        "Download Excel File",
        data=xlsx,
        file_name="muuto_item_conversion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Paste your IDs above and click **Convert IDs** to run the lookup.")
