import streamlit as st
import pandas as pd
from io import BytesIO
import os
import re
from typing import Dict, List
import time
import zipfile

# --- File-dependent Constants ---
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    LOGO_PATH = os.path.join(BASE_DIR, "muuto_logo.png")
except NameError:
    BASE_DIR = "."
    LOGO_PATH = "muuto_logo.png"

# -----------------------------
# Constants
# -----------------------------
# Vi forventer enten en ren Excel eller en zip med mapping.xlsx indeni
MAPPING_XLSX = os.path.join(BASE_DIR, "mapping.xlsx")
MAPPING_ZIP = os.path.join(BASE_DIR, "mapping.xlsx.zip")

# Output-kolonner (det kunden ser og får i filen)
OUTPUT_HEADERS = [
    "New Item No.",
    "Old Item no.",
    "Ean No.",
    "Description",
    "Family",
    "Category",
]

# Navne vi bruger til lookup i mapping-filen
OLD_COL_NAME = "Old Item no."
EAN_COL_NAME = "Ean No."

# -----------------------------
# Page configuration
# -----------------------------
st.set_page_config(
    layout="wide",
    page_title="Muuto Item Number Converter",
    page_icon="favicon.png",
)

# -----------------------------
# Styling
# -----------------------------
st.markdown(
    """
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
    div[data-testid="stAlert"] div[data-testid="stMarkdownContainer"],
    div[data-testid="stAlert"] div[data-testid="stMarkdownContainer"] p { color: #31333F !important; }
    div[data-testid="stAlert"] svg { fill: #5B4A14 !important; }

    div[data-testid="stTextArea"] textarea:focus,
    div[data-testid="stTextInput"] input:focus,
    div[data-testid="stSelectbox"] div[da]()
