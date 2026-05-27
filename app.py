import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# =========================
# PAGE CONFIG
# =========================

st.set_page_config(layout="wide")

st.title("IB Macro Event Dashboard")

# =========================
# LOAD FILE
# =========================

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

if uploaded_file:

    file = pd.ExcelFile(uploaded_file)

    sheet = st.selectbox(
        "Select Sheet",
        ["YIBc1.", "YIBc2.", "YIBc3.", "YIBc4."]
    )

    df = pd.read_excel(
        uploaded_file,
        sheet_name=sheet
    )
