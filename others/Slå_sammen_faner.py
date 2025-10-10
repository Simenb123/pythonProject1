"""
Excel-browser med Streamlit
---------------------------
Start appen slik:
    streamlit run excel_browser.py

• Velg/last opp en Excel-fil (.xlsx/.xls).
• Bla mellom faner, filtrer kolonner, paginer rader.
• Last ned enkeltark eller alle ark slått sammen.
"""

from pathlib import Path
from io import BytesIO
import pandas as pd
import streamlit as st

# ---------- 1. Sidedesign ---------------------------------------------
st.set_page_config(page_title="Excel-browser", layout="wide")
st.title("📊 Excel-browser")

# ---------- 2. Velg eller last opp Excel-fil --------------------------
EXAMPLE_PATH = Path(__file__).parent / "Pakkejobber Norge 2024.xlsx"

with st.sidebar:
    st.header("1️⃣ Velg Excel-fil")
    uploaded_file = st.file_uploader(
        "Dra inn eller velg (.xlsx / .xls)",
        type=["xlsx", "xls"],
        help="Har du ingen fil? Bruk eksempel­filen under.",
    )

    if uploaded_file is None:
        if EXAMPLE_PATH.exists():
            if st.button("Bruk eksempel­fil"):
                uploaded_file = EXAMPLE_PATH.open("rb")
                st.success("Eksempel­fil valgt ✔")
        else:
            st.warning("❗ Last opp en Excel-fil for å fortsette.")
            st.stop()

# ---------- 3. Les inn arbeidsboken -----------------------------------
@st.cache_data(show_spinner=False)
def load_workbook(file) -> pd.ExcelFile:
    return pd.ExcelFile(file)

xls = load_workbook(uploaded_file)
sheet_names = xls.sheet_names

# ---------- 4. Velg ark + visnings-innstillinger ----------------------
with st.sidebar:
    st.header("2️⃣ Velg ark (fane)")
    sheet = st.selectbox("Ark:", sheet_names)

    st.header("3️⃣ Visning")
    page_size = st.number_input(
        "Rader per side",
        min_value=5, max_value=1000, value=50, step=5,
    )

# ---------- 5. Les valgt ark ------------------------------------------
@st.cache_data(show_spinner=False)
def read_sheet(file, sheet_name) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name=sheet_name)

df = read_sheet(uploaded_file, sheet)
st.subheader(f"Ark: **{sheet}**  –  {len(df):,} rader × {df.shape[1]} kolonner")

# ---------- 6. Kolonne-filter -----------------------------------------
with st.expander("🔍 Filtrer kolonner (valgfritt)"):
    cols_to_show = st.multiselect(
        "Velg kolonner som skal vises",
        options=df.columns.tolist(),
        default=list(df.columns),
    )

df_view = df[cols_to_show]  # alltid definert

# ---------- 7. Paginert visning ---------------------------------------
total_pages = (len(df_view) - 1) // page_size + 1
page = st.number_input("Side", 1, total_pages, 1)
start, stop = (page - 1) * page_size, page * page_size

st.dataframe(
    df_view.iloc[start:stop].reset_index(drop=True),
    use_container_width=True,
    hide_index=True,
)

st.caption(
    f"Viser rader {start+1:,}–{min(stop, len(df_view)):,} "
    f"av totalt {len(df_view):,} rader."
)

# ---------- 8. Nedlasting ---------------------------------------------
def df_to_xlsx_bytes(dataframe: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False)
    return buf.getvalue()

col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "📥 Last ned dette arket (.xlsx)",
        data=df_to_xlsx_bytes(df),
        file_name=f"{sheet}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with col2:
    if st.button("⬇️ Slå sammen og last ned alle ark"):
        combined = pd.concat(
            [read_sheet(uploaded_file, s).assign(Arknavn=s) for s in sheet_names],
            ignore_index=True,
        )
        st.download_button(
            "📥 Last ned kombinert (.xlsx)",
            data=df_to_xlsx_bytes(combined),
            file_name="alle_ark_kombinert.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
