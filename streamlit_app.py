"""
PBIX Metadata Extractor — Streamlit App
-----------------------------------------
Upload a Power BI (.pbix) file in the browser, preview its extracted
metadata, and download a formatted Excel workbook.

Run locally:
    streamlit run streamlit_app.py

Deploy free on Streamlit Community Cloud by connecting this repo at
https://share.streamlit.io — set "Main file path" to streamlit_app.py.
"""

import io
import tempfile
from pathlib import Path

import streamlit as st

from extract_pbix_metadata import extract_all, write_excel

st.set_page_config(page_title="PBIX Metadata Extractor", page_icon="📊", layout="wide")

st.title("📊 PBIX Metadata Extractor")
st.caption(
    "Upload a Power BI Desktop file (.pbix) to extract its tables, DAX measures, "
    "calculated columns/tables, relationships, Power Query (M) code, and parameters "
    "into a downloadable Excel workbook — no Power BI Desktop needed."
)

uploaded_file = st.file_uploader("Upload a .pbix file", type=["pbix"])

if uploaded_file is not None:
    with st.spinner(f"Parsing {uploaded_file.name} ..."):
        try:
            with tempfile.NamedTemporaryFile(suffix=".pbix", delete=False) as tmp:
                tmp.write(uploaded_file.getbuffer())
                tmp_path = Path(tmp.name)

            sheets = extract_all(tmp_path)

            excel_buffer = io.BytesIO()
            write_excel(sheets, excel_buffer)
            excel_buffer.seek(0)

        except Exception as e:
            st.error(f"Couldn't parse this file: {e}")
            st.stop()
        finally:
            tmp_path.unlink(missing_ok=True)

    st.success(f"Extracted {len(sheets)} metadata sheets from **{uploaded_file.name}**")

    output_name = f"{Path(uploaded_file.name).stem}_metadata.xlsx"
    st.download_button(
        "⬇️ Download Excel workbook",
        data=excel_buffer,
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.divider()
    st.subheader("Preview")

    # Quick counts across the top
    cols = st.columns(4)
    highlight_order = ["Tables", "Measures (DAX)", "Calculated Columns", "Relationships"]
    for col, sheet_name in zip(cols, highlight_order):
        if sheet_name in sheets:
            col.metric(sheet_name, len(sheets[sheet_name]))

    tabs = st.tabs(list(sheets.keys()))
    for tab, (name, df) in zip(tabs, sheets.items()):
        with tab:
            if df.empty:
                st.info("No data of this type in the model.")
            else:
                st.dataframe(df, use_container_width=True, height=400)
else:
    st.info("👆 Upload a .pbix file to get started.")
    st.markdown(
        """
        **What gets extracted:**
        - Tables & column schema
        - DAX measures (full expressions)
        - Calculated columns & calculated tables
        - Relationships (cardinality, active/inactive, cross-filter direction)
        - Power Query (M) source code
        - Parameters, perspectives, and row-level security roles (if present)

        Your file is parsed in-memory for this session only and is not stored.
        """
    )
