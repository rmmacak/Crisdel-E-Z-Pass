"""
Crisdel Toll Reconciliation Dashboard
--------------------------------------
Run locally with:

    pip install -r requirements.txt
    streamlit run app.py

Then open the local URL Streamlit prints (usually http://localhost:8501).

Drop in the E-ZPass PDF and SunPass PDF for the month, plus the Crisdel
equipment/transponder mapping spreadsheet (only needs re-uploading when the
fleet list changes -- the app remembers the last one you gave it), and
click Generate to get the combined, cross-referenced Excel report.
"""
import io
import os
import tempfile
from datetime import datetime

import streamlit as st

import crisdel_toll_report as ctr

CACHE_DIR = os.path.join(os.path.dirname(__file__), ".cache")
CACHED_MAPPING_PATH = os.path.join(CACHE_DIR, "last_equipment_list.xlsx")
os.makedirs(CACHE_DIR, exist_ok=True)

st.set_page_config(page_title="Crisdel Toll Reconciliation", layout="centered")

st.title("Crisdel Toll Reconciliation Dashboard")
st.caption(
    "Upload this month's E-ZPass and SunPass transaction PDFs. "
    "Upload the equipment/transponder mapping spreadsheet too -- only "
    "needed again when the fleet list changes."
)

with st.sidebar:
    st.header("Fraud flag threshold")
    threshold = st.number_input(
        "Flag transactions over ($)", min_value=0.0, value=10.0, step=1.0,
        help="Any toll transaction above this amount is flagged for review."
    )
    st.divider()
    if os.path.exists(CACHED_MAPPING_PATH):
        mtime = datetime.fromtimestamp(os.path.getmtime(CACHED_MAPPING_PATH))
        st.success(f"Equipment list on file (saved {mtime.strftime('%m/%d/%Y %I:%M %p')})")
        if st.button("Clear saved equipment list"):
            os.remove(CACHED_MAPPING_PATH)
            st.rerun()
    else:
        st.warning("No equipment list saved yet -- upload one below.")

col1, col2 = st.columns(2)
with col1:
    ezpass_file = st.file_uploader("E-ZPass transaction PDF", type=["pdf"], key="ezpass")
with col2:
    sunpass_file = st.file_uploader("SunPass transaction PDF", type=["pdf"], key="sunpass")

equipment_file = st.file_uploader(
    "Crisdel equipment / transponder mapping (.xlsx) -- optional if you've uploaded it before",
    type=["xlsx"], key="equipment",
)

generate = st.button("Generate Report", type="primary", use_container_width=True,
                      disabled=not (ezpass_file and sunpass_file))

if generate:
    # Resolve which mapping file to use
    if equipment_file is not None:
        with open(CACHED_MAPPING_PATH, "wb") as f:
            f.write(equipment_file.getbuffer())
        mapping_path = CACHED_MAPPING_PATH
    elif os.path.exists(CACHED_MAPPING_PATH):
        mapping_path = CACHED_MAPPING_PATH
    else:
        st.error("No equipment/transponder mapping file available. Please upload one.")
        st.stop()

    with tempfile.TemporaryDirectory() as tmp:
        ez_path = os.path.join(tmp, "ezpass.pdf")
        sp_path = os.path.join(tmp, "sunpass.pdf")
        with open(ez_path, "wb") as f:
            f.write(ezpass_file.getbuffer())
        with open(sp_path, "wb") as f:
            f.write(sunpass_file.getbuffer())

        status = st.empty()
        progress_lines = []

        def log(msg):
            progress_lines.append(msg)
            status.info("\n\n".join(progress_lines[-4:]))

        with st.spinner("Processing..."):
            df = ctr.build_combined_dataframe(
                ez_path, sp_path, mapping_path,
                fraud_threshold=threshold, progress_cb=log,
            )
            wb = ctr.build_workbook(df, fraud_threshold=threshold)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

    status.empty()
    st.success("Report generated.")

    toll_df = df[df['Transaction Type'] == 'Toll']
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Toll transactions", f"{len(toll_df):,}")
    m2.metric("Total spend", f"${toll_df['Amount'].sum():,.2f}")
    m3.metric("Unmatched", f"{(toll_df['Match Status']=='UNMATCHED').sum():,}")
    m4.metric(f"Flagged (>${threshold:.0f})", f"{toll_df['Fraud Flag'].sum():,}")

    fname = f"Crisdel Toll Reconciliation Report - {datetime.now().strftime('%b %Y')}.xlsx"
    st.download_button(
        "Download Excel Report", data=buf, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    with st.expander("Preview: unmatched transponders/plates"):
        st.dataframe(
            df[df['Match Status'] == 'UNMATCHED'][
                ['Source', 'Transaction Date', 'Transponder/Plate (raw)', 'Agency', 'Amount']
            ],
            use_container_width=True,
        )
