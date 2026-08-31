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
import base64
import io
import os
import tempfile
from datetime import datetime

import streamlit as st

import crisdel_toll_report as ctr

CACHE_DIR = os.path.join(os.path.dirname(__file__), ".cache")
CACHED_MAPPING_PATH = os.path.join(CACHE_DIR, "last_equipment_list.xlsx")
LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "crisdel_logo.png")
os.makedirs(CACHE_DIR, exist_ok=True)

# Crisdel brand colors, pulled from the company logo
NAVY = "#163581"
LIGHT_BLUE = "#C1E1F3"
PAGE_TINT = "#F5F9FC"

st.set_page_config(
    page_title="Crisdel Toll Reconciliation",
    page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else None,
    layout="centered",
)

st.markdown(f"""
<style>
    .stApp {{ background-color: #FFFFFF; }}

    h1, h2, h3, h4 {{ font-family: 'Times New Roman', Times, serif; color: {NAVY}; }}

    .crisdel-banner {{
        background-color: {NAVY};
        border-radius: 10px;
        padding: 1.1rem 1.5rem;
        display: flex;
        align-items: center;
        gap: 1rem;
        margin-bottom: 1.25rem;
    }}
    .crisdel-banner img {{ height: 54px; border-radius: 4px; }}
    .crisdel-banner .crisdel-title {{
        font-family: 'Times New Roman', Times, serif;
        color: #FFFFFF;
        font-size: 1.5rem;
        font-weight: 700;
        margin: 0;
        line-height: 1.2;
    }}
    .crisdel-banner .crisdel-subtitle {{
        color: {LIGHT_BLUE};
        font-size: 0.85rem;
        margin: 0.15rem 0 0 0;
    }}

    div.stButton > button, div.stDownloadButton > button {{
        background-color: {NAVY};
        color: #FFFFFF;
        border: 1px solid {NAVY};
        border-radius: 6px;
        font-weight: 600;
    }}
    div.stButton > button:hover, div.stDownloadButton > button:hover {{
        background-color: {LIGHT_BLUE};
        color: {NAVY};
        border: 1px solid {NAVY};
    }}
    div.stButton > button:disabled {{
        background-color: #B9C4D9;
        color: #F0F0F0;
        border: none;
    }}

    [data-testid="stMetric"] {{
        background-color: {LIGHT_BLUE};
        border: 1px solid {NAVY};
        border-radius: 8px;
        padding: 0.85rem 0.75rem;
    }}
    [data-testid="stMetricValue"] {{ color: {NAVY}; font-family: 'Times New Roman', Times, serif; }}
    [data-testid="stMetricLabel"] {{ color: {NAVY}; }}

    [data-testid="stFileUploaderDropzone"] {{
        border: 2px dashed {NAVY};
        background-color: {PAGE_TINT};
        border-radius: 8px;
    }}

    section[data-testid="stSidebar"] {{
        background-color: {PAGE_TINT};
        border-right: 2px solid {NAVY};
    }}
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] label {{
        color: {NAVY};
    }}

    div[data-testid="stExpander"] summary {{ color: {NAVY}; font-weight: 600; }}
</style>
""", unsafe_allow_html=True)


def _logo_b64():
    if not os.path.exists(LOGO_PATH):
        return None
    with open(LOGO_PATH, "rb") as f:
        return base64.b64encode(f.read()).decode()


_logo_data = _logo_b64()
_logo_html = f'<img src="data:image/png;base64,{_logo_data}"/>' if _logo_data else ""

st.markdown(f"""
<div class="crisdel-banner">
    {_logo_html}
    <div>
        <p class="crisdel-title">Crisdel Toll Reconciliation Dashboard</p>
        <p class="crisdel-subtitle">E-ZPass + SunPass, cross-referenced to the Crisdel fleet</p>
    </div>
</div>
""", unsafe_allow_html=True)

st.caption(
    "Upload this month's E-ZPass and SunPass transaction PDFs. "
    "Upload the equipment/transponder mapping spreadsheet too -- only "
    "needed again when the fleet list changes."
)

with st.sidebar:
    st.markdown("### Fraud flag threshold")
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

tab_generate, tab_history = st.tabs(["Generate Report", "Report History"])

with tab_generate:
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
                ctr.save_report_to_archive(wb, df, fraud_threshold=threshold)

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)

        status.empty()
        st.success("Report generated and saved to Report History.")

        st.markdown("#### Summary")
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

with tab_history:
    reports = ctr.list_archived_reports()
    if not reports:
        st.info("No reports generated yet. Once you generate one on the other tab, it'll show up here.")
    else:
        st.caption(f"{len(reports)} report(s) on file, newest first.")
        for entry in reports:
            with st.container(border=True):
                c1, c2 = st.columns([3, 1])
                with c1:
                    st.markdown(f"**{entry['label']}**")
                    st.caption(
                        f"{entry['toll_transactions']:,} transactions  |  "
                        f"${entry['total_spend']:,.2f} total  |  "
                        f"{entry['unmatched']:,} unmatched  |  "
                        f"{entry['flagged']:,} flagged (>${entry.get('fraud_threshold', 10):.0f})"
                    )
                with c2:
                    fpath = ctr.get_archived_report_path(entry['filename'])
                    if fpath:
                        with open(fpath, "rb") as f:
                            st.download_button(
                                "Download", data=f.read(), file_name=entry['filename'],
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_{entry['filename']}", use_container_width=True,
                            )
                    else:
                        st.caption("File no longer available")
