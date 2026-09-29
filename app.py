from __future__ import annotations

from datetime import datetime
import pandas as pd
import streamlit as st

from mapping_layer import canonicalize

from engine import (
    STATUS_COL,
    appraisal_value_checker,
    deed_audit_checker,
    export_results_xlsx,
    generate_statistics,
    prepare_data,
    quality_checker,
    ratio_checker,
    read_uploaded_workbook,
    sale_date_checker,
    use_code_checker,
)

st.set_page_config(page_title="Sales Ratio Quality Checker", page_icon="📊", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1.4rem; padding-bottom: 3rem;}
.tool-card {border:1px solid #d9e2ec;border-radius:14px;padding:14px 16px;background:#f8fbfe;margin-bottom:8px;}
.small-note {color:#64748b;font-size:.9rem;}
</style>
""", unsafe_allow_html=True)

st.title("📊 Sales Ratio Quality Checker")
st.caption("Browser-based version of the Excel/VBA Sales Ratio Quality Checker workflow")

for key, default in {
    "df": None, "original_df": None, "raw_df": None, "source_name": None, "last_summary": {},
    "extra_tables": {}, "last_tool": None, "prepared": False,
    "mapping_report": None, "vendor_detection": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

with st.sidebar:
    st.header("1 · Load data")
    upload = st.file_uploader("Upload a CAMA / sales-ratio workbook", type=["xlsx", "xlsm", "xls", "xlsb"])
    if upload is not None:
        signature = (upload.name, upload.size)
        if st.session_state.get("upload_signature") != signature:
            try:
                raw_df, sheet = read_uploaded_workbook(upload.getvalue(), upload.name)
                df, mapping_report, vendor_detection = canonicalize(raw_df)
                st.session_state.raw_df = raw_df.copy()
                st.session_state.df = df
                st.session_state.original_df = df.copy()
                st.session_state.mapping_report = mapping_report
                st.session_state.vendor_detection = vendor_detection
                st.session_state.source_name = upload.name
                st.session_state.upload_signature = signature
                st.session_state.last_summary = {
                    "source_sheet": sheet,
                    "detected_vendor": vendor_detection.get("vendor", "Unknown"),
                    "vendor_confidence": vendor_detection.get("confidence", 0),
                    "rows_loaded": len(df),
                    "columns_loaded": len(df.columns),
                }
                st.session_state.extra_tables = {}
                st.session_state.prepared = False
                vendor_name = vendor_detection.get("vendor", "Unknown")
                confidence = vendor_detection.get("confidence", 0)
                st.success(f"Loaded {len(df):,} rows from {sheet}. Detected vendor: {vendor_name} ({confidence:.0%} profile match).")
            except Exception as e:
                st.error(f"Could not open the workbook: {e}")

    st.divider()
    st.header("Workspace")
    if st.button("🧹 Clear All", use_container_width=True):
        for k in ["df", "original_df", "raw_df", "source_name", "last_summary", "extra_tables", "last_tool", "prepared", "upload_signature", "mapping_report", "vendor_detection"]:
            if k in {"df", "original_df", "raw_df", "source_name", "last_tool", "upload_signature", "mapping_report", "vendor_detection"}:
                st.session_state[k] = None
            elif k in {"last_summary", "extra_tables"}:
                st.session_state[k] = {}
            else:
                st.session_state[k] = False
        st.rerun()
    if st.session_state.df is not None and st.button("↩️ Restore Uploaded Data", use_container_width=True):
        st.session_state.df = st.session_state.original_df.copy()
        st.session_state.last_summary = {"status": "Restored original upload"}
        st.session_state.extra_tables = {}
        st.session_state.prepared = False
        st.rerun()

if st.session_state.df is None:
    st.info("Upload an Excel workbook to begin. The app looks for an **Analysis** sheet first and uses the first worksheet when Analysis is not present.")
    st.markdown("### Online workflow")
    st.write("Upload → Prepare Data → choose a checker → review flagged records and dashboard → download processed Excel → run another checker or Clear All.")
    st.stop()

# Header/status area
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Rows", f"{len(st.session_state.df):,}")
c2.metric("Columns", f"{len(st.session_state.df.columns):,}")
vd = st.session_state.vendor_detection or {}
c3.metric("Vendor", vd.get("vendor", "Unknown"))
c4.metric("Prepared", "Yes" if st.session_state.prepared else "No")
c5.metric("Last Tool", st.session_state.last_tool or "—")

with st.expander("🔀 Vendor Mapping Report", expanded=False):
    st.caption("The application detects the vendor format first, maps recognized source fields to canonical Analysis fields, then sends that standardized data to the checkers.")
    if st.session_state.mapping_report is not None:
        st.dataframe(st.session_state.mapping_report, use_container_width=True, hide_index=True)
    else:
        st.caption("No mapping report is available.")

st.subheader("2 · Tool Menu")
st.caption("Run tools in any order. Use **Prepare Data** between diagnostics when you want to strip prior generated status columns and start the next check cleanly.")

# 3x3 icon menu matching the workbook's conceptual tools.
r1 = st.columns(3)
r2 = st.columns(3)
r3 = st.columns(3)

with r1[0]: prep = st.button("🧰 Prepare Data", use_container_width=True, help="Clean and normalize the uploaded Analysis data")
with r1[1]: date_btn = st.button("📅 Sale Date Checker", use_container_width=True)
with r1[2]: use_btn = st.button("🏷️ Use Code Checker", use_container_width=True)
with r2[0]: value_btn = st.button("🏠 Appraisal Value Checker", use_container_width=True)
with r2[1]: deed_btn = st.button("📜 Deed / MH / Comment Audit", use_container_width=True)
with r2[2]: ratio_btn = st.button("📐 Ratio Checker", use_container_width=True)
with r3[0]: quality_btn = st.button("🛡️ Quality Checker", use_container_width=True)
with r3[1]: stats_btn = st.button("📊 Generate Statistics", use_container_width=True)
with r3[2]: reset_results = st.button("🧽 Clear Results", use_container_width=True, help="Remove the current checker output but keep uploaded data")

# Sale Date needs a tax year input.
with st.expander("Sale Date Checker setting", expanded=False):
    default_year = datetime.now().year
    tax_year = st.number_input("Tax Year of Study", min_value=1900, max_value=2200, value=default_year, step=1)
    st.caption(f"Tax Year {int(tax_year)} checks October 1, {int(tax_year)-2} through September 30, {int(tax_year)-1}.")


def run_tool(name, fn, *args):
    try:
        with st.spinner(f"Running {name}..."):
            result = fn(st.session_state.df.copy(), *args)
            if len(result) == 2:
                new_df, summary = result
                extra = {}
            else:
                new_df, extra, summary = result
            st.session_state.df = new_df
            st.session_state.last_summary = summary
            st.session_state.extra_tables = extra
            st.session_state.last_tool = name
        st.success(f"{name} complete.")
    except Exception as e:
        st.error(f"{name} stopped: {e}")

if prep:
    run_tool("Prepare Data", prepare_data)
    st.session_state.prepared = True
if date_btn: run_tool("Sale Date Checker", sale_date_checker, int(tax_year))
if use_btn: run_tool("Use Code Checker", use_code_checker)
if value_btn: run_tool("Appraisal Value Checker", appraisal_value_checker)
if deed_btn: run_tool("Deed / MH / Comment Audit", deed_audit_checker)
if ratio_btn: run_tool("Ratio Checker", ratio_checker)
if quality_btn: run_tool("Quality Checker", quality_checker)
if stats_btn: run_tool("Generate Statistics", generate_statistics)
if reset_results:
    # Reset to prepared/original form without forcing a fresh upload.
    base, _ = prepare_data(st.session_state.df)
    st.session_state.df = base
    st.session_state.last_summary = {"status": "Prior generated results cleared"}
    st.session_state.extra_tables = {}
    st.session_state.last_tool = None
    st.rerun()

st.divider()
st.subheader("3 · Dashboard")
summary = st.session_state.last_summary or {}
if summary:
    pairs = list(summary.items())
    for start in range(0, len(pairs), 4):
        cols = st.columns(4)
        for col, (k, v) in zip(cols, pairs[start:start+4]):
            if isinstance(v, float):
                if "cod" in k.lower(): display = f"{v:.2%}"
                elif "ratio" in k.lower() or "median" in k.lower() or "mean" in k.lower(): display = f"{v:.4f}"
                else: display = f"{v:,.2f}"
            else: display = str(v)
            col.metric(k.replace("_", " ").title(), display)
else:
    st.caption("Run a checker to populate the dashboard.")

for name, table in (st.session_state.extra_tables or {}).items():
    with st.expander(name, expanded=True):
        st.dataframe(table, use_container_width=True, hide_index=True)

st.subheader("4 · Analysis Results")
df_view = st.session_state.df
show_only_issues = st.toggle("Show only rows requiring review", value=False)
if show_only_issues and STATUS_COL in df_view.columns:
    mask = ~df_view[STATUS_COL].astype(str).str.contains("Verified|Perfect", case=False, regex=True)
    df_view = df_view.loc[mask]
st.dataframe(df_view, use_container_width=True, hide_index=True, height=560)

st.subheader("5 · Export")
try:
    xlsx_bytes = export_results_xlsx(st.session_state.df, st.session_state.last_summary or {}, st.session_state.extra_tables or {})
    stem = (st.session_state.source_name or "sales_ratio").rsplit(".", 1)[0]
    st.download_button(
        "⬇️ Download Processed Workbook",
        data=xlsx_bytes,
        file_name=f"{stem}_web_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
except Exception as e:
    st.warning(f"Excel export is not available for the current results: {e}")

st.caption("The web version performs the calculations in Python; it does not execute VBA or require Microsoft Excel on the server.")
