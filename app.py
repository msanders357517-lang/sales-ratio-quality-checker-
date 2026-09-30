from __future__ import annotations

from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st

from mapping_layer import canonicalize, mapping_options, DERIVED_RATIO_OPTION, DO_NOT_MAP_OPTION

from engine import (
    STATUS_COL,
    DEFAULT_RATIO_THRESHOLDS,
    appraisal_value_checker,
    deed_audit_checker,
    export_results_xlsx,
    generate_statistics,
    prepare_data,
    neighborhood_ratio_statistics,
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
.help-callout {
    border: 1px solid #cbd5e1;
    border-left: 5px solid #2563eb;
    border-radius: 10px;
    padding: 0.8rem 1rem;
    background: #f8fafc;
    margin: 0.25rem 0 0.75rem 0;
}
</style>
""", unsafe_allow_html=True)

def load_readme_instructions() -> str:
    """
    Load README.md from the deployed application directory.

    README.md is the single source of truth for the in-app user guide.
    """
    candidates = [
        Path(__file__).resolve().with_name("README.md"),
        Path.cwd() / "README.md",
    ]
    for readme_path in candidates:
        if readme_path.exists():
            try:
                return readme_path.read_text(encoding="utf-8")
            except Exception as exc:
                return f"### Instructions unavailable\n\nCould not read `README.md`: {exc}"
    return (
        "### Instructions unavailable\n\n"
        "`README.md` was not found beside `app.py`. "
        "Add `README.md` to the same GitHub repository and redeploy the app."
    )


def render_user_instructions():
    """Render the repository README directly inside the app."""
    st.caption(
        "This guide is loaded directly from README.md. "
        "Update README.md in GitHub and the in-app guide updates after deployment."
    )
    st.markdown(load_readme_instructions(), unsafe_allow_html=False)


st.title("📊 Sales Ratio Quality Checker")
st.caption("Browser-based version of the Excel/VBA Sales Ratio Quality Checker workflow")

with st.expander("📘 Instructions / User Guide — Click to Open", expanded=False):
    render_user_instructions()


for key, default in {
    "df": None, "original_df": None, "raw_df": None, "source_name": None, "last_summary": {},
    "extra_tables": {}, "last_tool": None, "prepared": False,
    "mapping_report": None, "vendor_detection": None, "mapping_overrides": {},
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

if "ratio_thresholds" not in st.session_state:
    st.session_state.ratio_thresholds = DEFAULT_RATIO_THRESHOLDS.copy()

with st.sidebar:
    st.subheader("📘 Help")
    with st.popover("Open Instructions / User Guide", use_container_width=True):
        render_user_instructions()
    st.caption("Instructions are loaded directly from README.md.")

    st.divider()
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
                st.session_state.mapping_overrides = {}
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
        for k in ["df", "original_df", "raw_df", "source_name", "last_summary", "extra_tables", "last_tool", "prepared", "upload_signature", "mapping_report", "vendor_detection", "mapping_overrides"]:
            if k in {"df", "original_df", "raw_df", "source_name", "last_tool", "upload_signature", "mapping_report", "vendor_detection"}:
                st.session_state[k] = None
            elif k in {"last_summary", "extra_tables", "mapping_overrides"}:
                st.session_state[k] = {}
            else:
                st.session_state[k] = False
        st.session_state.ratio_thresholds = DEFAULT_RATIO_THRESHOLDS.copy()
        for widget_key in [
            "thr_global_too_low", "thr_global_perfect_low", "thr_global_perfect_high",
            "thr_global_too_high", "thr_neighborhood_method",
            "thr_neighborhood_lower_percentile", "thr_neighborhood_upper_percentile",
            "thr_neighborhood_fixed_low", "thr_neighborhood_fixed_high",
        ]:
            st.session_state.pop(widget_key, None)
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
    st.write("Upload → Review / Apply Field Mapping → Prepare Data → choose a checker → review flagged records and dashboard → generate statistics if needed → download processed Excel.")
    st.caption("Need help? Click **📘 Instructions / User Guide — Click to Open** under the app title, or use **📘 Help** in the sidebar. Both load directly from README.md.")
    st.stop()

# Header/status area
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Rows", f"{len(st.session_state.df):,}")
c2.metric("Columns", f"{len(st.session_state.df.columns):,}")
vd = st.session_state.vendor_detection or {}
c3.metric("Vendor", vd.get("vendor", "Unknown"))
c4.metric("Prepared", "Yes" if st.session_state.prepared else "No")
c5.metric("Last Tool", st.session_state.last_tool or "—")

st.subheader("2 · Field Mapping Review")
st.caption("Review the proposed source field for each standardized application field. Change any selection before running the checkers. Your selection becomes authoritative for the analytical engine.")
with st.expander("▶ Review / Change Field Mapping", expanded=False):

    raw_map_df = st.session_state.raw_df
    map_report = st.session_state.mapping_report
    if raw_map_df is not None and map_report is not None:
        with st.form("field_mapping_form"):
            mapping_choices = {}
            display_rows = map_report[map_report["Canonical Header"].astype(str).str.len() > 0].copy()
            for start in range(0, len(display_rows), 2):
                cols = st.columns(2)
                for ui_col, (_, row) in zip(cols, display_rows.iloc[start:start+2].iterrows()):
                    field = str(row["System Field"])
                    canonical = str(row["Canonical Header"])
                    options = mapping_options(raw_map_df, field, row)
                    current = st.session_state.mapping_overrides.get(field)
                    if current is None:
                        detected = str(row.get("Detected Header", "") or "")
                        if detected.startswith("DERIVED:"):
                            current = DERIVED_RATIO_OPTION
                        elif detected in options:
                            current = detected
                        elif field == "sales_ratio" and str(row.get("Status", "")) == "Missing":
                            current = DERIVED_RATIO_OPTION
                        else:
                            current = DO_NOT_MAP_OPTION
                    if current not in options:
                        options.insert(0, current)
                    idx = options.index(current)
                    mapping_choices[field] = ui_col.selectbox(
                        canonical, options, index=idx, key=f"map_select_{field}",
                        help=f"Choose which uploaded source column should become the standardized {canonical} field."
                    )
            apply_mapping = st.form_submit_button("✅ Apply Field Mapping", use_container_width=True)

        if apply_mapping:
            try:
                mapped_df, new_report, new_detection = canonicalize(raw_map_df, overrides=mapping_choices)
                st.session_state.mapping_overrides = mapping_choices.copy()
                st.session_state.df = mapped_df
                st.session_state.original_df = mapped_df.copy()
                st.session_state.mapping_report = new_report
                st.session_state.vendor_detection = new_detection
                st.session_state.last_summary = {
                    "status": "Field mapping applied",
                    "detected_vendor": new_detection.get("vendor", "Unknown"),
                    "rows_loaded": len(mapped_df),
                    "columns_loaded": len(mapped_df.columns),
                }
                st.session_state.extra_tables = {}
                st.session_state.last_tool = None
                st.session_state.prepared = False
                st.success("Field mapping applied. The checker engine will now use these standardized selections.")
                st.rerun()
            except Exception as e:
                st.error(f"Could not apply the selected field mapping: {e}")

        with st.expander("🔎 Mapping Details", expanded=False):
            st.dataframe(st.session_state.mapping_report, use_container_width=True, hide_index=True)

# Ratio threshold controls. Defaults mirror the original/current code behavior.
thresholds = st.session_state.ratio_thresholds
if thresholds.get("neighborhood_method") == "fixed":
    neighborhood_summary = (
        f"Fixed {float(thresholds['neighborhood_fixed_low']):.3f}–"
        f"{float(thresholds['neighborhood_fixed_high']):.3f}"
    )
else:
    neighborhood_summary = (
        f"P{float(thresholds['neighborhood_lower_percentile']):g}–"
        f"P{float(thresholds['neighborhood_upper_percentile']):g} "
        "(current code default = Q1–Q3)"
    )

st.markdown("### ⚙️ Ratio Threshold Settings")
st.caption(
    f"Applied now: Overall Too Low < {float(thresholds['global_too_low']):.3f} · "
    f"Perfect Global {float(thresholds['global_perfect_low']):.3f}–"
    f"{float(thresholds['global_perfect_high']):.3f} · "
    f"Overall Too High > {float(thresholds['global_too_high']):.3f} · "
    f"Neighborhood {neighborhood_summary}"
)

with st.expander("⚙️ Change Overall / Global and Neighborhood Thresholds", expanded=False):
    st.info(
        "The fields are automatically pre-set to the current checker requirements. "
        "Change them only when your study requires different limits, then click Apply."
    )

    st.markdown("#### Overall / Global")
    g1, g2, g3, g4 = st.columns(4)
    global_too_low = g1.number_input(
        "Too Low below",
        min_value=0.0,
        value=float(thresholds["global_too_low"]),
        step=0.01,
        format="%.4f",
        key="thr_global_too_low",
        help="Ratios below this value are Too Low for Global.",
    )
    global_perfect_low = g2.number_input(
        "Perfect Global minimum",
        min_value=0.0,
        value=float(thresholds["global_perfect_low"]),
        step=0.01,
        format="%.4f",
        key="thr_global_perfect_low",
    )
    global_perfect_high = g3.number_input(
        "Perfect Global maximum",
        min_value=0.0,
        value=float(thresholds["global_perfect_high"]),
        step=0.01,
        format="%.4f",
        key="thr_global_perfect_high",
    )
    global_too_high = g4.number_input(
        "Too High above",
        min_value=0.0,
        value=float(thresholds["global_too_high"]),
        step=0.01,
        format="%.4f",
        key="thr_global_too_high",
        help="Ratios above this value are Too High for Global.",
    )

    st.caption(
        "Values between Too Low and Perfect, and between Perfect and Too High, "
        "are classified as Acceptable Global."
    )

    st.markdown("#### Neighborhood")
    method_options = [
        "Percentile / Quartile Limits (current code)",
        "Fixed Ratio Limits",
    ]
    current_method_index = 1 if thresholds.get("neighborhood_method") == "fixed" else 0
    method_label = st.radio(
        "Neighborhood threshold method",
        method_options,
        index=current_method_index,
        horizontal=True,
        key="thr_neighborhood_method",
    )

    lower_pct = float(thresholds["neighborhood_lower_percentile"])
    upper_pct = float(thresholds["neighborhood_upper_percentile"])
    fixed_low = float(thresholds["neighborhood_fixed_low"])
    fixed_high = float(thresholds["neighborhood_fixed_high"])

    if method_label.startswith("Percentile"):
        n1, n2 = st.columns(2)
        lower_pct = n1.number_input(
            "Too Low below neighborhood percentile",
            min_value=0.0,
            max_value=100.0,
            value=lower_pct,
            step=1.0,
            format="%.1f",
            key="thr_neighborhood_lower_percentile",
            help="Current code default is the 25th percentile (Q1).",
        )
        upper_pct = n2.number_input(
            "Too High above neighborhood percentile",
            min_value=0.0,
            max_value=100.0,
            value=upper_pct,
            step=1.0,
            format="%.1f",
            key="thr_neighborhood_upper_percentile",
            help="Current code default is the 75th percentile (Q3).",
        )
        st.caption(
            "Default behavior: a ratio below Q1 is Neighborhood Too Low; "
            "above Q3 is Neighborhood Too High; between Q1 and Q3 is Acceptable."
        )
        selected_method = "percentile"
    else:
        n1, n2 = st.columns(2)
        fixed_low = n1.number_input(
            "Neighborhood Too Low below",
            min_value=0.0,
            value=fixed_low,
            step=0.01,
            format="%.4f",
            key="thr_neighborhood_fixed_low",
        )
        fixed_high = n2.number_input(
            "Neighborhood Too High above",
            min_value=0.0,
            value=fixed_high,
            step=0.01,
            format="%.4f",
            key="thr_neighborhood_fixed_high",
        )
        st.caption(
            "Fixed Ratio Limits use the same low/high cutoffs for every neighborhood "
            "instead of calculating each neighborhood's percentiles."
        )
        selected_method = "fixed"

    apply_col, reset_col = st.columns(2)
    apply_thresholds = apply_col.button(
        "✅ Apply Threshold Settings",
        use_container_width=True,
        key="apply_ratio_thresholds",
    )
    reset_thresholds = reset_col.button(
        "↩️ Reset to Current Code Defaults",
        use_container_width=True,
        key="reset_ratio_thresholds",
    )

    if apply_thresholds:
        candidate = {
            "global_too_low": float(global_too_low),
            "global_perfect_low": float(global_perfect_low),
            "global_perfect_high": float(global_perfect_high),
            "global_too_high": float(global_too_high),
            "neighborhood_method": selected_method,
            "neighborhood_lower_percentile": float(lower_pct),
            "neighborhood_upper_percentile": float(upper_pct),
            "neighborhood_fixed_low": float(fixed_low),
            "neighborhood_fixed_high": float(fixed_high),
        }

        global_order_valid = (
            candidate["global_too_low"]
            <= candidate["global_perfect_low"]
            <= candidate["global_perfect_high"]
            <= candidate["global_too_high"]
        )
        neighborhood_valid = (
            0.0 <= candidate["neighborhood_lower_percentile"]
            < candidate["neighborhood_upper_percentile"] <= 100.0
            if selected_method == "percentile"
            else candidate["neighborhood_fixed_low"] < candidate["neighborhood_fixed_high"]
        )

        if not global_order_valid:
            st.error(
                "Overall / Global values must be ordered: "
                "Too Low ≤ Perfect Minimum ≤ Perfect Maximum ≤ Too High."
            )
        elif not neighborhood_valid:
            st.error(
                "Neighborhood settings are not valid. The lower threshold must be "
                "less than the upper threshold."
            )
        else:
            st.session_state.ratio_thresholds = candidate
            st.session_state.last_summary = {
                "status": "Ratio threshold settings updated. Rerun Ratio Checker or Quality Checker."
            }
            st.session_state.extra_tables = {}
            st.session_state.last_tool = None
            st.success("Threshold settings applied.")
            st.rerun()

    if reset_thresholds:
        st.session_state.ratio_thresholds = DEFAULT_RATIO_THRESHOLDS.copy()
        for widget_key in [
            "thr_global_too_low", "thr_global_perfect_low", "thr_global_perfect_high",
            "thr_global_too_high", "thr_neighborhood_method",
            "thr_neighborhood_lower_percentile", "thr_neighborhood_upper_percentile",
            "thr_neighborhood_fixed_low", "thr_neighborhood_fixed_high",
        ]:
            st.session_state.pop(widget_key, None)
        st.session_state.last_summary = {
            "status": "Ratio thresholds reset to current code defaults."
        }
        st.session_state.extra_tables = {}
        st.session_state.last_tool = None
        st.rerun()

st.subheader("3 · Tool Menu")
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
            if name in {"Ratio Checker", "Quality Checker"}:
                neighborhood_table = neighborhood_ratio_statistics(new_df, st.session_state.ratio_thresholds)
                if not neighborhood_table.empty:
                    extra = dict(extra)
                    extra["🏘️ Neighborhood Statistics"] = neighborhood_table
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
if ratio_btn: run_tool("Ratio Checker", ratio_checker, st.session_state.ratio_thresholds)
if quality_btn: run_tool("Quality Checker", quality_checker, st.session_state.ratio_thresholds)
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
st.subheader("4 · Dashboard")
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
    with st.expander(name, expanded=False):
        st.dataframe(table, use_container_width=True, hide_index=True)

st.subheader("5 · Analysis Results")
df_view = st.session_state.df
show_only_issues = st.toggle("Show only rows requiring review", value=False)
if show_only_issues and STATUS_COL in df_view.columns:
    mask = ~df_view[STATUS_COL].astype(str).str.contains("Verified|Perfect", case=False, regex=True)
    df_view = df_view.loc[mask]
st.dataframe(df_view, use_container_width=True, hide_index=True, height=560)

st.subheader("6 · Export")
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
