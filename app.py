from __future__ import annotations

from datetime import datetime
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

APP_INSTRUCTIONS = """
## 📘 Sales Ratio Quality Checker — In-App User Guide

These instructions are built directly into the application and are the user guide for the live program. The **Help** control remains in the left sidebar for quick reminders, while the full instructions stay here at the top of the main page.

### Contents
- [Quick Start](#quick-start)
- [Upload and Field Mapping](#upload-and-field-mapping)
- [Prepare Data and Tool Menu](#prepare-data-and-tool-menu)
- [Ratio Threshold Settings](#ratio-threshold-settings)
- [Neighborhood Statistics](#neighborhood-statistics)
- [Checker Reference](#checker-reference)
- [Generate Statistics and VERIFY PASS](#generate-statistics-and-verify-pass)
- [Dashboard, Results, and Export](#dashboard-results-and-export)
- [Help and Review Tips](#help-and-review-tips)

### Quick Start
1. Upload the CAMA / sales-ratio workbook from **1 · Load data** in the left sidebar.
2. Review the detected vendor and open **2 · Field Mapping Review**.
3. Verify each proposed source column. If more than one field could qualify, choose the field that should be authoritative for the analysis.
4. Click **Apply Field Mapping**.
5. Click **Prepare Data** and confirm the top summary shows **Prepared = Yes**.
6. For ratio-based work, review **Ratio Threshold Settings** before running **Ratio Checker** or **Quality Checker**.
7. Run the desired checker from **3 · Tool Menu**.
8. Review **4 · Dashboard**, expand **🏘️ Neighborhood Statistics** when it is available, and review **5 · Analysis Results**.
9. If you run **Generate Statistics**, treat **VERIFY PASS** as an overall screening result—not as proof that every neighborhood or subgroup is reliable. Review the neighborhood/valuation-zone tables and sale counts before relying on the overall status.
10. Turn on **Show only rows requiring review** to focus on exceptions.
11. Use **6 · Export** to download the processed workbook.

### Upload and Field Mapping
The app looks for an **Analysis** worksheet first and otherwise uses the first worksheet in the uploaded Excel file. Supported upload types are `.xlsx`, `.xlsm`, `.xls`, and `.xlsb`.

**Field Mapping Review** connects vendor-specific source columns to the standardized fields used by the checker engine.

- Open **Review / Change Field Mapping** with its arrow.
- Review every proposed field before relying on checker output.
- If more than one source field is a reasonable match, select the source column you want the checker to use.
- Click **Apply Field Mapping** after making changes.
- Expand **Mapping Details** when you want to verify the detected header, standardized field, populated-row count, candidate headers, or mapping method.
- For Delta files, **USE CODE** remains the authoritative Use Code unless you explicitly choose another field.
- Four-digit Use Codes are normalized by removing one outside zero when applicable: `0100 → 100`, `1000 → 100`, `0101 → 101`, `1010 → 101`; `1001` remains `1001`.

### Prepare Data and Tool Menu
**Prepare Data** creates the clean, standardized working dataset used by the analytical tools. Run it after applying field mapping and whenever you upload a new workbook or materially change the mapping.

The Tool Menu contains **Prepare Data**, **Sale Date Checker**, **Use Code Checker**, **Appraisal Value Checker**, **Deed / MH / Comment Audit**, **Ratio Checker**, **Quality Checker**, **Generate Statistics**, and **Clear Results**.

**Clear Results** removes prior generated checker output while keeping the uploaded workbook available, so you can start another analysis without re-uploading the file.

### Ratio Threshold Settings
The ratio controls are located between **Field Mapping Review** and the **Tool Menu**. The current code defaults load automatically, but you may change them when the study requires different limits.

#### Overall / Global
**Global** means the ratio is compared with one set of limits for the overall applicable sales population, regardless of neighborhood.

- **Too Low below**: ratios below this value are **Too Low for Global**. Current default: `0.5000`.
- **Perfect Global minimum / maximum**: ratios inside this range are **Perfect Global**. Current defaults: `0.7000–1.2000`.
- Ratios between the Too Low boundary and Perfect range, or between the Perfect range and Too High boundary, are **Acceptable Global**.
- **Too High above**: ratios above this value are **Too High for Global**. Current default: `1.5000`.

The four Global settings must remain in this order:
`Too Low ≤ Perfect Minimum ≤ Perfect Maximum ≤ Too High`.

#### Neighborhood
Neighborhood analysis compares a sale with the other usable ratios in the **same neighborhood**. A record can therefore be acceptable globally but unusual for its neighborhood, or the reverse.

Choose one neighborhood method:

- **Percentile / Quartile Limits** — the current default. The default lower boundary is the **25th percentile (Q1)** and the upper boundary is the **75th percentile (Q3)** for each neighborhood. Below the lower boundary = **Neighborhood Too Low**; above the upper boundary = **Neighborhood Too High**; between them = **Acceptable**.
- **Fixed Ratio Limits** — uses the same neighborhood low/high cutoffs for every neighborhood instead of calculating separate percentile limits.

After changing any Global or Neighborhood setting, click **Apply Threshold Settings**, then rerun **Ratio Checker** or **Quality Checker** so the statuses, Dashboard, and Neighborhood Statistics are recalculated. **Reset to Current Code Defaults** restores the original settings.

**Critical dual outlier:** this is not a separate threshold you set. In the Quality Checker, a record is treated as **Critical: Global and Neighborhood Ratio Outlier** when the same ratio is outside both its Global limit and its Neighborhood limit.

### Neighborhood Statistics
After **Ratio Checker** or **Quality Checker** runs, the app automatically creates a **🏘️ Neighborhood Statistics** panel under the Dashboard when neighborhood and ratio data are available. Click its arrow to expand it; the panel stays collapsed by default because the table can be large.

The table can show, by neighborhood:
- valid sale count;
- threshold method and the neighborhood low/high thresholds;
- Q1, Median, Q3, Mean, Minimum, and Maximum ratio;
- Neighborhood Low and Neighborhood High counts;
- Global Low and Global High counts; and
- **Critical Dual Outliers** — records that are outliers under both tests.

Use this table together with the row-level result. A neighborhood or global flag means the record needs review; it does not automatically mean the assessment or sale is wrong.

### Checker Reference
| Tool | Purpose | Main review focus |
|---|---|---|
| **Prepare Data** | Standardize and clean working data | Prepared status and mapped fields |
| **Sale Date Checker** | Validate the study-period dates | Missing or out-of-range sale dates |
| **Use Code Checker** | Review Use Code/property-component conditions | Missing/unsupported codes and land/improvement conflicts |
| **Appraisal Value Checker** | Crossfoot component values to total value | Land + improvement + miscellaneous vs. total |
| **Deed / MH / Comment Audit** | Review transaction/documentation conditions | Parties, deeds, comments, qualification, MH, duplicates, and bad-sale conditions |
| **Ratio Checker** | Calculate and evaluate L, B, and L&B ratios | Global and neighborhood ratio classifications |
| **Quality Checker** | Consolidated integrity screening | Flag Status/reasons, including critical dual ratio outliers |
| **Generate Statistics** | Create statistical summaries | Current prepared/analysis population |
| **Clear Results** | Reset prior checker output | Begin another analysis without re-uploading |

**Sale Date Checker setting:** select the Tax Year of Study. The app checks the period from **October 1 of Tax Year − 2 through September 30 of Tax Year − 1**.

**Ratio formulas:** `L = Land Value ÷ Sale Price`; `B = (Improvement Value + Miscellaneous Value) ÷ Sale Price`; `L&B = Total Value ÷ Sale Price`.

**Bad Sale with No Comment:** the audit and Quality Checker treat this as its own review condition. A bad sale does not need to already have a comment in order to be flagged for missing documentation.

### Generate Statistics and VERIFY PASS
**Generate Statistics** summarizes the current prepared data so the analyst can evaluate the study as a whole and by important groups. If a Sales Ratio field is not already available, the statistics routine first derives/creates the ratios using the mapped ratio logic. The statistics are based primarily on sales classified as **GOOD** by the mapped qualification field.

The generated output includes:
- **Use Code Statistics** — groups the study by normalized Use Code.
- **Valuation Zone Statistics** — groups the study by the mapped **Neighborhood** field so you can review each neighborhood/zone separately.
- Counts for **Total Sales, Good Sales, Bad Sales, Undetermined Sales, and Good Ratios**.
- Statistical measures such as **Mean, Median, Weighted Mean, Minimum, Maximum, Range, PRD, and COD**, plus appraisal and sale-price summaries when the required fields are available.
- A Dashboard count of **Zones with 10+ Good Sales**, which shows how many neighborhoods/zones have at least 10 sales classified as GOOD.

#### What VERIFY PASS means
The **County Study Certification** shown by Generate Statistics is an **overall screening indicator**. The current program displays **VERIFY PASS** when both of these overall conditions are met:
- the overall median ratio is between **0.9750 and 1.0244**; and
- the overall COD is **20% or less**.

If either condition is not met, the status is **REVIEW REQUIRED**.

**Important: VERIFY PASS does not mean every neighborhood, Use Code, or individual sale passed.** It also does not mean the analyst should automatically rely on the overall county status without reviewing the underlying groups. A strong overall result can still contain neighborhoods with too few sales, unusual distributions, outliers, or local patterns that deserve attention.

Before relying on **VERIFY PASS**, review the **Valuation Zone Statistics** and **Neighborhood Statistics** and ask:
- Does each important neighborhood have enough **Good Sales** and **Good Ratios** to make its statistics meaningful?
- Which neighborhoods have **10 or more Good Sales**, and which have fewer than 10? Neighborhoods with small samples should be interpreted more cautiously.
- Are the neighborhood medians, CODs, ranges, and outlier counts reasonable, or is one neighborhood being hidden by a strong overall result?
- Are there many **Bad** or **Undetermined** sales that could affect how representative the usable sample is?
- Do the Global and Neighborhood ratio classifications agree with what you see in the grouped statistics?

Think of **VERIFY PASS** as: **“the overall median and COD meet the program's current screening limits—now verify that the neighborhoods and underlying sales support that conclusion.”** Analyst review of sample size, neighborhood composition, outliers, qualification, and source data is still required before treating the overall result as dependable.

### Dashboard, Results, and Export
After a checker runs, review the **Dashboard** first, then the record-level **Analysis Results**.

- Use **Show only rows requiring review** to filter out records that appear verified/perfect and concentrate on exceptions.
- A completion message means the tool finished running; it does **not** mean every record passed.
- **OK / acceptable** means no configured exception was identified for that check.
- **Review / flagged** means one or more conditions require analyst attention.
- **Missing** means a needed field is blank or unavailable; verify the field mapping first, then the source record.
- **Calculated** means the app derived a value from mapped fields rather than relying on a supplied source value.
- **VERIFY PASS** is an overall Generate Statistics screening result based on the current overall median and COD limits. It is **not** a blanket approval of every neighborhood or sale. Review neighborhood/valuation-zone sample sizes and statistics before relying on it.
- Use **Download Processed Workbook** when the review is complete. The export carries the generated statuses/results for continued work in Excel.

### Help and Review Tips
The **📘 Help** control remains in the **left sidebar**. It is a quick-reference aid; the full user guide is the **Instructions / User Guide** panel at the top of the main page.

Recommended review practices:
- Verify field mapping before trusting any analytical result.
- Confirm **Prepared = Yes** before analytical checkers.
- Review Global and Neighborhood settings before ratio-based analysis.
- Rerun the Ratio or Quality Checker after changing thresholds.
- Use Global and Neighborhood results together instead of treating either one as the only standard.
- Use Neighborhood Statistics to understand local context before correcting a flagged ratio.
- When Generate Statistics shows **VERIFY PASS**, still review Valuation Zone/Neighborhood Statistics, especially **Good Sales Count** and **Good Ratio Count**. Small neighborhood samples can make local statistics less reliable even when the overall study passes.
- Treat flags and pass indicators as screening aids that support analyst judgment; neither one replaces review of the source data and the composition of the sales sample.
"""

SIDEBAR_HELP = """
**Quick Help**

The full **📘 Instructions / User Guide** is at the top of the main page.

- **Mapping:** verify the source fields, then click **Apply Field Mapping**.
- **Prepare:** run **Prepare Data** and confirm **Prepared = Yes**.
- **Ratio / Quality:** review **Ratio Threshold Settings** first.
- **Global:** compares the ratio with overall study limits.
- **Neighborhood:** compares the ratio with same-neighborhood limits.
- **Critical dual outlier:** the same record is outside both Global and Neighborhood limits.
- **Neighborhood Statistics:** appears under the Dashboard after Ratio or Quality analysis when data are available.
- **Generate Statistics / VERIFY PASS:** VERIFY PASS only means the overall median and COD meet the current screening limits. Still review the neighborhood/valuation-zone counts and statistics—especially whether neighborhoods have enough good sales to support the overall conclusion.

**Checker Quick Reference**

| Tool | Brief Description |
| :--- | :--- |
| **Prepare Data** | Standardizes and cleans the mapped data so the analytical checkers can run correctly. |
| **Sale Date Checker** | Checks for missing sale dates and dates outside the selected study period. |
| **Use Code Checker** | Reviews Use Codes and related land/improvement conditions for records needing review. |
| **Appraisal Value Checker** | Confirms land, improvement, and miscellaneous values reconcile to total appraised value. |
| **Deed / MH / Comment Audit** | Reviews deed, party, qualification, manufactured-home, comment, duplicate, and bad-sale conditions. |
| **Ratio Checker** | Calculates L, B, and L&B ratios and compares them with Global and Neighborhood limits. |
| **Quality Checker** | Runs a broader combined quality review and identifies the reasons a record requires attention. |
| **Generate Statistics** | Produces overall, Use Code, and neighborhood/valuation-zone statistics. A **VERIFY PASS** is an overall median/COD screening result; neighborhood sample sizes and local statistics still need review. |
| **Clear Results** | Removes the current checker output while keeping the uploaded workbook loaded. |

**Export / Download Processed Workbook**

Use **Download Processed Workbook** after reviewing the results. The exported Excel workbook keeps the original working data together with the generated checker statuses, flags, reasons, calculated fields, and other analysis output available from the current run. This lets you continue reviewing, documenting, filtering, or sharing the results outside the web application without recreating the analysis manually.
"""


def render_user_instructions():
    """Render the user-facing guide maintained directly in this application."""
    st.markdown(APP_INSTRUCTIONS, unsafe_allow_html=False)


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
    with st.popover("Help / Quick Reference", use_container_width=True):
        st.markdown(SIDEBAR_HELP)

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
    st.caption("Need help? Open **📘 Instructions / User Guide — Click to Open** at the top of the main page. The **📘 Help** control in the sidebar provides a quick reference.")
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
