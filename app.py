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
## 📘 Sales Ratio Quality Checker — User Guide

This is the live user guide for the application. The full guide stays here at the top of the main page, while **Help / Quick Reference** remains in the left sidebar for shorter reminders.

> **Best practice:** Treat the checker as a screening and review tool. Before relying on any result, verify the field mapping, confirm **Prepared = Yes**, review the active ratio settings, and make sure the neighborhood/valuation-zone data are strong enough to support the conclusion.

### Contents
- [Quick Start](#quick-start)
- [Upload and Field Mapping](#upload-and-field-mapping)
- [Prepare Data and Tool Menu](#prepare-data-and-tool-menu)
- [Review Thresholds](#review-thresholds)
- [Neighborhood Statistics](#neighborhood-statistics)
- [Checker Reference](#checker-reference)
- [Generate Statistics and VERIFY PASS](#generate-statistics-and-verify-pass)
- [Understanding Statuses and Results](#understanding-statuses-and-results)
- [Export and Highlighting](#export-and-highlighting)
- [Help and Final Review Checklist](#help-and-final-review-checklist)

### Quick Start
**Recommended workflow:** Upload → Verify Mapping → Apply Mapping → Prepare Data → Review Thresholds When Appropriate → Run Checker → Review Dashboard/Rows → Generate Statistics if needed → Export

1. Upload the CAMA / sales-ratio workbook from **1 · Load data** in the left sidebar.
2. Review the detected vendor and open **2 · Field Mapping Review**.
3. Verify every proposed source column. If more than one source field could qualify, choose the field that should be authoritative for the analysis.
4. Click **Apply Field Mapping**.
5. Click **Prepare Data** in **3 · Tool Menu** and confirm the top summary shows **Prepared = Yes**.
6. For **Ratio Checker** or **Quality Checker**, review **Ratio Threshold Settings** before running the analysis.
7. Run the desired checker from **3 · Tool Menu**.
8. Review **4 · Dashboard**, expand **🏘️ Neighborhood Statistics** when available, and then review **5 · Analysis Results**.
9. Use **Show only rows requiring review** when you want to concentrate on exceptions.
10. If you run **Generate Statistics**, review both the overall result and the grouped neighborhood/valuation-zone statistics before relying on **VERIFY PASS**.
11. Use **6 · Export** to download the processed workbook.

**Before trusting any result, confirm:** the mapping is correct; **Prepared = Yes**; the active Global/Neighborhood settings are appropriate; the relevant neighborhood has enough usable sales to make its statistics meaningful; and the source data support the result.

### Upload and Field Mapping
The app looks for an **Analysis** worksheet first and otherwise uses the first worksheet in the uploaded Excel file. Supported upload types are `.xlsx`, `.xlsm`, `.xls`, and `.xlsb`.

**Field Mapping Review** connects vendor-specific source columns to the standardized fields used by the checker engine.

- Open **Review / Change Field Mapping** with its arrow.
- Review every proposed field before relying on checker output.
- If more than one source field is a reasonable match, select the source column you want the checker to use.
- Click **Apply Field Mapping** after making changes.
- Expand **Mapping Details** to verify the detected header, standardized field, populated-row count, candidate headers, and mapping method.
- For Delta files, **USE CODE** remains the authoritative Use Code unless you explicitly choose another field.
- Four-digit Use Codes are normalized by removing one outside zero when applicable: `0100 → 100`, `1000 → 100`, `0101 → 101`, `1010 → 101`; `1001` remains `1001`.

**Why this matters:** a correct checker rule applied to the wrong source field can still produce a misleading result. Mapping should always be verified before analysis.

### Prepare Data and Tool Menu
**Prepare Data** creates the clean, standardized working dataset used by the analytical tools. Run it after applying field mapping and whenever you upload a new workbook or materially change the mapping.

**Prepared = Yes** means the data have been standardized and are ready for the analytical tools. It does **not** mean the records have passed review.

| Tool | Primary purpose |
|---|---|
| **Prepare Data** | Standardize and clean the mapped working data. |
| **Sale Date Checker** | Check missing dates and dates outside the selected study period. |
| **Use Code Checker** | Review Use Codes and related land/improvement conditions. |
| **Appraisal Value Checker** | Reconcile land + improvement + miscellaneous value to total appraised value. |
| **Deed / MH / Comment Audit** | Review deed, party, qualification, manufactured-home, comment, duplicate, and bad-sale conditions. |
| **Ratio Checker** | Calculate L, B, and L&B ratios and compare Global and Neighborhood limits. |
| **Quality Checker** | Run a broader combined screening and identify why a record needs attention. |
| **Generate Statistics** | Produce overall, Use Code, and neighborhood/valuation-zone statistics. |
| **Clear Results** | Remove current checker output while keeping the uploaded workbook loaded. |

**Clear Results** is useful when you want a clean result view for another checker without uploading the workbook again.

### Review Thresholds
Use **Ratio Threshold Settings** between **Field Mapping Review** and the **Tool Menu** to confirm the active limits before ratio-based analysis. Current code defaults load automatically, but you may change them when the study requires different limits.

#### Overall / Global
**Global** compares each applicable ratio with one set of limits for the overall sales population, regardless of neighborhood.

| Setting | Meaning | Current default |
|---|---|---:|
| **Too Low below** | Ratio below this value = Too Low for Global | `0.5000` |
| **Perfect Global minimum** | Lower edge of the Perfect Global range | `0.7000` |
| **Perfect Global maximum** | Upper edge of the Perfect Global range | `1.2000` |
| **Too High above** | Ratio above this value = Too High for Global | `1.5000` |

Ratios between the Too Low boundary and Perfect range, or between the Perfect range and Too High boundary, are classified as **Acceptable Global**. The settings must remain in this order: `Too Low ≤ Perfect Minimum ≤ Perfect Maximum ≤ Too High`.

#### Neighborhood
Neighborhood analysis compares a sale with the other usable ratios in the **same neighborhood**. A record can therefore look acceptable globally but unusual locally, or the reverse.

- **Percentile / Quartile Limits** — current default. The lower boundary is the selected low percentile (default **25th percentile / Q1**) and the upper boundary is the selected high percentile (default **75th percentile / Q3**) for each neighborhood.
- **Fixed Ratio Limits** — applies the same neighborhood low/high limits to every neighborhood instead of calculating separate percentile limits.

After changing any Global or Neighborhood setting, click **Apply Threshold Settings**, then rerun **Ratio Checker** or **Quality Checker**. The statuses, Dashboard, and Neighborhood Statistics must be recalculated under the new settings. **Reset to Current Code Defaults** restores the original settings.

**Critical dual outlier:** this is not a separate threshold. In the Quality Checker, **Critical: Global and Neighborhood Ratio Outlier** means the same ratio is outside both its Global limit and its Neighborhood limit.

### Neighborhood Statistics
After **Ratio Checker** or **Quality Checker** runs, the app creates a **🏘️ Neighborhood Statistics** panel under the Dashboard when neighborhood and ratio data are available. The panel is collapsed by default because the table can be large.

The table can show, by neighborhood:
- valid sale count;
- threshold method and neighborhood low/high thresholds;
- Q1, Median, Q3, Mean, Minimum, and Maximum ratio;
- Neighborhood Low and Neighborhood High counts;
- Global Low and Global High counts; and
- **Critical Dual Outliers** — records that are outliers under both tests.

**How to use it:** do not judge an unusual ratio only from the Global result. Check whether it is also unusual compared with sales in the same neighborhood and whether that neighborhood has enough usable sales to make the local comparison meaningful.

The **Zones with 10+ Good Sales** measure in Generate Statistics is a useful review aid. A neighborhood with fewer than 10 Good Sales is not automatically wrong or unusable, but its local statistics should be interpreted more cautiously. The app's 10-sale count is a screening indicator, not a substitute for the analyst's applicable study standards and judgment.

### Checker Reference
| Tool | Main review focus |
|---|---|
| **Sale Date Checker** | Missing or out-of-range sale dates for the selected study period. |
| **Use Code Checker** | Missing/unsupported Use Codes and land/improvement conflicts. |
| **Appraisal Value Checker** | Component-to-total value reconciliation. |
| **Deed / MH / Comment Audit** | Parties, deed data, comments, qualification, MH indicators, duplicates, and bad-sale documentation. |
| **Ratio Checker** | L, B, and L&B ratio calculation plus Global and Neighborhood classifications. |
| **Quality Checker** | Consolidated row-level integrity screening and detailed reason/status. |
| **Generate Statistics** | Overall, Use Code, and valuation-zone/neighborhood statistical summaries. |

**Sale Date Checker setting:** select the **Tax Year of Study**. The app checks the period from **October 1 of Tax Year − 2 through September 30 of Tax Year − 1**.

**Ratio formulas:** `L = Land Value ÷ Sale Price`; `B = (Improvement Value + Miscellaneous Value) ÷ Sale Price`; `L&B = Total Value ÷ Sale Price`.

**Bad Sale with No Comment:** the audit and Quality Checker treat this as its own review condition. A bad sale does not need to already have a comment in order to be flagged for missing documentation.

### Generate Statistics and VERIFY PASS
**Generate Statistics** summarizes the current prepared data so the analyst can evaluate the study overall and by important groups. If a Sales Ratio field is not already available, the routine first derives/creates ratios using the mapped ratio logic. Statistics are based primarily on sales classified as **GOOD** by the mapped qualification field.

The generated output includes:
- **Use Code Statistics** — grouped by normalized Use Code.
- **Valuation Zone Statistics** — grouped by the mapped Neighborhood field.
- counts for **Total Sales, Good Sales, Bad Sales, Undetermined Sales, and Good Ratios**;
- **Mean, Median, Weighted Mean, Minimum, Maximum, Range, PRD, and COD**; and
- a Dashboard count of **Zones with 10+ Good Sales**.

#### What the main statistics mean
| Statistic | Plain-language meaning in this app |
|---|---|
| **Mean** | Average of the usable Good Sale ratios. |
| **Median** | Middle usable Good Sale ratio after the ratios are ordered. |
| **Weighted Mean** | Total appraised value divided by total sale price for available Good Sales. |
| **Range** | Maximum ratio minus minimum ratio. |
| **PRD** | Mean divided by Weighted Mean; use it as a distribution review measure. |
| **COD** | Average absolute deviation from the median divided by the median; lower values indicate tighter ratio dispersion. |

#### What VERIFY PASS means
The **County Study Certification** is an **overall screening indicator**. The current program displays **VERIFY PASS** when both conditions are met:
- overall median ratio is between **0.9750 and 1.0244**; and
- overall COD is **20% or less**.

If either condition is not met, the program displays **REVIEW REQUIRED**.

> **VERIFY PASS is not a blanket approval.** It does not mean every neighborhood, Use Code, or individual sale passed, and it does not prove that every subgroup has enough sales to be dependable.

Before relying on VERIFY PASS, review the **Valuation Zone Statistics** and **Neighborhood Statistics** and ask:
- Does each important neighborhood have enough **Good Sales** and **Good Ratios** to make the statistics meaningful?
- Which neighborhoods have **10 or more Good Sales**, and which have fewer than 10?
- Are neighborhood medians, CODs, ranges, and outlier counts reasonable?
- Could a strong overall result be hiding a weak or unusual neighborhood?
- Are there many Bad or Undetermined sales that make the usable sample less representative?
- Do the Global and Neighborhood classifications agree with the grouped statistics and source records?

Think of **VERIFY PASS** as: **“The overall median and COD meet the program's current screening limits. Now verify that the neighborhoods and underlying sales support that conclusion.”**

### Understanding Statuses and Results
After a checker runs, review the **Dashboard** first and then the record-level **Analysis Results**.

| Status / message | What it means | What it does NOT mean |
|---|---|---|
| **Prepared = Yes** | Data are standardized and ready for checker use. | The file passed quality review. |
| **Verified Compliant / Verified Active Sale** | No configured exception was found for that row under the current checker. | Every other checker, neighborhood, or source record is automatically correct. |
| **Perfect Global / Acceptable Global** | The ratio falls within the configured Global classification band. | The ratio is also acceptable for its neighborhood. |
| **Review / Flagged / Critical / Error** | One or more configured conditions require analyst attention. | The assessment or sale is automatically wrong. |
| **Missing** | A field needed by the checker is blank or unavailable. | The source system definitely lacks the information; verify mapping and source data first. |
| **Calculated** | The app derived a value from mapped source fields. | The value came directly from a source sales-ratio field. |
| **VERIFY PASS** | Overall Generate Statistics median and COD meet current screening limits. | Every neighborhood, Use Code, or sale passed. |
| **Completion message** | The selected tool finished running. | Every record passed. |

Use **Show only rows requiring review** to hide rows that appear verified/perfect and focus on exceptions. Always read the **Flag Status**, reason/detail fields, and the relevant source cells before deciding that a correction is needed.

### Export and Highlighting
Use **Download Processed Workbook** after reviewing the results. The export keeps the working data together with the generated statuses, flags, reasons, calculated fields, and available analysis tables.

The **Analysis** worksheet uses highlighting to make exceptions easier to locate:
- **Issue row:** only a row requiring attention receives row-wide issue shading.
- **Specific issue cell(s):** the source cell(s) most directly tied to the issue receive a stronger contrasting highlight.
- **Critical/error issue:** critical/error rows use red-toned row and source-cell highlighting.
- **Verified/good row:** no row-wide issue shading is applied; the normal **Flag Status** text and status color remain.

The export also includes a Dashboard highlighting legend so the colors can be interpreted outside the app.

### Help and Final Review Checklist
The **📘 Help** control remains in the left sidebar for quick reminders. The full guide is the **Instructions / User Guide** panel at the top of the main page.

Before finishing a review:
- Verify field mapping and any manual mapping choices.
- Confirm **Prepared = Yes**.
- Confirm the active Global and Neighborhood settings are appropriate.
- Rerun Ratio/Quality after changing thresholds.
- Review the Dashboard before the row-level results.
- Review Neighborhood Statistics for local context and sample size.
- If Generate Statistics shows **VERIFY PASS**, still review neighborhood/valuation-zone counts and statistics.
- Treat flags as prompts for analyst review, not automatic proof of an error.
- Treat pass/verified indicators as screening results, not substitutes for source-record review.
- Export the processed workbook when you need a documented copy of the analysis.

"""

SIDEBAR_HELP = """
**Quick Help**

The full **📘 Instructions / User Guide** is at the top of the main page.

- **Mapping:** verify the source fields, then click **Apply Field Mapping**.
- **Prepare:** run **Prepare Data** and confirm **Prepared = Yes**. This means the file is ready for analysis—not that it passed.
- **Ratio / Quality:** review **Ratio Threshold Settings** first.
- **Global:** compares the ratio with the overall study limits.
- **Neighborhood:** compares the ratio with same-neighborhood limits.
- **Critical dual outlier:** the same ratio is outside both Global and Neighborhood limits.
- **Neighborhood Statistics:** use the local sale count and statistics before relying on a neighborhood classification.
- **Verified Compliant:** no configured exception was found for that row in the current checker; still review other relevant checks and source data.
- **Generate Statistics / VERIFY PASS:** VERIFY PASS only means the overall median and COD meet the current screening limits. It does **not** mean every neighborhood or sale passed. Review the neighborhood/valuation-zone counts, Good Sales, Good Ratios, and local statistics before relying on the overall status.

**Checker Quick Reference**

| Tool | Brief Description |
| :--- | :--- |
| **Prepare Data** | Standardizes and cleans the mapped data for analysis. |
| **Sale Date Checker** | Checks for missing sale dates and dates outside the selected study period. |
| **Use Code Checker** | Reviews Use Codes and related land/improvement conditions. |
| **Appraisal Value Checker** | Confirms component values reconcile to total appraised value. |
| **Deed / MH / Comment Audit** | Reviews deed, party, qualification, manufactured-home, comment, duplicate, and bad-sale conditions. |
| **Ratio Checker** | Calculates L, B, and L&B ratios and compares Global and Neighborhood limits. |
| **Quality Checker** | Runs a broader combined review and identifies why a record requires attention. |
| **Generate Statistics** | Produces overall, Use Code, and neighborhood/valuation-zone statistics and the overall VERIFY PASS / REVIEW REQUIRED screening status. |
| **Clear Results** | Removes current checker output while keeping the uploaded workbook loaded. |

**Export / Download Processed Workbook**

Use **Download Processed Workbook** after reviewing the results. Only rows with issues receive row-wide shading, and the specific source cell(s) tied to the issue receive a stronger contrasting highlight. Critical/error exceptions use red-toned highlighting. Good/verified rows remain unshaded except for the normal **Flag Status** color.

"""


def render_user_instructions():
    """Render the user-facing guide maintained directly in this application."""
    st.markdown(APP_INSTRUCTIONS, unsafe_allow_html=False)


st.title("📊 Sales Ratio Quality Checker")

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

