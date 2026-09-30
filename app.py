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
## 📘 SALES RATIO QUALITY CHECKER
### In-App Instructions / User Guide
**Updated September 30, 2026**

Use this guide with the live application to move from upload and field mapping through checker review, neighborhood context, statistics, and export. The full guide is built directly into the app; the sidebar **Help / Quick Reference** remains a shorter quick reference.

**Workflow:** **1 Upload → 2 Verify Mapping → 3 Prepare → 4 Review Thresholds when appropriate → 5 Run Checker → 6 Review Results → 7 Statistics / Export**

> **Screening supports analyst judgment; it does not replace source-record review.**

### Contents
- [1. Quick Start](#1-quick-start)
- [2. Upload and Field Mapping](#2-upload-and-field-mapping)
- [3. Prepare Data and Tool Menu](#3-prepare-data-and-tool-menu)
- [4. Review Thresholds](#4-review-thresholds)
- [5. Neighborhood Statistics](#5-neighborhood-statistics)
- [6. Checker Reference](#6-checker-reference)
- [7. Generate Statistics and VERIFY PASS](#7-generate-statistics-and-verify-pass)
- [8. Understanding Statuses and Results](#8-understanding-statuses-and-results)
- [9. Export and Highlighting](#9-export-and-highlighting)
- [10. Help and Final Review Checklist](#10-help-and-final-review-checklist)

> **Key idea:** “pass,” “verified,” “perfect,” and “prepared” are context-specific statuses. None of them should be read as a blanket approval of every neighborhood, record, or source field. Always review the scope of the status and the underlying data.

---

### 1. Quick Start
**Recommended workflow:** Upload → Verify Mapping → Apply Mapping → Prepare Data → Review Thresholds → Run Checker → Review Dashboard / Neighborhoods / Rows → Generate Statistics if needed → Export

1. Upload the CAMA / sales-ratio workbook from **1 · Load data** in the left sidebar.
2. Open **2 · Field Mapping Review** and verify every proposed source field.
3. If more than one source field could qualify, choose the field that should be authoritative for the analysis.
4. Click **Apply Field Mapping**.
5. Click **Prepare Data** and confirm **Prepared = Yes**.
6. For ratio-based work, review **Ratio Threshold Settings** before running **Ratio Checker** or **Quality Checker**.
7. Run the desired checker, review the **Dashboard**, open **Neighborhood Statistics** when available, and inspect **Analysis Results**.
8. Use **Show only rows requiring review** for exceptions; use **Generate Statistics** for overall/grouped statistics; then export the processed workbook.

> **Before relying on a result:** verify the mapping, confirm **Prepared = Yes**, confirm the active ratio settings, check whether the relevant neighborhood has enough usable sales to make the comparison meaningful, and compare the result with the source record.

---

### 2. Upload and Field Mapping
The app looks for an **Analysis** worksheet first and otherwise uses the first worksheet in the uploaded Excel file. Supported upload types are `.xlsx`, `.xlsm`, `.xls`, and `.xlsb`.

**Field Mapping Review** connects vendor-specific source columns to the standardized fields used by the checker engine.

- Open **Review / Change Field Mapping** with its arrow.
- Review every proposed field before relying on checker output.
- If more than one source field is a reasonable match, select the source column you want the checker to use.
- Click **Apply Field Mapping** after making changes.
- Expand **Mapping Details** to verify the detected header, standardized field, populated-row count, candidate headers, and mapping method.
- For Delta files, **USE CODE** remains the authoritative Use Code unless you explicitly choose another field.
- Four-digit Use Codes are normalized by removing one outside zero when applicable: `0100 → 100`, `1000 → 100`, `0101 → 101`, `1010 → 101`; `1001` remains `1001`.

> **Why this matters:** a correct analytical rule applied to the wrong source field can still create a misleading result. Field mapping should be verified before analysis.

---

### 3. Prepare Data and Tool Menu
**Prepare Data** creates the clean, standardized working dataset used by the analytical tools. Run it after applying field mapping and whenever you upload a new workbook or materially change the mapping.

> **Prepared = Yes** means the data are standardized and ready for analytical tools. It does **not** mean the records passed review.

| Tool | Purpose |
|---|---|
| **Prepare Data** | Standardize and clean mapped working data. |
| **Sale Date Checker** | Validate the selected study-period dates. |
| **Use Code Checker** | Review Use Code and land/improvement conditions. |
| **Appraisal Value Checker** | Crossfoot component values to total appraised value. |
| **Deed / MH / Comment Audit** | Review documentation and transaction conditions. |
| **Ratio Checker** | Calculate L, B, and L&B ratios and classify Global/Neighborhood results. |
| **Quality Checker** | Run consolidated integrity screening and detailed reasons. |
| **Generate Statistics** | Create overall, Use Code, and valuation-zone/neighborhood summaries. |
| **Clear Results** | Remove current checker output while keeping the workbook loaded. |

---

### 4. Review Thresholds
The ratio controls are located between **Field Mapping Review** and the **Tool Menu**. Current code defaults load automatically, but they can be changed when the study requires different limits.

#### Overall / Global
**Global** compares each applicable ratio with one set of limits for the overall sales population, regardless of neighborhood.

| Setting | Meaning | Default |
|---|---|---:|
| **Too Low below** | Below this value = Too Low for Global | `0.5000` |
| **Perfect Global minimum** | Lower edge of Perfect Global | `0.7000` |
| **Perfect Global maximum** | Upper edge of Perfect Global | `1.2000` |
| **Too High above** | Above this value = Too High for Global | `1.5000` |

Ratios between the Too Low boundary and Perfect range, or between the Perfect range and Too High boundary, are classified as **Acceptable Global**. The four settings must remain in this order: **Too Low ≤ Perfect Minimum ≤ Perfect Maximum ≤ Too High**.

#### Neighborhood
Neighborhood analysis compares a sale with the other usable ratios in the **same neighborhood**. A record can therefore look acceptable globally but unusual locally, or the reverse.

- **Percentile / Quartile Limits** — current default. Default lower boundary = **25th percentile (Q1)**; default upper boundary = **75th percentile (Q3)** for each neighborhood.
- **Fixed Ratio Limits** — uses the same neighborhood low/high cutoffs for every neighborhood instead of calculating separate percentile limits.

After changing settings, click **Apply Threshold Settings**, then rerun **Ratio Checker** or **Quality Checker**. **Reset to Current Code Defaults** restores the original settings.

> **Critical dual outlier:** this is not a separate threshold. It means the same ratio is outside both its Global limit and its Neighborhood limit.

---

### 5. Neighborhood Statistics
After **Ratio Checker** or **Quality Checker** runs, the app creates a **🏘️ Neighborhood Statistics** panel under the Dashboard when neighborhood and ratio data are available. The panel stays collapsed by default because the table can be large.

The table can show, by neighborhood:
- valid sale count;
- threshold method and neighborhood low/high thresholds;
- Q1, Median, Q3, Mean, Minimum, and Maximum ratio;
- Neighborhood Low and Neighborhood High counts;
- Global Low and Global High counts; and
- **Critical Dual Outliers** — records that are outliers under both tests.

#### How to judge the neighborhood context
- Do not rely on the Global classification alone; compare the sale with its local neighborhood result.
- Check how many usable **Good Sales / Good Ratios** support the neighborhood statistics.
- Look for unusual medians, wide ranges, high dispersion, or repeated outliers that may be hidden by a strong overall result.
- Compare the grouped statistics with the actual source records before deciding a correction is necessary.

> **Sample-size caution:** Generate Statistics reports **Zones with 10+ Good Sales** as a review aid. Fewer than 10 Good Sales does not automatically make a neighborhood invalid, but a small sample should be interpreted more cautiously. The 10-sale count is a screening indicator, not a replacement for applicable study standards or analyst judgment.

---

### 6. Checker Reference
| Tool | Main review focus | Key note |
|---|---|---|
| **Sale Date Checker** | Missing/out-of-range dates | Set/confirm **Tax Year of Study** before clicking the checker; the selected tax year controls the study period. |
| **Use Code Checker** | Use Code/property-component issues | Review code and land/improvement conditions together. |
| **Appraisal Value Checker** | Component-to-total reconciliation | Land + improvement + miscellaneous vs. total. |
| **Deed / MH / Comment Audit** | Transaction/documentation issues | Parties, deed, qualification, comments, MH, duplicates, bad-sale documentation. |
| **Ratio Checker** | Ratio classifications | L, B, L&B plus Global and Neighborhood status. |
| **Quality Checker** | Consolidated screening | Flag Status and detailed reason; includes critical dual ratio outliers. |
| **Generate Statistics** | Study summaries | Overall, Use Code, and valuation-zone/neighborhood statistics. |

#### Sale Date Checker — set the Tax Year first
When using the **Sale Date Checker**, set or confirm the tax year **before** clicking the checker button:

1. Under **3 · Tool Menu**, expand **Sale Date Checker setting**.
2. Enter or confirm the **Tax Year of Study**.
3. Review the study-period message shown beneath the setting to make sure it is the period you intend to test.
4. Then click **📅 Sale Date Checker**.

> **Important:** Always set/confirm the **Tax Year of Study first**, then run **Sale Date Checker**. The selected tax year determines which sale-date study period the checker uses.

**Sale Date study period:** October 1 of **Tax Year − 2** through September 30 of **Tax Year − 1**.

**Ratio formulas:** `L = Land Value ÷ Sale Price`; `B = (Improvement Value + Miscellaneous Value) ÷ Sale Price`; `L&B = Total Value ÷ Sale Price`.

**Bad Sale with No Comment:** this is its own review condition; a bad sale does not need to already have a comment to be flagged for missing documentation.

---

### 7. Generate Statistics and VERIFY PASS
**Generate Statistics** summarizes the current prepared data so the analyst can evaluate the study overall and by important groups. If a Sales Ratio field is not already available, the routine first derives/creates ratios using the mapped ratio logic. The statistics are based primarily on sales classified as **GOOD** by the mapped qualification field.

Generated output includes:
- **Use Code Statistics** — grouped by normalized Use Code.
- **Valuation Zone Statistics** — grouped by the mapped Neighborhood field.
- Counts for **Total Sales, Good Sales, Bad Sales, Undetermined Sales, and Good Ratios**.
- **Mean, Median, Weighted Mean, Minimum, Maximum, Range, PRD, and COD**.
- Dashboard count of **Zones with 10+ Good Sales**.

#### What the main statistics mean
| Statistic | Meaning in the app |
|---|---|
| **Mean** | Average of the usable Good Sale ratios. |
| **Median** | Middle usable Good Sale ratio after ordering the ratios. |
| **Weighted Mean** | Total appraised value ÷ total sale price for available Good Sales. |
| **Range** | Maximum ratio − minimum ratio. |
| **PRD** | Mean ÷ Weighted Mean; a distribution review measure. |
| **COD** | Average absolute deviation from the median ÷ median; lower values indicate tighter dispersion. |

#### What VERIFY PASS means
The **County Study Certification** is an **overall screening indicator**. The current program displays **VERIFY PASS** only when both of these conditions are met:
- Overall median ratio is between **0.9750 and 1.0244**.
- Overall COD is **20% or less**.

If either condition is not met, the program displays **REVIEW REQUIRED**.

> **VERIFY PASS is not a blanket approval.** It does not mean every neighborhood, Use Code, or individual sale passed, and it does not prove that every subgroup has enough sales to be dependable.

Before relying on VERIFY PASS, review **Valuation Zone Statistics** and **Neighborhood Statistics** and ask:
- Does each important neighborhood have enough **Good Sales** and **Good Ratios** to make its statistics meaningful?
- Which neighborhoods have **10 or more Good Sales**, and which have fewer than 10?
- Are neighborhood medians, CODs, ranges, and outlier counts reasonable?
- Could a strong overall result be hiding a weak or unusual neighborhood?
- Are there many Bad or Undetermined sales that make the usable sample less representative?
- Do the Global and Neighborhood classifications agree with the grouped statistics and source records?

> **Think of VERIFY PASS as:** “The overall median and COD meet the program’s current screening limits. Now verify that the neighborhoods and underlying sales support that conclusion.”

---

### 8. Understanding Statuses and Results
After a checker runs, review the **Dashboard** first and then the record-level **Analysis Results**. A status only describes the scope of the current checker or statistical test.

| Status / message | What it means | What it does NOT mean |
|---|---|---|
| **Prepared = Yes** | Data are standardized and ready for checker use. | The file passed quality review. |
| **Verified Compliant / Verified Active Sale** | No configured exception was found for that row under the current checker. | Every other checker, neighborhood, or source field is automatically correct. |
| **Perfect Global / Acceptable Global** | Ratio is within the configured Global classification band. | The ratio is also acceptable for its neighborhood. |
| **Review / Flagged / Critical / Error** | One or more configured conditions require attention. | The assessment or sale is automatically wrong. |
| **Missing** | A needed field is blank/unavailable to the checker. | The source system definitely lacks the information; verify mapping first. |
| **Calculated** | The app derived a value from mapped source fields. | The value came directly from a source ratio field. |
| **VERIFY PASS** | Overall Generate Statistics median and COD meet the current screening limits. | Every neighborhood, Use Code, or sale passed. |
| **Completion message** | The selected tool finished running. | Every record passed. |

A flag is a prompt for analyst review. Read the **Flag Status**, reason/detail fields, and the related source cells before deciding that a correction is necessary.

Use **Show only rows requiring review** to focus the table on exceptions.

---

### 9. Export and Highlighting
**Download Processed Workbook** keeps the working data together with generated statuses, flags, reasons, calculated fields, and available analysis tables so the review can continue in Excel.

The **Analysis** worksheet is designed to make exceptions easy to locate:

| Export element | What it means | Visual treatment |
|---|---|---|
| **Issue row** | Only rows requiring attention receive row-wide issue shading. | Yellow-toned review shading |
| **Specific issue cell(s)** | The source cell(s) most directly tied to the issue receive a stronger contrasting highlight. | Stronger orange-toned highlight |
| **Critical / error issue** | Critical/error exceptions use stronger row and source-cell emphasis. | Red-toned highlighting |
| **Verified / good row** | No row-wide issue shading. Flag Status text and its normal status color remain. | Normal verified/status color |

The export highlights likely source fields to make review faster, but the highlighted cell is a **locator—not an automatic correction instruction**. Confirm the issue against the Flag Status and source record.

The exported Dashboard also includes a highlighting legend so the colors can be interpreted when the workbook is opened outside the app.

---

### 10. Help and Final Review Checklist
The **Help** control remains in the left sidebar for quick reminders. The full guide is the **Instructions / User Guide** panel at the top of the main page.

#### Final review checklist
- [ ] Field mapping is correct, including any manual mapping choices.
- [ ] **Prepared = Yes** before analytical checkers are run.
- [ ] Global and Neighborhood settings match the intended study.
- [ ] Ratio or Quality Checker was rerun after any threshold change.
- [ ] Dashboard totals were reviewed before individual rows.
- [ ] Neighborhood Statistics were reviewed for local context and usable sale count.
- [ ] If Generate Statistics shows **VERIFY PASS**, neighborhood/valuation-zone counts and statistics were still reviewed.
- [ ] Flags were treated as review prompts, not automatic proof of errors.
- [ ] Pass/verified indicators were treated as screening results, not substitutes for source-record review.
- [ ] The processed workbook was exported when a documented copy of the analysis was needed.

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

