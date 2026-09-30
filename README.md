# 📊 Sales Ratio Quality Checker

> **A browser-based CAMA / sales-ratio quality-control and analysis application built with Python and Streamlit.**

The **Sales Ratio Quality Checker** converts the original Excel/VBA workflow into a browser-based application that can upload CAMA and sales-ratio reports, standardize different vendor layouts, run quality-control checks, calculate sales ratios and statistics, review flagged records, and export processed results to Excel.

[![Open the Live App](https://img.shields.io/badge/Open-Live%20Application-FF4B4B?logo=streamlit&logoColor=white)](https://sale-ratio-quality-checker.streamlit.app/)

---

## 🚀 Quick Start

1. Open the **Sales Ratio Quality Checker**.
2. Upload a CAMA / sales-ratio workbook.
3. Review the detected vendor.
4. Open **Field Mapping Review** and verify the proposed source fields.
5. Change any mapping that is not the field you want analyzed.
6. Click **Apply Field Mapping**.
7. Click **Prepare Data** and confirm **Prepared = Yes**.
8. Review **Ratio Threshold Settings** when using Ratio or Quality analysis. The defaults already match the current checker code.
9. Run the desired checker.
10. Review the **Dashboard** and **Analysis Results**.
11. Use **Show only rows requiring review** to focus on exceptions.
12. Generate statistics when needed.
13. Download the processed workbook.

> 💡 **In-app help:** The application now includes a collapsible **📘 Instructions / User Guide** in the sidebar. It remains available while you work so you do not have to leave the application to review the workflow or checker descriptions.

---


## 📘 Where to Open the Instructions in the App

The live Streamlit app gives users two easy ways to open this README as the user guide:

1. **Main page:** click **📘 Instructions / User Guide — Click to Open** directly below the app title.
2. **Sidebar:** under **📘 Help**, click **Open Instructions / User Guide**.

Both views load their content directly from this `README.md`. After you update and commit the README in GitHub and Streamlit redeploys the repository, the in-app guide displays the updated instructions.

---

## 🔄 README ↔ In-App Instructions Sync

`README.md` is the **single source of truth** for the app's instructions.

The Streamlit application reads `README.md` directly from the same deployment folder as `app.py` and renders it inside:

> **📘 Instructions / User Guide (README)**

### What this means

```text
Edit README.md in GitHub
        ↓
Commit / push the change
        ↓
Streamlit redeploys the repository
        ↓
The in-app Instructions panel shows the updated README
```

You no longer need to edit instruction text separately inside `app.py`.

> **Important:** The change appears in the live app after Streamlit has pulled/redeployed the new repository version. Editing the README on GitHub does not change an already-running deployment before that update is deployed.

---

## 📑 Table of Contents

- [README ↔ In-App Instructions Sync](#-readme--in-app-instructions-sync)
- [Purpose](#-purpose)
- [Application Workflow](#-application-workflow)
- [Main Features](#-main-features)
- [Supported Report Formats](#-supported-report-formats)
- [Vendor Mapping Layer](#-vendor-mapping-layer)
- [Application Tools](#-application-tools)
- [Ratio Threshold Settings](#️-ratio-threshold-settings)
- [Neighborhood Statistics](#️-neighborhood-statistics)
- [Dashboard and Results](#-dashboard-and-results)
- [Excel Export](#-excel-export)
- [Application File Structure](#-application-file-structure)
- [Running Locally](#-running-locally)
- [Deploying to Streamlit Community Cloud](#️-deploying-to-streamlit-community-cloud)
- [Recommended User Workflow](#-recommended-user-workflow)
- [Adding a New Vendor](#-adding-a-new-vendor)
- [Important Notes](#️-important-notes)

---

## 🎯 Purpose

CAMA and sales-ratio exports do not always use the same column names or report layouts.

The application solves this by using a **vendor mapping layer** to translate supported source fields into a common structure before analytical tools run. This allows the checker logic to operate on standardized fields instead of requiring every uploaded report to use identical headers.

---

## 🔄 Application Workflow

```text
Upload Report
      ↓
Detect Vendor
      ↓
Review / Override Field Mapping
      ↓
Apply Field Mapping
      ↓
Prepare Data
      ↓
Run Selected Checker
      ↓
Review Dashboard & Analysis Results
      ↓
Generate Statistics (when needed)
      ↓
Export Processed Workbook
```

---

## ✨ Main Features

- 📁 Upload Excel workbooks directly in the browser.
- 📊 Supports `.xlsx`, `.xlsm`, `.xls`, and `.xlsb`.
- 🔎 Uses the `Analysis` worksheet when one exists; otherwise uses the first worksheet.
- 🏢 Detects supported vendor/report layouts.
- 🔀 Maps vendor-specific fields into standardized application fields.
- 🧭 Provides a **collapsible Field Mapping Review** with user overrides.
- 🔎 Provides a collapsible **Mapping Details** report for verification.
- 💾 Preserves working data during the browser session.
- 🧰 Runs individual diagnostic tools without Microsoft Excel or VBA on the server.
- 🚩 Identifies records requiring analyst review.
- 📐 Calculates type-aware sales ratios.
- 📈 Generates sales-ratio statistics.
- 🏘️ Produces neighborhood-level ratio statistics and outlier counts.
- ⚙️ Provides user-adjustable **overall/global and neighborhood ratio thresholds**, pre-set to the current code defaults.
- 🗂️ Creates Use Code and Valuation Zone / Neighborhood statistical tables where available.
- 📥 Exports processed results and dashboards to Excel.
- 📘 Includes **in-app instructions loaded directly from `README.md`**, so documentation only needs to be maintained in one place.

---

## 🏢 Supported Report Formats

The mapping layer was developed from supplied CAMA report examples and the original VBA workflow.

| Vendor / Format | Supported |
|---|:---:|
| Assurance | ✅ |
| Delta | ✅ |
| S&W | ✅ |
| Capture | ✅ |
| Ingenuity | ✅ |

> **Note:** Vendor exports may use different names for the same underlying information. The mapping layer translates recognized source headers into standardized application fields.

### Standardized Application Fields

| Category | Standardized Fields |
|---|---|
| Property | Parcel Number, Neighborhood, Use Code |
| Sale | Sale Price, Sale Date, Sales Ratio, Qualification |
| Appraisal | Total Value, Land Value, Improvement Value, Miscellaneous Value |
| Parties | Grantor, Grantee |
| Deed | Deed Book, Deed Page |
| Review | Comments, Type |

The original vendor columns can remain available while standardized fields are created for the analytical engine.

---

## 🔀 Vendor Mapping Layer

The mapping layer runs before the checker engine.

### What it does

- Examines uploaded report headers.
- Identifies the report/vendor structure when possible.
- Applies known vendor mappings.
- Applies approved fallback aliases when necessary.
- Creates standardized fields for the checker engine.
- Reports how each application field was mapped.
- Allows the user to override a proposed field when more than one source field could qualify.

### Why the Mapping Layer Matters

Different vendors may use headers such as:

```text
APPRVAL
APPRAISED_VALUE
TOTAL VALUE
TOTAL VALUE OF PROPERTY
```

These can represent the same underlying concept. Rather than teaching every checker every possible vendor header, the mapping layer translates recognized source fields into one common structure.

### Collapsible Field Mapping Review

The **Field Mapping Review is collapsed by default**. Click its arrow to:

1. Review the proposed source field for each standardized field.
2. Select a different source field when appropriate.
3. Click **Apply Field Mapping**.
4. Expand **Mapping Details** to verify the final result.

The user's selection becomes authoritative for the analytical engine.

### Delta Use Code Rule

For Delta reports, the `USE CODE` header is the authoritative Use Code field. Other Delta code columns, including improvement-code fields, are not substitutes for `USE CODE`.

Where configured, common four-digit representations are normalized. Examples include:

```text
0100 → 100
1000 → 100
0101 → 101
1010 → 101
1001 → 1001
```

### Sales Ratio Mapping

When a usable Sales Ratio is supplied by the source report, the application can use the mapped ratio field.

When one is not available, the workflow can derive a ratio from mapped values where configured:

```text
Sales Ratio = Total Appraised Value ÷ Total Sales Price
```

A missing or zero sale price is not used as the denominator.

> The **Ratio Checker** can subsequently perform its own type-aware analytical ratio processing.

---

## 🧰 Application Tools

### 1. 🧰 Prepare Data

Creates a clean, standardized working dataset before analysis.

Preparation may include:

- Removing blank rows and empty columns.
- Cleaning prior generated diagnostic fields.
- Consolidating supported date structures.
- Preparing Use Code information.
- Preserving parcel and identifier fields as text.
- Removing prior checker status fields so the next diagnostic can begin cleanly.

> **Recommended:** Run **Prepare Data** after mapping changes and between diagnostic tools when you want a clean working dataset.

---

### 2. 📅 Sale Date Checker

Reviews sale dates against the selected tax-year study period.

For a selected tax year:

```text
October 1 of Tax Year - 2
through
September 30 of Tax Year - 1
```

It can identify:

- Valid active sales.
- Missing sale dates.
- Invalid dates.
- Sales outside the study period.
- Supported split Year / Month / Day structures.

---

### 3. 🏷️ Use Code Checker

Reviews Use Code and related valuation information.

Potential review conditions include:

- Missing Use Code.
- Unsupported or unrecognized Use Code.
- Missing or zero Land Value.
- Improvement value associated with a land-only code.
- Manufactured-home review conditions.
- Other valuation / Use Code inconsistencies.

> A flag means **review is needed**; it does not automatically mean the assessment is wrong.

---

### 4. 🏠 Appraisal Value Checker

Compares appraisal components with the stated total appraised value.

```text
Calculated Total
    = Land Value
    + Improvement Value
    + Miscellaneous Value
```

The application compares the calculated total with the mapped stated Total Value using the configured crossfoot tolerance.

---

### 5. 📜 Deed / MH / Comment Audit

Reviews transaction, deed, party, qualification, manufactured-home, and comment information.

Potential review conditions include:

- Missing Grantor or Grantee.
- **Bad sale with no supporting comment.**
- Conflicting or unknown qualification status.
- Missing deed information.
- Missing neighborhood information.
- Same deed appearing in different neighborhoods.
- Possible duplicate transactions.
- Grantor and Grantee being the same party.
- Manufactured-home references.
- LLC-to-LLC transactions.
- Configured review keywords.

A **bad sale with no comment** is its own review condition; it is not limited to records that already contain a bad-sale comment.

---

### 6. 📐 Ratio Checker

Calculates and evaluates sales ratios using type-aware logic.

| Type | Calculation |
|---|---|
| `L` | `Land Value ÷ Sale Price` |
| `B` | `(Improvement Value + Miscellaneous Value) ÷ Sale Price` |
| `L&B` | `Total Value ÷ Sale Price` |

The Ratio Checker can also evaluate:

- Global ratio ranges.
- Neighborhood ratio behavior.
- Neighborhood quartiles.
- Low observations.
- High observations.
- Acceptable observations.
- Global outliers.
- Neighborhood outliers.
- Critical records that are both global and neighborhood outliers.

---

### 7. 🛡️ Quality Checker

Runs a broader integrity review across the sales-ratio dataset.

Checks can include:

- Qualification status.
- Bad-sale comments.
- Parcel Number.
- Neighborhood.
- Sale Date.
- Use Code.
- Land Value.
- Appraisal component-to-total agreement.
- Grantor and Grantee.
- Duplicate transactions.
- Deed / neighborhood conflicts.
- Global ratio outliers.
- Neighborhood ratio outliers.

The resulting **Flag Status** identifies records that appear compliant versus records requiring review.

---

### 8. 📊 Generate Statistics

Creates statistical summaries from the prepared/current analysis data.

| Statistic | Description |
|---|---|
| Total Count | Number of records |
| Good Sales | Qualified / usable sales |
| Bad Sales | Excluded / bad sales |
| Undetermined | Unrecognized qualification |
| Mean | Average sales ratio |
| Median | Median sales ratio |
| Weighted Mean | Total appraisal ÷ total sale price |
| Minimum | Lowest ratio |
| Maximum | Highest ratio |
| Range | Maximum − Minimum |
| PRD | Price-Related Differential |
| COD | Coefficient of Dispersion |
| Appraisal Statistics | Minimum, maximum, and average appraisal |
| Average Sale Price | Average sale price for applicable records |

Additional tables can include:

- **Use Code Statistics**
- **Valuation Zone / Neighborhood Statistics**
- **Overall Study Summary**

> **Troubleshooting:** If generated statistics do not agree with the CAMA system, compare the sale type (`L`, `B`, or `L&B`) with the source-system classification. A type mismatch changes which values are used in the ratio calculation.

---

### 9. 🧽 Clear Results

Removes prior generated checker output while retaining the uploaded working data.

Use this when you want to begin another analysis without re-uploading the source workbook.

---


## ⚙️ Ratio Threshold Settings

The app includes a collapsible **Ratio Threshold Settings** area between **Field Mapping Review** and the **Tool Menu**.

The settings are automatically loaded with the existing checker requirements, so an analyst can simply leave them unchanged.

### Current Overall / Global Defaults

| Setting | Default |
|---|---:|
| Too Low for Global | Ratio `< 0.5000` |
| Perfect Global minimum | `0.7000` |
| Perfect Global maximum | `1.2000` |
| Too High for Global | Ratio `> 1.5000` |

Ratios between the Too Low cutoff and the Perfect range, or between the Perfect range and the Too High cutoff, are classified as **Acceptable Global**.

### Current Neighborhood Default

The existing neighborhood logic is retained as the default:

- **Too Low:** below the neighborhood's 25th percentile (`Q1`)
- **Acceptable:** from `Q1` through `Q3`
- **Too High:** above the neighborhood's 75th percentile (`Q3`)

The analyst can change the lower and upper neighborhood percentiles when needed.

The app also provides an optional **Fixed Ratio Limits** mode. In that mode, the analyst enters one neighborhood low cutoff and one neighborhood high cutoff that apply to every neighborhood.

### Applying or Resetting Settings

- Click **Apply Threshold Settings** to make the selected values authoritative for the current session.
- Click **Reset to Current Code Defaults** to restore the original checker requirements.
- The selected settings are used by the **Ratio Checker**, **Quality Checker**, and **Neighborhood Statistics** table.
- The Dashboard records the active global cutoffs and neighborhood threshold method after the Ratio or Quality Checker runs.

---

## 🏘️ Neighborhood Statistics

When the **Ratio Checker** or **Quality Checker** is run, the application can create a collapsible **Neighborhood Statistics** table.

Depending on available data, it can show:

- Valid Sales
- Threshold Method
- Neighborhood Low Threshold
- Neighborhood High Threshold
- Q1
- Median
- Q3
- Mean
- Minimum
- Maximum
- Neighborhood Low Outliers
- Neighborhood High Outliers
- Global Low Outliers
- Global High Outliers
- Critical Global + Neighborhood Outliers

This gives the analyst neighborhood context rather than treating every unusual ratio as the same kind of exception.

---

## 📊 Dashboard and Results

### Dashboard

After a checker runs, the Dashboard displays checker-specific summary information. Depending on the tool, this may include:

- Total rows reviewed.
- Verified / compliant records.
- Records requiring review.
- Missing-data counts.
- Ratio statistics.
- Median ratio.
- Mean ratio.
- COD.
- Other checker-specific results.

### Analysis Results

The working dataframe is displayed directly in the application.

Use:

> **Show only rows requiring review**

to focus the table on exception records.

### How to Interpret a Flag

A flagged row is a **quality-control prompt**, not an automatic determination that the underlying sale or assessment is incorrect. Review the reason/status fields and compare them with the CAMA record and supporting documentation.

---

## 📥 Excel Export

Use **Download Processed Workbook** to export the current checker results.

The exported workbook may contain:

```text
Analysis
Dashboard
Use Code Statistics
Valuation Zone / Neighborhood Statistics
Other generated statistical tables
```

Diagnostic statuses are visually formatted to assist analyst review.

---

## 📁 Application File Structure

```text
sales-ratio-quality-checker/
│
├── app.py
├── engine.py
├── mapping_layer.py
├── vendor_profiles.json
├── vendor_headers.json
├── requirements.txt
└── README.md
```

| File | Responsibility |
|---|---|
| `app.py` | Streamlit UI, upload workflow, mapping review, tool controls, dashboard, in-app help, results, and Excel download |
| `engine.py` | Checker logic, statistics, workbook processing, and Excel export |
| `mapping_layer.py` | Vendor detection, mapping options, overrides, and field standardization |
| `vendor_profiles.json` | Vendor-specific mapping profiles |
| `vendor_headers.json` | Known source-header information |
| `requirements.txt` | Required Python packages |
| `README.md` | Project and application documentation |

---

## 💻 Running Locally

### 1. Install Requirements

```bash
pip install -r requirements.txt
```

### 2. Start the Application

```bash
streamlit run app.py
```

Streamlit will provide a local browser address.

---

## ☁️ Deploying to Streamlit Community Cloud

1. Create or open the GitHub repository.
2. Add or replace the updated application files.
3. Confirm the required project files are committed.
4. Connect the repository to Streamlit Community Cloud.
5. Select `app.py` as the main application file.
6. Deploy / reboot the application.
7. Test representative reports from the supported CAMA/vendor systems.

> When the GitHub repository is updated, the deployed application can be rebuilt from the updated repository.

---

## ✅ Recommended User Workflow

```text
1. Upload workbook
2. Review detected vendor
3. Open Field Mapping Review
4. Confirm / override mappings
5. Apply Field Mapping
6. Prepare Data
7. Run checker
8. Review Dashboard
9. Filter flagged rows
10. Correct source data when appropriate
11. Run other checker(s) as needed
12. Generate Statistics
13. Export processed workbook
```

---

## 🔎 Mapping Review

Always review mapping information when testing a new report format.

A recognizable header can still represent information differently between vendors. If a new vendor or report layout is introduced, update the mapping configuration rather than spreading vendor-specific assumptions throughout individual checkers.

---

## ➕ Adding a New Vendor

1. Obtain a representative vendor report.
2. Identify source headers and their meanings.
3. Compare them with standardized application fields.
4. Add or update the vendor profile.
5. Add approved aliases where appropriate.
6. Test vendor detection and Mapping Details.
7. Test user mapping overrides.
8. Test Prepare Data.
9. Test every checker that relies on the new fields.
10. Compare results with expected source/VBA behavior.
11. Deploy only after validation.

---

## ⚠️ Important Notes

- The application performs its calculations in **Python**.
- It does **not** execute the original VBA code on the Streamlit server.
- The original VBA workflow was used as the specification for recreating analytical behavior.
- Vendor mapping occurs before diagnostic checker logic.
- New or changed vendor exports should be tested before production use.
- Automated flags are quality-control indicators and should be reviewed by the appropriate analyst when professional judgment is required.

---

## 🎯 Project Goal

The goal of the Sales Ratio Quality Checker is to preserve the useful quality-control workflow of the original Excel/VBA system while making it easier to use across different CAMA report formats through a browser-based application and a centralized vendor mapping layer.

**Built for sales-ratio quality control, standardized CAMA report review, and repeatable analytical workflows.**
