# 📊 Sales Ratio Quality Checker

> **A browser-based sales-ratio quality-control and analysis application
> built with Python and Streamlit.**

The **Sales Ratio Quality Checker** converts the original Excel/VBA
workflow into a web application that can accept CAMA and sales-ratio
reports, standardize different vendor report formats, run
quality-control checks, calculate sales ratios and statistics, review
flagged records, and export processed results to Excel.

------------------------------------------------------------------------

## 🔗 Live Application

**Sales Ratio Quality Checker:**\
https://sale-ratio-quality-checker.streamlit.app/

------------------------------------------------------------------------

## 📑 Table of Contents

-   [Purpose](#-purpose)
-   [Application Workflow](#-application-workflow)
-   [Main Features](#-main-features)
-   [Supported Report Formats](#-supported-report-formats)
-   [Vendor Mapping Layer](#-vendor-mapping-layer)
-   [Application Tools](#-application-tools)
-   [Dashboard and Results](#-dashboard-and-results)
-   [Excel Export](#-excel-export)
-   [Application File Structure](#-application-file-structure)
-   [Running Locally](#-running-locally)
-   [Deploying to Streamlit Community
    Cloud](#-deploying-to-streamlit-community-cloud)
-   [Recommended User Workflow](#-recommended-user-workflow)
-   [Adding a New Vendor](#-adding-a-new-vendor)
-   [Important Notes](#-important-notes)

------------------------------------------------------------------------

## 🎯 Purpose

CAMA and sales-ratio exports do not always use the same column names or
report layouts.

The application solves this problem by using a **vendor mapping layer**
to translate supported vendor fields into a common structure **before**
the analytical tools run.

This allows the checker logic to work from standardized fields instead
of requiring every uploaded report to have identical headers.

------------------------------------------------------------------------

## 🔄 Application Workflow

``` text
Upload Report
      ↓
Detect Vendor / Map Fields
      ↓
Create Standardized Fields
      ↓
Prepare Data
      ↓
Run Selected Checker
      ↓
Review Dashboard & Analysis Results
      ↓
Generate Statistics
      ↓
Export Processed Workbook
```

------------------------------------------------------------------------

## ✨ Main Features

-   📁 Upload Excel workbooks directly in the browser.
-   📊 Supports `.xlsx`, `.xlsm`, `.xls`, and `.xlsb` files.
-   🔎 Reads the `Analysis` worksheet when one exists; otherwise uses
    the first worksheet.
-   🔀 Maps supported vendor/CAMA fields into standardized application
    fields.
-   🧭 Displays vendor and mapping information for review.
-   💾 Preserves working data during the browser session.
-   🧰 Runs individual diagnostic tools without requiring Microsoft
    Excel or VBA on the server.
-   🚩 Identifies records requiring analyst review.
-   📈 Generates sales-ratio statistics.
-   🗂️ Creates Use Code and Valuation Zone/Neighborhood statistical
    tables.
-   📥 Exports processed results and dashboards to Excel.

------------------------------------------------------------------------

## 🏢 Supported Report Formats

The mapping layer was developed from the supplied CAMA report examples
and the original VBA workflow.

### Known Vendor / Report Layouts

  Vendor / Format    Supported
  ----------------- -----------
  Assurance             ✅
  Delta                 ✅
  S&W                   ✅
  Capture               ✅
  Ingeunity             ✅

> **Note:** Vendor exports may use different names for the same
> underlying information. The mapping layer translates recognized source
> headers into standardized application fields.

### Standardized Application Fields

  -----------------------------------------------------------------------
  Category                            Standardized Fields
  ----------------------------------- -----------------------------------
  Property                            Parcel Number, Neighborhood, Use
                                      Code

  Sale                                Sale Price, Sale Date, Sales Ratio,
                                      Qualification

  Appraisal                           Total Value, Land Value,
                                      Improvement Value, Miscellaneous
                                      Value

  Parties                             Grantor, Grantee

  Deed                                Deed Book, Deed Page

  Review                              Comments, Type
  -----------------------------------------------------------------------

The original vendor columns can remain available while standardized
fields are created for the analytical engine.

------------------------------------------------------------------------

## 🔀 Vendor Mapping Layer

The **mapping layer runs before the checker engine**.

Its responsibilities are to:

1.  Examine the uploaded report headers.
2.  Identify the report/vendor structure when possible.
3.  Apply the appropriate known vendor mapping.
4.  Apply approved fallback aliases when necessary.
5.  Create standardized fields for the checker engine.
6.  Report how each application field was mapped.

### Why the Mapping Layer Matters

Without the mapping layer, each checker would have to independently
determine whether fields such as:

``` text
APPRVAL
APPRAISED_VALUE
TOTAL VALUE
TOTAL VALUE OF PROPERTY
```

all represent the same underlying concept.

Instead, the mapping layer translates recognized source fields into a
common application structure before analysis begins.

### Delta Use Code Rule

For **Delta** reports, the `USE CODE` header is the authoritative Use
Code field.

Other Delta code columns---including improvement-code fields---are **not
substitutes** for the Delta `USE CODE` field.

### Sales Ratio Mapping

When a usable Sales Ratio is supplied by the source report, the
application can use the mapped ratio field.

When a report does **not** contain a usable Sales Ratio, the mapping
process can derive one from the mapped values:

``` text
Sales Ratio = Total Appraised Value ÷ Total Sales Price
```

A missing or zero sale price is not used as the denominator.

> The application's **Ratio Checker** can subsequently perform its own
> type-aware analytical ratio processing.

------------------------------------------------------------------------

# 🧰 Application Tools

## 1. 🧹 Prepare Data

Cleans and normalizes the working Analysis data before diagnostics are
run.

### Preparation may include:

-   Removing blank rows.
-   Removing completely empty columns.
-   Cleaning prior generated diagnostic fields.
-   Consolidating supported date structures.
-   Consolidating Use Code information where applicable.
-   Preserving parcel and identifier fields as text.
-   Removing prior checker status fields so the next diagnostic can
    start cleanly.

> **Recommended:** Use **Prepare Data** between diagnostic tools when
> you want a clean working dataset.

------------------------------------------------------------------------

## 2. 📅 Sale Date Checker

Reviews sale-date information against the selected tax-year study
period.

### Study Period

For a selected tax year:

``` text
October 1 of Tax Year - 2
through
September 30 of Tax Year - 1
```

### Reviews

-   Valid active sales.
-   Missing sale-date information.
-   Invalid dates.
-   Sales outside the study period.
-   Supported split Year / Month / Day structures.

------------------------------------------------------------------------

## 3. 🏷️ Use Code Checker

Reviews Use Code and related valuation information.

### Potential Issues

-   Missing Use Code.
-   Unsupported or unrecognized Use Code.
-   Missing or zero Land Value.
-   Improvement value associated with a land-only code.
-   Manufactured-home review conditions.
-   Other valuation/use-code inconsistencies.

------------------------------------------------------------------------

## 4. 🏠 Appraisal Value Checker

Compares appraisal components with the stated Total Appraised Value.

### Calculation

``` text
Calculated Total
    = Land Value
    + Improvement Value
    + Miscellaneous Improvement Value
```

The application then compares:

``` text
Calculated Total ↔ Stated Total Value
```

A crossfoot tolerance is used when determining whether the component
total and stated total agree.

------------------------------------------------------------------------

## 5. 📜 Deed / MH / Comment Audit

Reviews transaction, deed, party, qualification, and comment
information.

### Potential Review Conditions

-   Missing Grantor or Grantee.
-   Bad sale without a supporting comment.
-   Conflicting qualification status.
-   Unknown qualification status.
-   Missing deed information.
-   Missing neighborhood information.
-   Same deed appearing in different neighborhoods.
-   Possible duplicate transactions.
-   Grantor and Grantee being the same party.
-   Manufactured-home references.
-   LLC-to-LLC transactions.
-   Configured review keywords.

------------------------------------------------------------------------

## 6. 📐 Ratio Checker

Calculates and evaluates sales ratios.

### Type-Aware Ratio Logic

  Type      Calculation
  --------- ----------------------------------------------------------
  **L**     `Land Value ÷ Sale Price`
  **B**     `(Improvement Value + Miscellaneous Value) ÷ Sale Price`
  **L&B**   `Total Value ÷ Sale Price`

The Ratio Checker also evaluates:

-   Global ratio ranges.
-   Neighborhood ratio behavior.
-   Neighborhood quartiles.
-   Low observations.
-   High observations.
-   Acceptable observations.

------------------------------------------------------------------------

## 7. 🛡️ Quality Checker

Runs a broader integrity review across the sales-ratio dataset.

### Checks Can Include

-   Qualification status.
-   Bad-sale comments.
-   Parcel Number.
-   Neighborhood.
-   Sale Date.
-   Use Code.
-   Land Value.
-   Appraisal component-to-total agreement.
-   Grantor and Grantee.
-   Duplicate transactions.
-   Deed/neighborhood conflicts.
-   Global ratio outliers.
-   Neighborhood ratio outliers.

The resulting **Flag Status** identifies records that are compliant or
require review.

------------------------------------------------------------------------

## 8. 📊 Generate Statistics

Creates statistical summaries from the working dataset.

### Statistical Outputs Can Include

  Statistic              Description
  ---------------------- -------------------------------------------
  Total Count            Number of records
  Good Sales             Qualified/usable sales
  Bad Sales              Excluded/bad sales
  Undetermined           Unrecognized qualification
  Mean                   Average sales ratio
  Median                 Median sales ratio
  Weighted Mean          Total appraisal ÷ total sale price
  Minimum                Lowest ratio
  Maximum                Highest ratio
  Range                  Maximum − Minimum
  PRD                    Price-Related Differential
  COD                    Coefficient of Dispersion
  Appraisal Statistics   Min, max, and average appraisal
  Average Sale Price     Average sale price for applicable records

Additional tables can include:

-   **Use Code Statistics**
-   **Valuation Zone / Neighborhood Statistics**
-   **Overall Study Summary**

------------------------------------------------------------------------

# 📊 Dashboard and Results

## Dashboard

After a tool runs, the application displays a dashboard containing the
summary generated by that checker.

Depending on the selected tool, the dashboard may display:

-   Total rows reviewed.
-   Verified/compliant records.
-   Records requiring review.
-   Missing-data counts.
-   Ratio statistics.
-   Median ratio.
-   Mean ratio.
-   COD.
-   Other checker-specific results.

## Analysis Results

The working dataframe is displayed directly in the application.

Use:

> **Show only rows requiring review**

to focus the table on records that are not currently classified as
verified/perfect results.

------------------------------------------------------------------------

# 📥 Excel Export

The processed dataset can be downloaded as an Excel workbook.

### Exported Workbook May Contain

``` text
Analysis
Dashboard
Use Code Statistics
Valuation Zone Statistics
Other generated statistical tables
```

Diagnostic statuses are visually formatted in the exported workbook to
assist analyst review.

------------------------------------------------------------------------

# 📁 Application File Structure

A typical deployment contains:

``` text
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

### File Responsibilities

  -----------------------------------------------------------------------
  File                                Purpose
  ----------------------------------- -----------------------------------
  `app.py`                            Streamlit user interface and
                                      application workflow

  `engine.py`                         Checker logic, statistics, workbook
                                      processing, and Excel export

  `mapping_layer.py`                  Vendor detection and field
                                      standardization

  `vendor_profiles.json`              Vendor-specific mapping profiles

  `vendor_headers.json`               Known source-header information

  `requirements.txt`                  Required Python packages

  `README.md`                         Application documentation
  -----------------------------------------------------------------------

------------------------------------------------------------------------

# 💻 Running Locally

## 1. Install Requirements

From the project directory:

``` bash
pip install -r requirements.txt
```

## 2. Start the Application

``` bash
streamlit run app.py
```

Streamlit will provide a local browser address for opening the
application.

------------------------------------------------------------------------

# ☁️ Deploying to Streamlit Community Cloud

1.  Create or open the GitHub repository for the application.
2.  Upload the application files.
3.  Make sure the following files are committed:

``` text
app.py
engine.py
mapping_layer.py
vendor_profiles.json
vendor_headers.json
requirements.txt
README.md
```

4.  Connect the GitHub repository to Streamlit Community Cloud.
5.  Select `app.py` as the main application file.
6.  Deploy the application.
7.  Test representative reports from the supported CAMA/vendor systems.

> When the GitHub repository is updated, the deployed application can be
> rebuilt from the updated repository.

------------------------------------------------------------------------

# ✅ Recommended User Workflow

1.  **Open** the Sales Ratio Quality Checker.
2.  **Upload** the CAMA or sales-ratio workbook.
3.  **Review** the detected vendor and mapping information.
4.  **Confirm** that important fields were mapped correctly.
5.  **Run Prepare Data.**
6.  **Select** the desired checker.
7.  **Review** the Dashboard.
8.  **Review or filter** the Analysis Results.
9.  **Correct source data** when appropriate.
10. **Run additional checkers** as needed.
11. **Generate statistics** when appropriate.
12. **Download** the processed workbook.

------------------------------------------------------------------------

# 🔎 Mapping Review

Always review the mapping information when testing a new report format.

A report can contain a recognizable column name while still representing
information differently from another vendor. The mapping report is
therefore an important quality-control step.

If a new vendor or report layout is introduced, update the **mapping
configuration** rather than adding vendor-specific assumptions
throughout every checker.

------------------------------------------------------------------------

# ➕ Adding a New Vendor

When adding another report format:

1.  Obtain a representative vendor report.
2.  Identify the source headers and their meanings.
3.  Compare the source fields with the standardized application fields.
4.  Add or update the vendor profile.
5.  Add approved header aliases where appropriate.
6.  Test the Mapping Report.
7.  Test **Prepare Data**.
8.  Test every checker that relies on the newly mapped fields.
9.  Compare results with the expected VBA/source behavior.
10. Deploy only after validation is complete.

------------------------------------------------------------------------

# ⚠️ Important Notes

-   The web application performs calculations in **Python**.
-   It does **not** execute the original VBA code on the Streamlit
    server.
-   The original VBA workflow was used as the specification for
    recreating analytical behavior.
-   Vendor mapping occurs **before** diagnostic checker logic.
-   New or changed vendor exports should be tested before production
    use.
-   Automated flags are quality-control indicators and should be
    reviewed by the appropriate analyst when professional judgment is
    required.

------------------------------------------------------------------------

# 🎯 Project Goal

The goal of the **Sales Ratio Quality Checker** is to preserve the
useful quality-control workflow of the original Excel/VBA system while
making it easier to use across different CAMA report formats through a
browser-based application and a centralized vendor mapping layer.

------------------------------------------------------------------------

**Built for sales-ratio quality control, standardized CAMA report
review, and repeatable analytical workflows.**
