Sales Ratio Quality Checker
A browser-based sales-ratio quality-control and analysis application
built with Python and Streamlit.
The Sales Ratio Quality Checker converts the original Excel/VBA workflow
into a web application that can accept CAMA and sales-ratio reports,
standardize different vendor report formats, run quality-control checks,
calculate sales ratios and statistics, review flagged records, and
export the processed results to Excel.
Purpose
CAMA and sales-ratio exports do not always use the same column names or
report layouts. The application uses a mapping layer to translate
supported vendor fields into a common structure before the analytical
tools run.
The general workflow is:
Upload Report → Detect/Map Fields → Prepare Data → Run Checker →
Review Dashboard and Results → Export Workbook
This allows the checker logic to work from standardized fields instead
of requiring every uploaded report to have identical headers.
Main Features
Upload Excel workbooks directly in the browser.
Supports `.xlsx`, `.xlsm`, `.xls`, and `.xlsb` files.
Reads the `Analysis` worksheet when one exists; otherwise uses the
first worksheet.
Maps supported vendor/CAMA fields into standardized application
fields.
Displays vendor/mapping information so field interpretation can be
reviewed.
Preserves the working data in the browser session.
Runs individual diagnostic tools without requiring Microsoft Excel
or VBA on the server.
Displays checker summaries and detailed Analysis results.
Filters the Analysis table to records requiring review.
Generates statistical tables.
Exports processed results and dashboards to an Excel workbook.
Supported Report Formats
The mapping layer was developed from the supplied CAMA report examples
and VBA workflow. Known report layouts include:
Assurance
Assurance2
Delta
S&W
Capture
Ingeunity
Different report formats can use different names for the same underlying
information. The mapping layer translates recognized source fields into
standardized application fields such as:
Parcel Number
Neighborhood
Sale Price
Total Value
Land Value
Improvement Value
Miscellaneous Value
Use Code
Qualification
Sale Date
Grantor
Grantee
Comments
Deed Book
Deed Page
Sales Ratio
Type
The original vendor columns can remain available while standardized
fields are created for the analytical engine.
Vendor Mapping Layer
The mapping layer runs before the checker engine.
Its job is to:
Examine the uploaded report headers.
Identify the report/vendor structure when possible.
Apply the appropriate known vendor mapping.
Apply approved fallback aliases when necessary.
Create standardized fields for the checker engine.
Report how each application field was mapped.
This separation is important because the checker functions should
analyze the meaning of the data rather than independently guessing which
vendor column represents each field.
Use Code Mapping
Use Code handling supports vendor-specific structures and standardized
aliases.
For Delta reports, the `USE CODE` header is the authoritative Use Code
field. Other Delta code columns, including improvement-code fields, are
not substitutes for the Delta `USE CODE` field.
Sales Ratio Mapping
When a usable Sales Ratio is supplied by the source report, the
application can use the mapped ratio field.
When a source report does not contain a usable Sales Ratio, the mapping
process can derive a standardized ratio from the mapped values:
``` text
Sales Ratio = Total Appraised Value / Total Sales Price
```
A missing or zero sale price is not used as the denominator.
The application's Ratio Checker can subsequently perform its own
type-aware analytical ratio processing as required by the checker
workflow.
Application Tools
Prepare Data
Cleans and normalizes the working Analysis data before diagnostics are
run.
Preparation includes operations such as:
Removing blank rows and empty columns.
Cleaning prior generated diagnostic fields.
Consolidating supported date structures.
Consolidating Use Code information where applicable.
Preserving parcel/identifier fields as text.
Removing prior checker status fields so another diagnostic can start
cleanly.
Use Prepare Data between diagnostic tools when a clean working
dataset is desired.
Sale Date Checker
Reviews sale-date information against the selected tax-year study
period.
For a tax year, the study window is:
``` text
October 1 of Tax Year - 2
through
September 30 of Tax Year - 1
```
Supported date structures include recognized Sale Date fields and
supported split Year/Month/Day fields.
Records can be identified as valid active sales, missing date
information, or invalid/out-of-period dates.
Use Code Checker
Reviews Use Code and related valuation information.
The checker can identify conditions such as:
Missing Use Code.
Unsupported or unrecognized Use Code.
Missing or zero Land Value.
Improvement value associated with a land-only code.
Manufactured-home review conditions.
Other valuation/use-code inconsistencies.
Appraisal Value Checker
Compares appraisal components with the stated total appraisal value.
Conceptually:
``` text
Calculated Total =
Land Value
+ Improvement Value
+ Miscellaneous Improvement Value
```
The calculated total is compared with the mapped stated Total Value.
The application uses a crossfoot tolerance when determining whether the
component total and stated total agree.
Deed / MH / Comment Audit
Reviews transaction, deed, party, qualification, and comment
information.
Potential review conditions include:
Missing Grantor or Grantee.
Bad sale without a supporting comment.
Conflicting or unknown qualification status.
Missing deed or neighborhood information.
Same deed appearing in different neighborhoods.
Possible duplicate transactions.
Grantor and Grantee being the same party.
Manufactured-home references.
LLC-to-LLC transactions.
Configured review keywords.
Ratio Checker
Calculates and evaluates sales ratios.
The checker supports type-aware ratio logic:
``` text
L   = Land Value / Sale Price

B   = (Improvement Value + Miscellaneous Value) / Sale Price

L&B = Total Value / Sale Price
```
The tool also evaluates global ratio ranges and neighborhood behavior.
Neighborhood review uses the distribution of valid ratios within the
neighborhood to identify low, high, or acceptable observations.
Quality Checker
Runs a broader integrity review across the sales-ratio dataset.
Checks can include:
Qualification status.
Bad-sale comments.
Parcel Number.
Neighborhood.
Sale Date.
Use Code.
Land Value.
Appraisal component-to-total agreement.
Grantor and Grantee.
Duplicate transactions.
Deed/neighborhood conflicts.
Global ratio outliers.
Neighborhood ratio outliers.
The resulting `Flag Status` identifies records that are compliant or
require review.
Generate Statistics
Creates statistical summaries from the working dataset.
Outputs can include:
Use Code Statistics.
Valuation Zone / Neighborhood Statistics.
Good-sale counts.
Bad-sale counts.
Undetermined qualification counts.
Mean ratio.
Median ratio.
Weighted mean.
Minimum and maximum ratios.
Range.
PRD.
COD.
Appraisal statistics.
Sale-price statistics.
Overall study summary.
Dashboard
After a tool runs, the application displays a dashboard containing the
summary generated by that checker.
Examples include:
Total rows reviewed.
Compliant records.
Records requiring review.
Missing data counts.
Ratio statistics.
Median ratio.
COD.
Other checker-specific results.
The dashboard changes depending on the most recently executed tool.
Analysis Results
The working dataframe is displayed directly in the application.
Use Show only rows requiring review to filter the table to records
that are not currently classified as verified/perfect results.
This makes it easier to focus on records requiring analyst attention.
Excel Export
The processed dataset can be downloaded as an Excel workbook.
The export can contain:
`Analysis`
`Dashboard`
Additional statistical tables produced by the selected tool
Diagnostic statuses are also visually formatted in the exported workbook
to assist review.
Application File Structure
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
`app.py`
Streamlit user interface.
Responsible for:
File upload.
Session state.
Mapping display.
Tool buttons.
Dashboard presentation.
Analysis table.
Excel download.
`engine.py`
Core analytical engine.
Contains the checker logic, data preparation, statistical calculations,
workbook reading, and Excel export functions.
`mapping_layer.py`
Translates different vendor/CAMA report structures into the standardized
fields used by the engine.
`vendor_profiles.json`
Contains known vendor-specific mapping profiles and rules used by the
mapping layer.
`vendor_headers.json`
Contains known header information used to help recognize supported
report structures.
`requirements.txt`
Lists the Python packages required to run the application.
Running Locally
Install Python and the required packages.
From the project directory:
``` bash
pip install -r requirements.txt
```
Then start the Streamlit application:
``` bash
streamlit run app.py
```
Streamlit will provide a local address for opening the application in a
browser.
Deploying to Streamlit Community Cloud
Create or open the GitHub repository for the application.
Upload the application files to the repository.
Make sure `app.py`, `engine.py`, `mapping_layer.py`, the vendor JSON
files, and `requirements.txt` are committed.
Connect the GitHub repository to Streamlit Community Cloud.
Select `app.py` as the main application file.
Deploy the application.
After deployment, test representative reports from the supported
CAMA/vendor systems.
When the GitHub repository is updated, the deployed application can be
rebuilt from the updated repository.
Recommended User Workflow
Open the Sales Ratio Quality Checker.
Upload the CAMA or sales-ratio workbook.
Review the detected vendor and mapping information.
Confirm that important fields were mapped correctly.
Run Prepare Data.
Select the desired checker.
Review the Dashboard.
Review or filter the Analysis Results.
Correct source data when appropriate.
Run additional checkers as needed.
Generate statistics when appropriate.
Download the processed workbook.
Mapping Review
Always review the mapping information when testing a new report format.
A report can contain a recognizable column name while still representing
information differently from another vendor. The mapping report is
therefore an important quality-control step.
If a new vendor or report layout is introduced, update the mapping
configuration rather than adding vendor-specific assumptions throughout
every checker.
Adding a New Vendor
When adding another report format:
Obtain a representative report from the vendor.
Identify the source headers and their meanings.
Compare them with the standardized application fields.
Add or update the vendor profile.
Add approved header aliases where appropriate.
Test the mapping report.
Test Prepare Data.
Test every checker that relies on the newly mapped fields.
Compare results with the expected source/VBA behavior before
deploying the change.
Important Notes
The web application performs calculations in Python.
It does not execute the original VBA code on the Streamlit
server.
The VBA workflow was used as the specification for recreating the
analytical behavior.
Vendor mapping occurs before the diagnostic checker logic.
New or changed vendor exports should be tested before production
use.
Automated flags are quality-control indicators and should be
reviewed by the appropriate analyst when professional judgment is
required.
Project Goal
The goal of the Sales Ratio Quality Checker is to preserve the useful
quality-control workflow of the original Excel/VBA system while making
it easier to use across different CAMA report formats through a
browser-based application and a centralized vendor mapping layer.

https://sale-ratio-quality-checker.streamlit.app/
