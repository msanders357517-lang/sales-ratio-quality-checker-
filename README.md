Sales Ratio Quality Checker — Web App
A browser-based conversion of the Excel/VBA Sales Ratio Quality Checker workflow supplied by the project owner.

Included tools
Prepare Data
Sale Date Checker
Use Code Checker
Appraisal Value Checker
Deed / Manufactured Home / Bad Sale Comment Audit
Ratio Checker
Quality Checker
Generate Statistics
Clear Results / Clear All
Interactive Analysis table and dashboard
Download processed results as `.xlsx`
The app accepts `.xlsx`, `.xlsm`, `.xls`, and `.xlsb` uploads. It searches for an `Analysis` worksheet first; if one does not exist, it uses the first worksheet.

Run locally
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
# source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```
Streamlit will print a local browser address, normally `http://localhost:8501`.

Put it online
Streamlit Community Cloud
Create a GitHub repository.
Upload `app.py`, `engine.py`, and `requirements.txt`.
In Streamlit Community Cloud, choose Create app and select the repository.
Set the main file to `app.py`.
Deploy.

Internal / agency deployment
For private CAMA or taxpayer-related data, use an environment approved by your organization rather than a public file-processing service. The project can be hosted behind authentication on Azure App Service, AWS, Google Cloud, an internal server, or another approved platform.

Architecture
The browser UI is Streamlit. Excel/VBA logic is executed in Python/pandas, so the server does not need Microsoft Excel. Uploaded data remains in the app session unless your hosting platform is configured to persist it.

Source-rule notes
This first conversion implements the supplied VBA rules for the primary menu tools, including:
Tax-year sale-date window (`Oct 1, tax year - 2` through `Sep 30, tax year - 1`).
Approved, land-only, and manufactured-home use-code lists.
Land-only valuation conflict checks.
Appraisal cross-foot tolerance of one cent.
Type-aware ratios (`L`, `B`, `L&B`).
Global ratio thresholds and neighborhood quartile checks.
Deed/comment audit priority checks and configured review-keyword families.
Good-sale statistical calculations including median, weighted mean, PRD, and COD.
County-level median/COD review threshold logic.

The original VBA package is very large and contains extensive Excel-specific presentation, worksheet-navigation, formula-protection, drill-down, comment, chart, and reconciliation behavior. Those Excel-only interface behaviors are intentionally represented as web tables/dashboard interactions rather than VBA objects.


URL
https://sale-ratio-quality-checker.streamlit.app/

