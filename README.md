# Vendor Cost Reconciliation Toolkit

A Google Sheets + Apps Script system that ingests vendor usage exports and invoice PDFs, 
attributes transactions to cost centers, reconciles totals against invoices, and generates 
per-cost-center reports ready for downstream expense-management systems.

## What it does (generic terms)

- Reads periodic vendor usage files (spreadsheets) from a Drive inbox folder
- Deduplicates transactions by transaction ID across cumulative uploads
- Parses invoice PDFs via an LLM API to extract billing metadata
- Reconciles ledger sums against invoice totals per period + vertical
- Auto-attributes transactions to canonical cost centers via alias mapping
- Handles multi-cost-center assignments with automatic even-split distribution
- Generates a per-cost-center report grouped by invoice for downstream submission

## Setup

1. Copy the Apps Script files into your Google Sheets project
   (Extensions → Apps Script → paste each `.gs` file)
2. Enable the Advanced Drive Service:
   Services → `+` → Drive API → v2 → Add
3. Reload the spreadsheet — an **⚙️ Setup** menu appears
4. Run **⚙️ Setup → Configure keys, IDs, and recipients**:
   - LLM API key (used for invoice PDF parsing)
   - Notification email address
   - Parent Drive folder ID (where the system will create its inbox/archive folders)
5. Run **⚙️ Setup → Initialize folder structure** (or the equivalent menu item in your fork)
6. Run **⚙️ Setup → Initialize master sheets**
7. Populate `_master_cost_centers` with your own cost centers and eligibility flags

## Caveat on data-source specifics

Some data-parsing logic in this template is deliberately simplified or presented as 
example schemas. The author works with confidential HR data and cannot publish the 
production import mappings or exact vendor column layouts. If you adapt this repo, 
you will need to:

- Replace the demonstration ledger schema with your real vendor export columns
- Update the header maps in the importer to match your vendors' column names
- Adjust the invoice-parser prompt to match your invoice formats

The reconciliation logic, attribution engine, alias handling, multi-cost-center 
distribution, and reporting layer are complete and production-ready.

## Attribution

100% AI-generated, 100% human-directed. Directed and iterated by 
Yordhan Fitrians Akhmad B. based on real HR pain points — refined repeatedly for 
business fit, user experience, and workflow comfort.
