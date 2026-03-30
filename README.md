# ⚙️ Ops Automation Dashboard

A multi-page internal Streamlit application built to automate order fulfillment, vendor portal workflows, and business intelligence reporting across Amazon, Walmart, and Snap-on supply chains.

Built and maintained as an internal tool used daily by a team of 5–10 people across operations, sales, and logistics.

---

## What It Does

Manual order management across multiple vendor portals is slow, error-prone, and hard to hand off. This dashboard replaces a collection of standalone scripts with a unified, team-facing application that non-technical staff can operate without developer involvement.

**Time saved: ~5–10 hours per week across the team.**

---

## Features

### 📦 Amazon Vendor Central
- **Backorder Pipeline** — Multi-stage Streamlit workflow that logs into Amazon Vendor Central via Selenium, downloads backorder and substitute inventory reports, processes them against warehouse stock levels, and routes each item to one of several decisions: fulfill from committed stock, transfer from an alternate warehouse, substitute with another SKU, or reject. Generates a color-coded HTML email table and opens a pre-populated Outlook reply draft.
- **Order Paperwork Generator** — Exports pack worksheets from NetSuite, validates items against a cross-reference file (prompting for missing BPN/UPC data inline), splits the combined export into per-PO files, generates pack sheets and carton labels as XLSX/PDF, and assembles a zip attachment for an Outlook email draft.
- **Weekly Sales Report** — Pulls WTD and YTD sales metrics from Amazon Vendor Central's retail analytics dashboard, merges with NetSuite COGS data, generates YoY revenue line charts and category pie charts using matplotlib, and builds a formatted HTML email with embedded images for management distribution.

### 🛒 Walmart Vendor Portal
- **Order Confirmation Pipeline** — Pre-confirms transfer availability, confirms orders through the Walmart portal, and generates routed order upload files with warehouse priority logic (Overflow → PDX HQ → 3PL PDX).
- **Pack Slip & Carton Label Generator** — Produces warehouse-ready pack slip and carton label files, including LTL/TL load documents and pallet labels.
- **Walmart CA** — Parallel pipeline for Canadian Walmart orders with region-specific formatting.

### 🔧 Snap-on SupplyWeb
- **Pack Slip Generator** — Generates properly formatted pack slips matching Snap-on's specific document requirements.

### 📊 Analytics & Reporting
- **Inventory Health Report** — Queries NetSuite live via SuiteQL (TBA auth), calculates weeks-of-stock per SKU, flags at-risk items, and exports a formatted report across all accounts.

---

## Tech Stack

| Layer | Tools |
|---|---|
| Frontend | Streamlit (multi-page) |
| Automation | Selenium, webdriver-manager |
| Data | Pandas, NumPy |
| Reporting | matplotlib, openpyxl, BeautifulSoup |
| ERP Integration | NetSuite SuiteQL via Token-Based Authentication (TBA) |
| Email | win32com (Outlook), threading |
| Credentials | Streamlit secrets.toml |
| AI-Assisted Dev | Claude (Anthropic) used throughout for code generation and iteration |

---

## Architecture

```
main.py                  # App entry point, navigation, home page
pages/
  Amazon/
    backorder_pipeline.py        # Multi-stage backorder workflow
    generate_paperwork.py        # Pack sheet + label automation
    sales_report.py              # Weekly sales reporting pipeline
    backorder_upload.py          # Upload file generator
  Walmart/
    pre_confirm_transfer.py
    confirm_orders.py
    pack_slip.py
    carton_labels.py
    ltl_tl_documents.py
  WalmartCA/
    process_orders.py
    add_ship_data.py
  Snapon/
    pack_slip_generator.py
  Analytics/
    inventory_health_report.py
Utilities/
  amazon_login.py        # Selenium login helpers
  netsuite_login.py      # NetSuite TBA auth helpers
  netsuite_utils.py      # SuiteQL query wrappers
```

---

## Notable Patterns

**Stage-based UI state machine** — Long-running pipelines (backorder processing, paperwork generation) use `st.session_state` to drive a multi-stage workflow: idle → running → decisions needed → ready → complete. This lets users make mid-pipeline decisions (e.g. choosing transfer vs. substitute) without losing progress.

**Background threading with queue-based status updates** — Selenium automation and file processing run in daemon threads. A `queue.Queue` passes status updates back to the Streamlit UI, which polls and reruns to show live progress cards and a scrollable activity log.

**Inline cross-reference editing** — When the paperwork generator encounters items missing from the cross-reference file, it surfaces an inline form in the UI for the user to enter BPN/UPC data, then writes the update back to the CSV and continues — no separate file editing required.

**HTML email assembly** — Email bodies are built programmatically using BeautifulSoup, with color-coded table cells, embedded transfer summaries, and a color key legend. Styled for Outlook rendering quirks.

---

## Setup

> ⚠️ This repo contains sanitized demo code. Credentials, internal paths, and proprietary data have been removed. The app requires a live NetSuite account, Amazon Vendor Central access, and Walmart vendor portal credentials to run.

```bash
pip install streamlit pandas selenium webdriver-manager beautifulsoup4 openpyxl pyotp pywin32
streamlit run main.py
```

Credentials go in `.streamlit/secrets.toml` (see `secrets.toml.example` for required keys).

---

## Background

This dashboard grew out of a series of standalone Python scripts written to automate repetitive order management tasks. Over time the scripts were consolidated into a unified Streamlit app, making the tooling accessible to non-technical colleagues across the team without requiring them to run Python directly.

Development has been ongoing since 2019, with AI-assisted development (Claude) adopted more recently to accelerate feature iteration and script generation.

---

*Built by Rich Hardisty · [linkedin.com/in/richhardisty](https://linkedin.com/in/richhardisty)*
