# ──────────────────────────────────────────────────────────────────────────────
# DEMO VERSION
# Sanitized for portfolio use. Credentials, internal file paths, and email
# addresses have been replaced with placeholders. Requires a live NetSuite
# account and a configured secrets.toml to run. See README.md for context.
# ──────────────────────────────────────────────────────────────────────────────

import streamlit as st
import time
import re
import traceback
import pyotp
import shutil
import pythoncom
import os
import copy
import threading
import csv
import io
from datetime import datetime, timedelta, date
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import win32com.client

# ──────────────────────────────────────────────
# Credentials from secrets.toml
# ──────────────────────────────────────────────
os.environ['LONG_EMAIL']        = st.secrets['LONG_EMAIL']
os.environ['NETSUITE_PASSWORD'] = st.secrets['NETSUITE_PASSWORD']
os.environ['NETSUITE_KEY']      = st.secrets['NETSUITE_KEY']

# ──────────────────────────────────────────────
# Session state initialisation
# ──────────────────────────────────────────────
for key, default in {
    'phase': 'ready',          # ready | running | processing | done | error
    'modified_table': None,
    'slim_table': None,
    'transfers_by_location': {},
    'total_orders': 0,
    'log_details': [],
    'error_message': None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ──────────────────────────────────────────────
# Date helpers
# ──────────────────────────────────────────────
current_week_monday = datetime.now() - timedelta(days=datetime.now().weekday())
formatted_date      = current_week_monday.strftime("%d-%b")
formatted_monday    = current_week_monday.strftime("%Y-%m-%d")
folder_path_today   = datetime.now().strftime("%Y-%m-%d")

def get_timestamp():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def is_valid_date(date_str):
    if not date_str or str(date_str).lower() in ('none', 'nat', 'nan', ''):
        return None
    for fmt in ["%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"]:
        try:
            return datetime.strptime(str(date_str).strip(), fmt).date()
        except ValueError:
            continue
    return None

def normalize_date(date_str):
    d = is_valid_date(date_str)
    return d.strftime("%m/%d/%Y") if d else None

def clean_text(text):
    return text.strip().lower() if text else ""

# ──────────────────────────────────────────────
# Logging
# ──────────────────────────────────────────────
downloads_directory = st.secrets['DOWNLOAD_FOLDER']
bo_file_prefix      = st.secrets['WALMART_BO_FILE_PREFIX']
output_path         = st.secrets['WALMART_BO_REPORT_GROUPED']
bo_dstDir           = st.secrets['WALMART_BO_REPORT_FOLDER']
bo_dstFileName      = "Walmart Backorder Item List.csv"

WALMART_BASE_PROJECTS_FOLDER = st.secrets.get(
    'WALMART_BASE_PROJECTS_FOLDER',
    st.secrets.get("WALMART_BASE_PROJECTS_FOLDER", "./projects")
)
WALMART_WEEKLY_FOLDER  = os.path.join(WALMART_BASE_PROJECTS_FOLDER, f"Walmart Orders - Week of {formatted_monday}")
WALMART_AUTOMATION_LOG = os.path.join(WALMART_WEEKLY_FOLDER, r"Uploads\Automation Logs")

os.makedirs(WALMART_AUTOMATION_LOG, exist_ok=True)
os.makedirs(bo_dstDir, exist_ok=True)

log_file_path = os.path.join(
    WALMART_AUTOMATION_LOG,
    f"walmart_backorder_app - {get_timestamp().replace(':', '-')}.txt"
)

def append_log(message):
    entry = f"{get_timestamp()} - {message}"
    st.session_state.log_details.append(entry)

def save_log():
    with open(log_file_path, "w", encoding='utf-8') as f:
        for line in st.session_state.log_details:
            f.write(line + "\n")

# ──────────────────────────────────────────────
# Selenium helpers
# ──────────────────────────────────────────────
def wait_for_element_to_be_clickable(driver, by, value, timeout=30):
    try:
        return WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((by, value))
        )
    except Exception as e:
        append_log(f"Timeout waiting for element ({value}): {e}")
        return None

def wait_for_file_to_appear(directory, prefix, timeout=300):
    start = time.time()
    while True:
        for fn in os.listdir(directory):
            if fn.startswith(prefix) and not fn.endswith('.crdownload'):
                return os.path.join(directory, fn)
        if time.time() - start > timeout:
            raise Exception(f"Timeout: no file with prefix '{prefix}' found in {timeout}s")
        time.sleep(1)

def read_csv_with_encoding(path):
    try:
        df = pd.read_csv(path, encoding='utf-8')
    except UnicodeDecodeError:
        df = pd.read_csv(path, encoding='ISO-8859-1')
    df.columns = df.columns.str.strip()
    return df

# ──────────────────────────────────────────────
# Outlook helpers
# ──────────────────────────────────────────────
def get_outlook_app():
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
        append_log("Attached to existing Outlook instance")
        return outlook
    except:
        outlook = win32com.client.Dispatch("Outlook.Application")
        append_log("Created new Outlook instance")
        return outlook

def create_outlook_draft(styled_body: str, to: str, cc: str, subject: str):
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.HTMLBody = styled_body
        mail.To      = to
        mail.CC      = cc
        mail.Subject = subject
        mail.Display()
        append_log("Outlook draft opened")
    except Exception as e:
        append_log(f"Outlook thread error: {e}")
    finally:
        pythoncom.CoUninitialize()

# ──────────────────────────────────────────────
# Phase 1 — NetSuite download + grouping
# ──────────────────────────────────────────────
def run_phase1():
    """Download, group, and save the Walmart backorder CSV."""
    driver = None
    try:
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": downloads_directory
        })
        driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=chrome_options
        )

        netsuite_url = "https://4423908.app.netsuite.com/app/common/search/searchresults.nl?searchid=11060"
        driver.get(netsuite_url)
        append_log("Navigated to NetSuite login")

        driver.find_element(By.ID, "email").send_keys(os.environ['LONG_EMAIL'])
        driver.find_element(By.ID, "password").send_keys(os.environ['NETSUITE_PASSWORD'])
        driver.find_element(By.ID, "login-submit").click()
        append_log("Submitted login credentials")
        time.sleep(5)

        try:
            choose_role = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.LINK_TEXT, "Choose Role"))
            )
            choose_role.click()
            append_log("Clicked 'Choose Role'")
        except:
            append_log("No 'Choose Role' link found")

        try:
            key  = os.environ['NETSUITE_KEY']
            totp = pyotp.TOTP(key)
            two_fa = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "uif60_input"))
            )
            two_fa.send_keys(totp.now())
            driver.find_element(By.ID, "uif76").click()
            append_log("Entered 2FA code")
            time.sleep(3)
        except Exception as e:
            append_log(f"2FA exception: {e}")
            raise

        # Export CSV
        export_btn = wait_for_element_to_be_clickable(
            driver, By.CSS_SELECTOR,
            'div.uir-list-icon-button.uir-list-export-csv[title="Export - CSV"]'
        )
        if not export_btn:
            raise Exception("Export CSV button not found")
        export_btn.click()
        append_log("Clicked Export CSV button")

        file_path = wait_for_file_to_appear(downloads_directory, bo_file_prefix)
        append_log(f"File downloaded: {file_path}")

        # Group by Item + SO, then summarise by Item
        df = read_csv_with_encoding(file_path)

        merged_df = df.groupby(['Item', 'SO']).agg({
            'Description':               'first',
            'Ship From':                 'first',
            'Quantity':                  'first',
            'Quantity Committed':        'first',
            'Latest Delivery Date':      'first',
            'Item Status':               'first',
            'PDX HQ':                    'sum',
            'PDX HQ 2':                  'sum',
            'Overflow':                  'sum',
            '3PL PDX':                   'sum',
            'In Transit':                'sum',
            'Next Restock Date':         'first',
            'SALES - ACTION REQUESTED?': 'first',
        }).reset_index()

        final_df = merged_df.groupby('Item').agg({
            'Description':               'first',
            'Ship From':                 'first',
            'SO':                        'nunique',
            'Quantity':                  'sum',
            'Quantity Committed':        'sum',
            'Latest Delivery Date':      'min',
            'Item Status':               'first',
            'PDX HQ':                    'first',
            'PDX HQ 2':                  'first',
            'Overflow':                  'first',
            '3PL PDX':                   'first',
            'In Transit':                'first',
            'Next Restock Date':         'min',
            'SALES - ACTION REQUESTED?': 'first',
        }).reset_index()

        final_df = final_df.rename(columns={'SO': '# of Orders'})
        final_df['Quantity Committed'] = pd.to_numeric(final_df['Quantity Committed'], errors='coerce').fillna(0).astype(int)
        final_df['Quantity']           = pd.to_numeric(final_df['Quantity'],           errors='coerce').fillna(0).astype(int)
        final_df['Quantity Required']  = (final_df['Quantity'] - final_df['Quantity Committed']).clip(lower=0)

        numeric_cols = ['PDX HQ', 'PDX HQ 2', 'Overflow', '3PL PDX', 'In Transit']
        for col in numeric_cols:
            if col in final_df.columns:
                final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0).astype(int)

        final_df['Next Restock Date'] = (
            final_df['Next Restock Date'].astype(str)
            .replace(['NaT', 'nan', 'None', ''], 'None')
            .fillna('None')
        )
        final_df['Item'] = final_df['Item'].astype(str).str.strip()
        final_df = final_df[final_df['Item'].str.lower() != 'total']
        final_df = final_df.sort_values(by=['Ship From', 'Item'], ascending=[True, True])

        column_order = [
            'Item', '# of Orders', 'Quantity', 'Quantity Committed', 'Quantity Required',
            'Ship From', 'Description', 'Latest Delivery Date', 'Item Status',
            'PDX HQ', 'PDX HQ 2', 'Overflow', '3PL PDX',
            'In Transit', 'Next Restock Date',
            'SALES - ACTION REQUESTED?',
        ]
        column_order = [c for c in column_order if c in final_df.columns]
        final_df = final_df[column_order]

        final_df.to_csv(file_path, index=False)
        shutil.move(file_path, os.path.join(bo_dstDir, bo_dstFileName))
        final_df.to_csv(output_path, index=False)
        append_log(f"CSV saved: {output_path}")

    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass
        append_log("Driver closed")

# ──────────────────────────────────────────────
# Phase 2 — Process CSV into styled HTML table
# ──────────────────────────────────────────────
def get_transfer_location(pdx_hq, pdx_hq_2, overflow, three_pl_pdx, qty_required):
    """Returns (location_label, action_note, available_stock) for the best transfer source,
    or (None, None, 0) if no stock exists. Caller checks the 75% threshold."""

    # Priority 1: Overflow (digital transfer - physically at 3PL PDX, minimises touches)
    if overflow > 0:
        return "3PL Overflow", "Accept - Transfer from Overflow", overflow

    # Priority 2: PDX HQ / PDX HQ 2 - prefer whichever can cover the full qty alone;
    #             if neither can, pick the larger to minimise stops
    if pdx_hq > 0 or pdx_hq_2 > 0:
        hq_covers  = pdx_hq  >= qty_required
        hq2_covers = pdx_hq_2 >= qty_required

        if hq_covers and not hq2_covers:
            return "PDX HQ",   "Accept - Transfer from PDX HQ",   pdx_hq
        elif hq2_covers and not hq_covers:
            return "PDX HQ 2", "Accept - Transfer from PDX HQ 2", pdx_hq_2
        else:
            # Both cover, or neither covers - pick larger to minimise stops
            if pdx_hq >= pdx_hq_2:
                return "PDX HQ",   "Accept - Transfer from PDX HQ",   pdx_hq
            else:
                return "PDX HQ 2", "Accept - Transfer from PDX HQ 2", pdx_hq_2

    # Priority 3: 3PL PDX
    if three_pl_pdx > 0:
        return "3PL - Portland Expeditors", "Accept - Transfer from 3PL PDX", three_pl_pdx

    return None, None, 0

def run_phase2():
    """Read the saved CSV and build a styled BeautifulSoup table."""
    df = read_csv_with_encoding(os.path.join(bo_dstDir, bo_dstFileName))

    transfers_by_location = {
        "3PL Overflow": [],
        "PDX HQ": [],
        "PDX HQ 2": [],
        "3PL - Portland Expeditors": [],
    }
    total_orders = 0

    # Drop YYZ Classic before building table
    if "YYZ Classic" in df.columns:
        df = df.drop(columns=["YYZ Classic"])

    # Build HTML table from DataFrame
    table_html = df.to_html(index=False, border=1, classes="backorder-table")
    soup = BeautifulSoup(table_html, "html.parser")
    table = soup.find("table")

    rows = table.find_all("tr")
    headers = [clean_text(cell.get_text(separator=" ")) for cell in rows[0].find_all(["th", "td"])]
    append_log(f"Processing columns: {headers}")

    # Column index lookups
    def idx(name):
        try:
            return headers.index(name)
        except ValueError:
            return None

    item_index          = idx("item")
    qty_index           = idx("quantity")
    qty_committed_index = idx("quantity committed")
    qty_required_index  = idx("quantity required")
    ship_from_index     = idx("ship from")
    pdx_hq_index        = idx("pdx hq")
    pdx_hq_2_index      = idx("pdx hq 2")
    overflow_index      = idx("overflow")
    three_pl_pdx_index  = idx("3pl pdx")
    restock_index       = idx("next restock date")
    orders_index        = idx("# of orders")
    action_index        = idx("sales - action requested?")
    item_status_index   = idx("item status") or idx("status")

    if any(v is None for v in [item_index, qty_required_index, ship_from_index,
                                pdx_hq_index, pdx_hq_2_index, overflow_index,
                                three_pl_pdx_index, restock_index, orders_index,
                                action_index]):
        raise Exception(f"Missing required columns. Found: {headers}")

    # Style header row
    for th in rows[0].find_all(["th", "td"]):
        th['style'] = (
            'background-color: #1f497d; color: white; font-weight: bold; '
            'font-family: Helvetica, Arial, sans-serif; font-size: 11pt; '
            'padding: 6px 10px; text-align: center;'
        )

    for row in rows[1:]:
        cells = row.find_all("td")
        if len(cells) < len(headers):
            continue

        def cell_val(i, as_int=False):
            v = cells[i].get_text(strip=True) if i is not None and i < len(cells) else ""
            if as_int:
                try:
                    return int(float(v or 0))
                except:
                    return 0
            return v

        item         = cell_val(item_index)
        qty           = cell_val(qty_index,           as_int=True)
        qty_committed = cell_val(qty_committed_index, as_int=True)
        qty_required  = cell_val(qty_required_index,  as_int=True)
        ship_from    = cell_val(ship_from_index).lower()
        pdx_hq       = cell_val(pdx_hq_index, as_int=True)
        pdx_hq_2     = cell_val(pdx_hq_2_index, as_int=True)
        overflow     = cell_val(overflow_index, as_int=True)
        three_pl_pdx = cell_val(three_pl_pdx_index, as_int=True)
        restock_str  = cell_val(restock_index)
        if not restock_str or restock_str.lower() in ('nan', 'nat', 'none', ''):
            restock_str = 'None'
            if restock_index is not None and restock_index < len(cells):
                cells[restock_index].string = 'None'
        orders       = cell_val(orders_index, as_int=True)
        item_status  = cell_val(item_status_index).lower() if item_status_index is not None else ""

        total_orders += orders
        total_pdx_stock = pdx_hq + pdx_hq_2 + overflow + three_pl_pdx
        restock_date    = is_valid_date(restock_str)
        is_portland_hq  = any(v in ship_from for v in ["portland hq", "pdx hq", "portland"])

        # Base row style
        for cell in cells:
            existing = cell.get('style', '')
            cell['style'] = (
                'font-family: Helvetica, Arial, sans-serif; font-size: 11pt; '
                'padding: 4px 8px; ' + existing
            )

        # Determine action
        new_note       = None
        transfer_from  = None
        action_style   = 'text-align: center;'

        if "discontinued" in item_status:
            new_note     = "Reject - Discontinued Item"
            action_style = 'background-color: #FFC7CE; color: #9C0006; font-weight: bold; text-align: center;'

        elif is_portland_hq and total_pdx_stock > 0:
            transfer_loc, transfer_note, avail_stock = get_transfer_location(pdx_hq, pdx_hq_2, overflow, three_pl_pdx, qty_required)
            if transfer_loc:
                threshold = qty_required * 0.75
                if qty_required == 0 or avail_stock >= threshold:
                    new_note      = transfer_note
                    transfer_from = transfer_loc
                    if transfer_loc == "3PL Overflow":
                        action_style = 'background-color: #90EE90; color: #333333; font-weight: bold; text-align: center;'
                    else:
                        action_style = 'background-color: #FFFF99; color: #333333; font-weight: bold; text-align: center;'
                else:
                    new_note     = f"Reject - Insufficient stock for transfer ({avail_stock} available, {qty_required} needed)"
                    action_style = 'background-color: #FFC7CE; color: #9C0006; font-weight: bold; text-align: center;'

        elif total_pdx_stock == 0:
            if "mto" in item_status or "made to order" in item_status:
                new_note     = "Reject - MTO Item"
                action_style = 'background-color: #FFC7CE; color: #9C0006; font-weight: bold; text-align: center;'
            elif not restock_str or restock_str.lower() in ('none', 'nat', 'nan', '', 'none'):
                new_note     = "Reject - OOS - No Restock Date"
                action_style = 'text-align: center;'
            elif restock_date:
                if restock_date <= date.today() + timedelta(days=2):
                    new_note     = f"Reject - OOS until {restock_str}"
                    action_style = 'background-color: #20B2AA; color: black; font-weight: bold; text-align: center;'
                else:
                    new_note     = f"Reject - OOS until {restock_str}"
                    action_style = 'text-align: center;'

        if new_note and action_index is not None and action_index < len(cells):
            cells[action_index].string = new_note
            cells[action_index]['style'] = (
                'font-family: Helvetica, Arial, sans-serif; font-size: 11pt; '
                'padding: 4px 8px; ' + action_style
            )

        if transfer_from:
            transfers_by_location[transfer_from].append(
                f"{item} - Quantity Required: {qty_required}"
            )

    st.session_state.modified_table       = table
    st.session_state.transfers_by_location = transfers_by_location
    st.session_state.total_orders          = total_orders
    append_log(f"Phase 2 complete. {total_orders} total order lines processed.")


# ──────────────────────────────────────────────
# Build slim email table (drop inventory columns)
# ──────────────────────────────────────────────
COLUMNS_TO_DROP = {
    "pdx hq", "pdx hq 2", "overflow", "3pl pdx",
    "yyz classic", "in transit", "next restock date", "available substitutes",
    "quantity", "quantity committed"
}

def build_slim_table(full_table):
    slim = copy.deepcopy(full_table)
    drop_indices = []
    header_row = slim.find("tr")
    if header_row:
        for i, cell in enumerate(header_row.find_all(["th", "td"])):
            if clean_text(cell.get_text(strip=True)) in COLUMNS_TO_DROP:
                drop_indices.append(i)
    for row in slim.find_all("tr"):
        cells = row.find_all(["th", "td"])
        for i in sorted(drop_indices, reverse=True):
            if i < len(cells):
                cells[i].decompose()
    return slim

def table_to_csv(table_soup):
    output = io.StringIO()
    writer = csv.writer(output)
    for row in table_soup.find_all("tr"):
        writer.writerow([c.get_text(strip=True) for c in row.find_all(["th", "td"])])
    return output.getvalue()


# ──────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────
st.title("Walmart Backordered Items — Automation & Reply Generator")

st.markdown("""
This tool runs in two automatic phases:

1. **Download & Group** — Logs into NetSuite, exports the Walmart backorder report, groups it by item, saves a CSV.
2. **Process & Preview** — Reads the CSV, applies transfer/reject logic, builds a styled table ready for the reply email.

Click **Run** to start both phases, then review the table and send the reply draft to Outlook.
""")

user_email = st.text_input(
    "Your email address (for CC)",
    value="you@your-company.com"
)

# ── Run button ──
if st.session_state.phase == 'ready':
    if st.button("▶ Run Automation", type="primary", use_container_width=True, disabled=not user_email):
        st.session_state.phase = 'running'
        st.rerun()

# ── Phase 1 execution ──
if st.session_state.phase == 'running':
    with st.status("Phase 1 — Downloading from NetSuite…", expanded=True) as status:
        try:
            run_phase1()
            status.update(label="Phase 1 complete ✓", state="complete")
        except Exception as e:
            st.session_state.phase = 'error'
            st.session_state.error_message = f"{str(e)}\n{traceback.format_exc()}"
            append_log(st.session_state.error_message)
            save_log()
            status.update(label="Phase 1 failed ✗", state="error")
            st.error(st.session_state.error_message)
            st.stop()

    with st.status("Phase 2 — Processing table…", expanded=True) as status:
        try:
            run_phase2()
            st.session_state.slim_table = build_slim_table(st.session_state.modified_table)
            st.session_state.phase = 'done'
            save_log()
            status.update(label="Phase 2 complete ✓", state="complete")
        except Exception as e:
            st.session_state.phase = 'error'
            st.session_state.error_message = f"{str(e)}\n{traceback.format_exc()}"
            append_log(st.session_state.error_message)
            save_log()
            status.update(label="Phase 2 failed ✗", state="error")
            st.error(st.session_state.error_message)
            st.stop()

    st.rerun()

# ── Error state ──
if st.session_state.phase == 'error':
    st.error(st.session_state.error_message)
    if st.button("Reset", use_container_width=True):
        for key in ['phase', 'modified_table', 'slim_table',
                    'transfers_by_location', 'total_orders',
                    'log_details', 'error_message']:
            del st.session_state[key]
        st.rerun()

# ── Results ──
if st.session_state.phase == 'done':
    st.success("Both phases complete!")

    # ── Transfer summary ──
    st.subheader("Transfer Summary")
    st.write(f"**Total Order Lines:** {st.session_state.total_orders}")
    total_transfers = sum(len(v) for v in st.session_state.transfers_by_location.values())
    if total_transfers:
        for loc, items in st.session_state.transfers_by_location.items():
            if items:
                st.write(f"**Transfer from {loc}:**")
                for entry in items:
                    st.write(f"- {entry}")
    else:
        st.write("No transfers needed this week.")

    # ── Full table preview ──
    st.subheader("Processed Table Preview (full — includes inventory columns)")
    st.markdown(str(st.session_state.modified_table), unsafe_allow_html=True)

    # ── CSV download ──
    st.subheader("Download Slim Table as CSV")
    slim_csv = table_to_csv(st.session_state.slim_table)
    st.download_button(
        label="📥 Download CSV",
        data=slim_csv,
        file_name=f"Walmart_Backorder_{formatted_date.replace('-', '')}.csv",
        mime="text/csv",
        use_container_width=True,
    )

    # ── Generate Outlook draft ──
    st.subheader("Reply Email Draft")
    if st.button("📧 Generate Reply Draft in Outlook", type="primary", use_container_width=True):
        with st.spinner("Opening Outlook draft…"):
            try:
                email_table = st.session_state.slim_table

                # Transfer list HTML
                transfer_html = (
                    f'<p style="font-family: Helvetica, Arial, sans-serif; font-size: 12pt;">'
                    f'<strong>Total Order Lines:</strong> {st.session_state.total_orders}</p>'
                )
                if total_transfers:
                    transfer_html += (
                        '<p style="font-family: Helvetica, Arial, sans-serif; font-size: 12pt; margin-top: 16pt;">'
                        '<strong style="color: #1f497d; font-size: 13pt;">Transfer List (by Location):</strong></p>'
                    )
                    for loc, items in st.session_state.transfers_by_location.items():
                        if items:
                            transfer_html += (
                                f'<div style="margin-left: 30px; margin-bottom: 12pt;">'
                                f'<p style="font-family: Helvetica, Arial, sans-serif; font-size: 12pt; margin: 10pt 0 6pt 0;">'
                                f'<strong style="color: #1f497d;">Transfer from {loc}:</strong></p>'
                                f'<div style="font-family: Helvetica, Arial, sans-serif; font-size: 12pt; line-height: 1.5; margin: 0 0 8pt 20px;">'
                            )
                            for entry in items:
                                transfer_html += f'<div style="margin-bottom: 4pt;">{entry}</div>'
                            transfer_html += '</div></div>'
                else:
                    transfer_html += (
                        '<p style="font-family: Helvetica, Arial, sans-serif; font-size: 12pt; '
                        'color:#666666; font-style:italic;">No transfers needed this week.</p>'
                    )

                color_key_html = """
                <div style="margin-bottom: 20pt; font-family: Helvetica, Arial, sans-serif; font-size: 12pt;">
                    <strong>Action Color Key:</strong><br><br>
                    <table style="border-collapse: collapse; width: auto; font-size: 11pt;">
                        <tr>
                            <td style="width:30px; height:24px; background-color:#90EE90; border:1px solid #555;"></td>
                            <td style="padding:4px 12px;">Transfer from Overflow (digital transfer)</td>
                        </tr>
                        <tr>
                            <td style="width:30px; height:24px; background-color:#FFFF99; border:1px solid #555;"></td>
                            <td style="padding:4px 12px;">Transfer from PDX HQ / PDX HQ 2 / 3PL PDX</td>
                        </tr>
                        <tr>
                            <td style="background-color:#FFC7CE; border:1px solid #555;"></td>
                            <td style="padding:4px 12px;">Reject — Discontinued / MTO / Insufficient stock for transfer</td>
                        </tr>
                        <tr>
                            <td style="background-color:#20B2AA; border:1px solid #555;"></td>
                            <td style="padding:4px 12px;">OOS — Restock within 2 days</td>
                        </tr>
                    </table>
                </div>
                """

                body_content = f"""
                <p>Team,</p>
                <p>Please advise on the following Walmart items on backorder:</p>
                <br>
                {str(email_table)}
                <br>
                {transfer_html}
                <br>
                {color_key_html}
                <br>
                <p>Thank you,</p>
                <br>
                <hr>
                <p style="font-family: Helvetica, Arial, sans-serif; font-size: 10pt; color: #666666;">Full inventory detail:</p>
                {str(st.session_state.modified_table)}
                """

                # Style table cells
                final_soup  = BeautifulSoup(body_content, "html.parser")
                final_table = final_soup.find("table")
                if final_table:
                    for cell in final_table.find_all(["td", "th"]):
                        style = cell.get("style", "")
                        cell["style"] = (
                            "font-family: Helvetica, Arial, sans-serif !important; "
                            "font-size: 12pt !important; " + style
                        )
                    for row in final_table.find_all("tr"):
                        row["style"] = "height: 24pt;"

                styled_body = f"""
                <html>
                <head>
                    <style type="text/css">
                        body, p, div {{ font-family: Helvetica, Arial, sans-serif !important; font-size: 12pt !important; }}
                    </style>
                </head>
                <body>{final_soup.prettify()}</body>
                </html>
                """

                # ── UPDATE THESE RECIPIENTS ──
                to_emails = "recipient1@your-company.com; recipient2@your-company.com"
                cc_emails = f"{user_email}; colleague@yourcompany.com"
                subject   = f"Walmart Orders - Week of {formatted_date} - {folder_path_today}"

                threading.Thread(
                    target=create_outlook_draft,
                    args=(styled_body, to_emails, cc_emails, subject),
                    daemon=True
                ).start()

                st.success("Reply draft opened in Outlook!")

            except Exception as e:
                err = f"Error generating reply: {str(e)}\n{traceback.format_exc()}"
                append_log(err)
                save_log()
                st.error(err)

    # ── Reset for another run ──
    st.divider()
    if st.button("↺ Reset for another run", use_container_width=True):
        for key in ['phase', 'modified_table', 'slim_table',
                    'transfers_by_location', 'total_orders',
                    'log_details', 'error_message']:
            del st.session_state[key]
        st.rerun()