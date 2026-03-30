# ──────────────────────────────────────────────────────────────────────────────
# DEMO VERSION
# Sanitized for portfolio use. Credentials, internal file paths, and email
# addresses have been replaced with placeholders. Requires a live NetSuite
# account, Walmart CA vendor portal access, and a configured secrets.toml
# to run. See README.md for context.
# ──────────────────────────────────────────────────────────────────────────────

import streamlit as st
import os
import sys
import time
import shutil
from datetime import datetime
import pandas as pd
import glob
import zipfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoAlertPresentException

# Set environment variables from secrets
os.environ['LONG_EMAIL'] = st.secrets['LONG_EMAIL']
os.environ['NETSUITE_PASSWORD'] = st.secrets['NETSUITE_PASSWORD']
os.environ['NETSUITE_KEY'] = st.secrets['NETSUITE_KEY']
source_folder = st.secrets['DOWNLOAD_FOLDER']
target_folder_root = st.secrets.get("WALMART_BASE_PROJECTS_FOLDER", "./projects")

# Utilities folder
UTILITIES_PATH = os.path.join(os.path.dirname(__file__), "..", "..", "Utilities")
sys.path.insert(0, UTILITIES_PATH)
from netsuite_login import netsuite_login

st.title("🍁 Walmart Canada - Order Processing")

st.markdown("""
This app automates:
- The process of updating carrier data (such as DC #, SCAC, Carrier, and BOL #) for Walmart Canada orders in NetSuite.
- Exports the Label Line Flat CSV file for the orders.
- The file is named based on unique PO numbers.
You can download the processed file directly from here.
""")

def edit_field(driver, span_id, new_value):
    span = driver.find_element(By.ID, span_id)
    actions = ActionChains(driver)
    actions.double_click(span).perform()

    # Wait for the input field to appear in the parent td
    td = span.find_element(By.XPATH, "..")
    input_field = WebDriverWait(td, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "input"))
    )

    input_field.clear()
    input_field.send_keys(new_value)
    input_field.send_keys(Keys.ENTER)

    # Wait briefly for the change to take effect
    WebDriverWait(driver, 5).until(
        EC.text_to_be_present_in_element((By.ID, span_id), new_value)
    )


def accept_alert_if_present(driver):
    try:
        alert = driver.switch_to.alert
        alert.accept()
    except NoAlertPresentException:
        pass


if st.button("Run Automation"):
    with st.spinner("Processing... Please wait."):
        # Define today_str and target_folder early
        today_str = datetime.today().strftime('%Y-%m (%b)-%d')
        target_folder = os.path.join(target_folder_root, f'Walmart Orders - CA - {today_str}')

        # Create the target folder if it doesn't exist
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)

        # Set up Chrome options for downloads
        chrome_options = Options()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": source_folder,
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True,
            "plugins.always_open_pdf_externally": True  # Force PDF download
        })

        # Initialize Selenium WebDriver
        driver = webdriver.Chrome(options=chrome_options)

        # Navigate to the first search results page
        driver.get("https://4423908.app.netsuite.com/app/common/search/searchresults.nl?searchid=8474")


        # Perform login without displaying log messages
        def log_message(message):
            pass  # Do nothing to suppress output


        netsuite_login(driver, log_message)

        # Wait for the table to load
        wait = WebDriverWait(driver, 30)
        table = wait.until(EC.presence_of_element_located((By.ID, "div__body")))

        # Get all data rows (skip header and hidden rows)
        rows = table.find_elements(By.CSS_SELECTOR, "tbody tr.uir-list-row-tr")

        # Collect data for all orders
        orders = []
        for row in rows:
            row_id = row.get_attribute("id")  # e.g., "row0"
            row_index = int(row_id.replace("row", ""))

            tds = row.find_elements(By.TAG_NAME, "td")

            # Get SO and PO
            so = tds[3].text.strip()
            po = tds[4].text.strip()

            # Get Cases
            cases_str = tds[23].text.strip()
            cases = int(cases_str) if cases_str else 0

            # Determine Carrier and SCAC
            carrier = "UPS STANDARD" if cases <= 35 else "FastFrate"
            scac = "UPSS" if cases <= 35 else "CFFO"

            # Get Node, BILL OF LADING
            node = tds[7].text.strip()
            bill_of_lading = tds[16].text.strip()

            orders.append({
                'row_index': row_index,
                'so': so,
                'po': po,
                'node': node,
                'cases': cases,
                'bill_of_lading': bill_of_lading,
                'carrier': carrier,
                'scac': scac
            })

        # Display green bubbles for each order
        for order in orders:
            st.markdown(
                f'<span style="background-color:green;color:white;padding:5px;border-radius:5px;">{order["so"]} - {order["po"]} - {order["cases"]} Cases - {order["carrier"]}</span>',
                unsafe_allow_html=True)

        # Update DC # for all rows
        for order in orders:
            dc_span_id = f"lstinln_{order["row_index"]}_1"
            edit_field(driver, dc_span_id, order['node'])
            time.sleep(2)  # Pause after each edit to avoid concurrent edit errors

        # Update SCAC for all rows
        for order in orders:
            scac_span_id = f"lstinln_{order["row_index"]}_2"
            edit_field(driver, scac_span_id, order['scac'])
            time.sleep(2)

        # Update Carrier for all rows
        for order in orders:
            carrier_span_id = f"lstinln_{order["row_index"]}_3"
            edit_field(driver, carrier_span_id, order['carrier'])
            time.sleep(2)

        # Update BOL # for all rows
        for order in orders:
            bol_span_id = f"lstinln_{order["row_index"]}_5"
            edit_field(driver, bol_span_id, order['bill_of_lading'])
            time.sleep(2)

        # Click off the last editable field to ensure save
        header_cell = driver.find_element(By.ID, "div__lab1")
        header_cell.click()

        # Add another 5-second pause after the last order
        time.sleep(5)

        # Export the label line flat file first
        driver.get("https://4423908.app.netsuite.com/app/common/search/searchresults.nl?searchid=4993&whence=")

        # Wait for the page to load and find the Export CSV button
        export_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@title="Export - CSV"]')))
        export_button.click()

        # Wait for the file to download
        file_pattern = os.path.join(source_folder, "WalmartLabelLinesFlatALLLOADSCA*.csv")
        timeout = 60  # seconds
        start_time = time.time()
        while not glob.glob(file_pattern):
            if time.time() - start_time > timeout:
                st.error("Download timeout.")
                driver.quit()
                break
            time.sleep(1)

        # Once downloaded, process the file
        processed_file_path = None
        for filename in os.listdir(source_folder):
            if filename.startswith("WalmartLabelLinesFlatALLLOADS") and filename.endswith('.csv'):
                source_file_path = os.path.join(source_folder, filename)
                df = pd.read_csv(source_file_path)  # Read the CSV file
                unique_pos = df['PO'].unique()  # Get unique PO values
                unique_pos_str = '_'.join(map(str, unique_pos))  # Convert each to string and concatenate

                # New file name with unique POs
                new_filename = f"Walmart Label Lines Flat - PO #{unique_pos_str}.csv"
                target_file_path = os.path.join(target_folder, new_filename)

                # Move and rename the file
                shutil.move(source_file_path, target_file_path)
                st.write(f"Generated {new_filename}")
                processed_file_path = target_file_path

        # # Offer download button if file was processed
        # if processed_file_path:
        #     with open(processed_file_path, "rb") as f:
        #         st.download_button(
        #             label="Download Processed CSV",
        #             data=f,
        #             file_name=new_filename,
        #             mime="text/csv"
        #         )

        # Now create the pack slip import CSV
        sos = [order['so'] for order in orders]
        pos = [order['po'] for order in orders]
        unique_pos_str = '_'.join(pos)
        pack_slip_filename = f"Walmart CA Pack Slip Import - PO #{unique_pos_str}.csv"
        pack_slip_path = os.path.join(target_folder, pack_slip_filename)

        # Create DataFrame for pack slip
        pack_slip_data = []
        for so in sos:
            pack_slip_data.append([so, 'Sales Order', '', '', 'N'])

        pack_slip_df = pd.DataFrame(pack_slip_data,
                                    columns=['Order Number', 'Transaction Type', 'Weight', 'Tracking Number',
                                             'Label Integration'])
        pack_slip_df.to_csv(pack_slip_path, index=False)
        # st.write(f"Created pack slip import file: {pack_slip_filename}")
        #
        # # Offer download for pack slip
        # with open(pack_slip_path, "rb") as f:
        #     st.download_button(
        #         label="Download Pack Slip Import CSV",
        #         data=f,
        #         file_name=pack_slip_filename,
        #         mime="text/csv"
        #     )

        # After offering label line, proceed to generate pack slips (upload and submit)
        driver.get("https://4423908.app.netsuite.com/app/accounting/transactions/salesordermanager.nl?type=importcsv")

        # Wait for the page to load
        wait.until(EC.presence_of_element_located((By.NAME, "order")))

        # Upload the pack slip import file
        file_input = driver.find_element(By.NAME, "order")
        file_input.send_keys(pack_slip_path)
        time.sleep(2)
        accept_alert_if_present(driver)

        # Set bosslocation to "3PL - Ontario Expeditors"
        location_input = driver.find_element(By.ID, "bosslocation_display")
        location_input.send_keys("3PL - Ontario Expeditors")
        time.sleep(1)  # Wait for suggestions
        location_input.send_keys(Keys.ARROW_DOWN)
        location_input.send_keys(Keys.ENTER)
        time.sleep(2)  # Pause to ensure selection
        accept_alert_if_present(driver)

        # Set shipstatus to "Packed"
        shipstatus_input = driver.find_element(By.ID, "inpt_shipstatus_3")
        shipstatus_input.clear()
        shipstatus_input.send_keys("Packed")
        time.sleep(1)  # Wait for suggestions if needed
        shipstatus_input.send_keys(Keys.ENTER)
        time.sleep(2)  # Pause to ensure selection
        accept_alert_if_present(driver)

        # Additional pause and alert check before submit
        time.sleep(2)
        accept_alert_if_present(driver)

        # Wait for the submission results page to load
        wait.until(EC.presence_of_element_located((By.ID, "div__body")))

        # Monitor the latest submission by Richard Hardisty
        complete = False
        while not complete:
            # Refresh the page
            driver.refresh()
            wait.until(EC.presence_of_element_located((By.ID, "div__body")))

            # Get all rows
            rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr.uir-list-row-tr")

            # Find rows created by Richard Hardisty
            richard_rows = []
            for row in rows:
                tds = row.find_elements(By.TAG_NAME, "td")
                if len(tds) >= 7:
                    created_by = tds[6].text.strip()
                    if "Richard Hardisty" in created_by:
                        submission_id = int(tds[0].text.strip())
                        status = tds[2].text.strip()
                        percent = tds[3].text.strip()
                        richard_rows.append({
                            'submission_id': submission_id,
                            'status': status,
                            'percent': percent
                        })

            if richard_rows:
                # Find the latest (max submission_id)
                latest = max(richard_rows, key=lambda x: x['submission_id'])
                st.write(f"Current fulfillment progress: {latest['percent']}")
                if latest['status'] == "Complete" and latest['percent'] == "100.0%":
                    complete = True
                else:
                    time.sleep(20)  # Wait before next refresh
            else:
                time.sleep(20)  # Wait if no rows found

        # Once complete, navigate to the next URL
        driver.get("https://4423908.app.netsuite.com/app/common/search/searchresults.nl?scrollid=8474&searchid=11050")

        st.write("Generating Carton Labels...")

        # Wait for the table to load
        wait.until(EC.presence_of_element_located((By.ID, "div__body")))

        # Get all data rows
        rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr.uir-list-row-tr")

        # For each row, copy IF # to PRO #
        for row in rows:
            row_id = row.get_attribute("id")  # e.g., "row0"
            row_index = int(row_id.replace("row", ""))

            tds = row.find_elements(By.TAG_NAME, "td")

            # Get IF #
            if_num = tds[6].text.strip()

            # Edit PRO # (span id lstinln_{row_index}_5)
            pro_span_id = f"lstinln_{row_index}_5"
            edit_field(driver, pro_span_id, if_num)
            time.sleep(2)

        # Wait a few seconds after edits
        time.sleep(5)

        # Click on Auto-Pack links
        for row in rows:
            tds = row.find_elements(By.TAG_NAME, "td")
            auto_pack_link = tds[21].find_element(By.TAG_NAME, "a").get_attribute("href")

            # Open in new tab
            driver.execute_script(f"window.open('{auto_pack_link}', '_blank');")
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(5)
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        # Open View links, wait 5 sec, close tab
        for row in rows:
            tds = row.find_elements(By.TAG_NAME, "td")
            view_link = tds[2].find_elements(By.TAG_NAME, "a")[0].get_attribute("href")

            # Open in new tab
            driver.execute_script(f"window.open('{view_link}', '_blank');")
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(5)
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        # Then open Batch Print Labels links
        for row in rows:
            tds = row.find_elements(By.TAG_NAME, "td")
            batch_print_link = tds[22].find_element(By.TAG_NAME, "a").get_attribute("href")

            # Open in new tab
            driver.execute_script(f"window.open('{batch_print_link}', '_blank');")
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(5)
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        # Open View links, wait 5 sec, close tab
        for row in rows:
            tds = row.find_elements(By.TAG_NAME, "td")
            view_link = tds[2].find_elements(By.TAG_NAME, "a")[0].get_attribute("href")

            # Open in new tab
            driver.execute_script(f"window.open('{view_link}', '_blank');")
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(5)
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        # Refresh the search page
        driver.refresh()
        wait.until(EC.presence_of_element_located((By.ID, "div__body")))

        # Get rows again after refresh
        rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr.uir-list-row-tr")

        # Collect unique POs and their first row
        po_to_row = {}
        for row in rows:
            tds = row.find_elements(By.TAG_NAME, "td")
            po = tds[7].text.strip()
            if po not in po_to_row:
                po_to_row[po] = row

        # For each unique PO, click the print label link in the first row
        for po, row in po_to_row.items():
            tds = row.find_elements(By.TAG_NAME, "td")
            print_labels_link = tds[23].find_element(By.TAG_NAME, "a")
            driver.execute_script("arguments[0].click();", print_labels_link)
            time.sleep(2)  # Brief pause between clicks to allow download initiation

        # After all clicks, run the renaming and merging
        import re
        from PyPDF2 import PdfReader, PdfMerger

        def rename_to_order_number(source_directory):
            """Rename PDF files in source directory based on order number."""
            if not os.path.exists(source_directory):
                print(f"Directory {source_directory} does not exist.")
                return []

            renamed_files = []
            for filename in os.listdir(source_directory):
                if filename.lower().endswith('.pdf'):
                    file_path = os.path.join(source_directory, filename)
                    try:
                        with open(file_path, 'rb') as file:
                            pdf_reader = PdfReader(file)
                            order_number = None
                            if len(pdf_reader.pages) > 0:
                                page = pdf_reader.pages[0]
                                text = page.extract_text()
                                if text:
                                    lines = text.split('\n')
                                    for line in lines:
                                        if line.startswith("ORDER #: "):
                                            order_number = line[len("ORDER #: "):].strip()
                                            break

                        if order_number:
                            sanitized_order = ''.join(c for c in order_number if c.isalnum() or c in ('-', '_'))
                            new_filename = f"{sanitized_order}.pdf"
                            new_file_path = os.path.join(source_directory, new_filename)

                            if os.path.exists(new_file_path):
                                print(f"Cannot rename {filename}: File {new_filename} already exists.")
                            else:
                                max_retries = 3
                                for attempt in range(max_retries):
                                    try:
                                        os.rename(file_path, new_file_path)
                                        print(f"Renamed {filename} to {new_filename}")
                                        renamed_files.append(new_filename)
                                        break
                                    except PermissionError as e:
                                        if "[WinError 32]" in str(e):
                                            print(
                                                f"Attempt {attempt + 1}/{max_retries}: File {filename} is locked. Waiting 2 seconds...")
                                            time.sleep(2)
                                        else:
                                            raise e
                                else:
                                    print(
                                        f"Failed to rename {filename} after {max_retries} attempts: File still in use.")
                        else:
                            print(f"No order number found on first page of {filename}")

                    except Exception as e:
                        print(f"Error processing {filename}: {str(e)}")

            return renamed_files


        def move_and_merge_files(source_directory, target_directory):
            """Move renamed files to target directory and merge PDFs starting with a number."""
            # Create target directory if it doesn't exist
            os.makedirs(target_directory, exist_ok=True)

            # Move renamed files
            renamed_files = rename_to_order_number(source_directory)
            moved_files = []
            for filename in renamed_files:
                source_path = os.path.join(source_directory, filename)
                target_path = os.path.join(target_directory, filename)
                try:
                    os.rename(source_path, target_path)
                    print(f"Moved {filename} to {target_directory}")
                    moved_files.append(filename)
                except Exception as e:
                    print(f"Error moving {filename}: {str(e)}")

            # Merge PDFs that start with a number
            pdf_files = []
            pattern = re.compile(r'^\d+.*\.pdf$')
            for file_name in os.listdir(target_directory):
                if pattern.match(file_name):
                    pdf_files.append(file_name)

            pdf_files.sort()
            merger = PdfMerger()

            for pdf in pdf_files:
                merger.append(os.path.join(target_directory, pdf))

            if pdf_files:
                po_numbers = '_'.join([os.path.splitext(file_name)[0] for file_name in pdf_files])
                new_file_name = f"Walmart CA Carton Labels - PO #{po_numbers}.pdf"
                new_file_path = os.path.join(target_directory, new_file_name)

                merger.write(new_file_path)
                merger.close()

                print(f"Merged PDF saved as: {new_file_path}")

                # Delete original PDF files
                for pdf in pdf_files:
                    try:
                        os.remove(os.path.join(target_directory, pdf))
                        print(f"Deleted original file: {pdf}")
                    except Exception as e:
                        print(f"Error deleting {pdf}: {str(e)}")
            else:
                print("No PDFs found that start with a number.")

            st.write(f"Carton Labels Successfully Generated as {new_file_name}.")
            st.write("Generating Pack Slips...")

        # Run the move and merge
        move_and_merge_files(source_folder, target_folder)

        # After previous steps, go to the packing slips page
        driver.get(
            "https://4423908.app.netsuite.com/app/accounting/print/returnform.nl?trantype=&printtype=packingslip&method=print&printtype=packingslip&whence=&title=Packing+Slips+and+Return+Forms")

        # Wait for page to load
        wait.until(EC.presence_of_element_located((By.ID, "item_Transaction_NAME_display")))

        # Set the input to "830999 WALMART CANADA"
        name_input = driver.find_element(By.ID, "item_Transaction_NAME_display")
        name_input.clear()
        name_input.send_keys("830999 WALMART CANADA")
        time.sleep(1)
        name_input.send_keys(Keys.ARROW_DOWN)
        name_input.send_keys(Keys.ENTER)
        time.sleep(2)

        # Ensure reprint checkbox is checked
        reprint_checkbox = driver.find_element(By.ID, "reprint_fs_inp")
        if not reprint_checkbox.is_selected():
            driver.execute_script("arguments[0].click();", reprint_checkbox)
        time.sleep(2)

        # Wait 10 seconds
        time.sleep(30)

        # Click Mark All
        mark_all_button = driver.find_element(By.ID, "markall")
        mark_all_button.click()
        time.sleep(2)

        # Click Print
        print_button = driver.find_element(By.ID, "nl_print")
        print_button.click()
        time.sleep(2)

        # Wait 10 seconds for PDF
        time.sleep(30)

        # Handle PDF download
        pdf_pattern = os.path.join(source_folder, "*.pdf")
        download_timeout = 30
        download_start = time.time()
        downloaded_pdf = None
        while not downloaded_pdf:
            downloaded_pdfs = glob.glob(pdf_pattern)
            if downloaded_pdfs:
                downloaded_pdf = max(downloaded_pdfs, key=os.path.getctime)  # Get latest
            if time.time() - download_start > download_timeout:
                st.error("PDF download timeout.")
                break
            time.sleep(1)

        if downloaded_pdf:
            # Rename and move
            pos_str = '_'.join(pos)  # pos from earlier
            new_pdf_name = f"Walmart Pack Slips - PO #{pos_str}.pdf"
            new_pdf_path = os.path.join(target_folder, new_pdf_name)
            shutil.move(downloaded_pdf, new_pdf_path)
            st.write(f"Saved PDF: {new_pdf_name}")

        # After carton labels, go to the load list search
        driver.get("https://4423908.app.netsuite.com/app/common/search/searchresults.nl?searchid=8449")

        # Wait for the page to load and find the Export CSV button
        export_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@title="Export - CSV"]')))
        export_button.click()

        # Wait for the file to download
        load_file_pattern = os.path.join(source_folder, "5WalmartEmailLoadList*.csv")
        load_timeout = 60
        load_start = time.time()
        while not glob.glob(load_file_pattern):
            if time.time() - load_start > load_timeout:
                st.error("Load list CSV download timeout.")
                driver.quit()
                break
            time.sleep(1)

        # Once downloaded, process with the script
        import win32com.client as win32
        import pythoncom
        import re
        from PyPDF2 import PdfReader, PdfMerger  # Assuming already used, but import again if needed


        def get_data_from_csv_as_html(filename):
            # Read CSV, ensuring all data is initially treated as strings
            data = pd.read_csv(filename, dtype=str)

            # Remove the presumed total row if it's the last row
            data = data[:-1]

            # Drop the 'Item' column if present
            if 'Item' in data.columns:
                data.drop('Item', axis=1, inplace=True)

            # Convert 'Cartons' to numeric, ensuring the column exists
            if 'Cartons' in data.columns:
                data['Cartons'] = pd.to_numeric(data['Cartons'], errors='coerce')
            else:
                raise ValueError("The 'Cartons' column is missing from the CSV file.")

            # Define aggregation logic
            # For simplicity, 'first' is used for non-numeric fields, which may need to be adjusted
            aggregation_logic = {'LoadID': 'first', 'SO #': 'first', 'DC #': 'first',
                                 'Dest.': 'first', 'SCAC': 'first', 'Carrier': 'first',
                                 'Cartons': 'sum', 'Mode': 'first'}

            # Group by 'PO #' and aggregate
            grouped_data = data.groupby('PO #', as_index=False).agg(aggregation_logic)

            # Specify the desired column order
            desired_order = ['LoadID', 'PO #', 'SO #', 'DC #', 'Dest.', 'SCAC',
                             'Carrier', 'Cartons', 'Mode']

            # Reorder the DataFrame columns
            grouped_data = grouped_data[desired_order]

            # Convert to HTML
            html_table = grouped_data.to_html(border=1, index=False, justify='center', classes='table table-condensed')

            # HTML styling and message
            html_table_style = """
            <style>
            table, th, td {
                border: 1px solid black;
                border-collapse: collapse;
                text-align: center;
            }
            table {
                width: 100%;
            }
            th, td {
                padding: 10px;
            }
            </style>
            """
            message = """
            Hello,<br><br>
            Attached is the documentation for the following orders:<br><br>
            """
            full_html = message + html_table_style + html_table
            return full_html


        def zip_files(date_str):
            folder_path = os.path.join(st.secrets.get("WALMART_BASE_PROJECTS_FOLDER", "./projects"), f"Walmart Orders - CA - {date_str}")
            zip_file_name = f"Walmart Orders - CA - {date_str}.zip"
            zip_file_path = os.path.join(folder_path, zip_file_name)
            ignore_file = f"Walmart Pack Slip Import - CA - {date_str}.csv"  # File to ignore
            with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for foldername, subfolders, filenames in os.walk(folder_path):
                    if "Uploads" in subfolders:
                        subfolders.remove("Uploads")
                    for filename in filenames:
                        if filename != zip_file_name and filename != ignore_file:  # Skip the zip file and the ignore file
                            file_path = os.path.join(foldername, filename)
                            zipf.write(file_path, os.path.relpath(file_path, folder_path))
            print(f"Files zipped successfully in {zip_file_path}")
            return zip_file_path


        def send_outlook_mail(subject, body, recipients, attachment_path=None):
            pythoncom.CoInitialize()
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.Display()

            default_signature = mail.HTMLBody
            formatted_body = f'<div style="font-family: Helvetica, Arial, sans-serif; font-size: 12pt;">{body}<br>{default_signature}</div>'

            mail.To = ";".join(recipient for recipient in recipients)
            mail.Subject = subject
            mail.HTMLBody = formatted_body

            if attachment_path:
                mail.Attachments.Add(attachment_path)

            time.sleep(5)
            mail.Send()
            pythoncom.CoUninitialize()


        def main():
            directory = st.secrets.get("DOWNLOAD_FOLDER", "./downloads")
            prefix = "5WalmartEmailLoadList"
            recipients = [
                "logistics1@your-3pl.com", "logistics2@your-3pl.com", "logistics3@your-3pl.com",
                "ops1@your-company.com", "mgmt1@your-company.com", "ops2@your-company.com",
                "mgmt2@your-company.com", "ops3@your-company.com", "ops4@your-company.com",
                "ops5@your-company.com",
                "ops6@your-company.com", "ops7@your-company.com", "you@your-company.com"
            ]

            date_str_for_zip = datetime.today().strftime('%Y-%m (%b)-%d')
            zip_file_path = zip_files(date_str_for_zip)

            todays_date = datetime.now().strftime("%d-%b")

            for filename in os.listdir(directory):
                if filename.startswith(prefix) and filename.endswith('.csv'):
                    filepath = os.path.join(directory, filename)
                    data = get_data_from_csv_as_html(filepath)
                    send_outlook_mail(f"Walmart CA Orders from {todays_date}", data, recipients, zip_file_path)
                    os.remove(filepath)
                    break


        # Run the main function
        main()

        st.success("Automation completed!")
        driver.quit()
