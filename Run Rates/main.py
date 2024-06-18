import os
import time
import logging
import pandas as pd
from datetime import datetime
from plyer import notification
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
from io import StringIO
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from requests_ntlm import HttpNtlmAuth

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s',
    handlers=[
        logging.FileHandler("app.log"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# Define the download directory
downloads_dir = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(downloads_dir):
    os.makedirs(downloads_dir)
else:
    for file in os.listdir(downloads_dir):
        file_path = os.path.join(downloads_dir, file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            logger.error(f"Error deleting file: {e}", exc_info=True)

#these act as something like a 'default' run rate for each cell, which can be updated by the user
cell_run_rates = {
    "800": 434.16,
    "801": 1806.20,
    "802": 10.00,
    "803": 1195.80,
    "804": 3545.20,
    "805": 0.00,
    "806": 0.00,
    "807": 0.00,
    "808": 0.00,
    "809": 73.00,
    "810": 3749.20,
    "811": 3066.80,
    "812": 1288.40,
    "813": 773.60,
    "814": 546.60,
    "815": 458.40,
    "816": 1610.20,
    "817": 2131.40,
    "818": 21.92,
    "819": 322.20,
    "820": 4696.00,
    "821": 3095.20,
    "822": 2634.00,
    "823": 1889.76,
    "824": 770.80,
    "825": 504.40,
    "826": 1681.20,
    "827": 824.40,
    "828": 3120.00,
    "829": 1095.40,
    "830": 12000,
    "831": 12000,
    "832": 75.00,
    "833": 1440.00,
    "834": 1413.00,
    "836": 205.80,
    "837": 83.40,
    "838": 1280.80,
    "840": 460.00,
    "841": 460.00,
    "842": 460.00,
    "843": 460.00,
    "844": 460.00,
    "850": 460.00,
    "G3A": 0.00,
    "LB": 1031.00,
    "MAN": 6976.00,
    "47P": 0.00,
    "TR4": 132.48,
    "83I": 6976.00,
    "83H": 1413.00,
    "835": 1413.00,
    "839": 1031.00,
    "KOP": 0.00,
    "HVY": 1413.00
}

df_cell_run = pd.DataFrame.from_dict(cell_run_rates, orient='index', columns=['Run Rate'])

def configure_options():
    options = Options()
    for arg in ["--headless","--disable-gpu", "--allow-running-insecure-content", "--disable-web-security", "--unsafely-treat-insecure-origin-as-secure=http://hffsuk02"]:
        options.add_argument(arg)
    prefs = {
        "download.default_directory": downloads_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--disable-features=InsecureDownloadWarnings")
    return options

def get_latest_file_path(directory, extension=".xlsx"):
    files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith(extension)]
    return max(files, key=os.path.getctime) if files else None

def is_file_downloaded(directory, initial_files, timeout=120):
    elapsed_time = 0
    while elapsed_time < timeout:
        current_files = set(os.listdir(directory))
        new_files = current_files - initial_files
        if new_files:
            new_file = new_files.pop()
            new_file_path = os.path.join(directory, new_file)
            if new_file_path.endswith(".xlsx"):
                return new_file_path
        time.sleep(1)
        elapsed_time += 1
    return None

def itemlistscraper():
    def get_latest_file_with_item_name(directory, timeout=60):
        start_time = time.time()
        while (time.time() - start_time) < timeout:
            for file_name in os.listdir(directory):
                if "Item" in file_name and file_name.endswith(".xlsx"):
                    return os.path.join(directory, file_name)
            time.sleep(1)
        return None

    try:
        item_url = "http://hffsuk02/Reports/report/ReportsUK/Item/ItemListMDeptWC"
        logger.debug(f"Starting itemlistscraper with URL: {item_url}")
        options = configure_options()
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 30)

        driver.get(item_url)
        driver.fullscreen_window()
        time.sleep(5)

        frames = driver.find_elements(By.TAG_NAME, "iframe")
        if frames:
            driver.switch_to.frame(frames[0])
            logger.debug("Navigated to frame")

        dropdown_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@id='ReportViewerControl_ctl04_ctl03_ctl01']")))
        driver.execute_script("arguments[0].click();", dropdown_button)
        time.sleep(4)

        select_all_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl04_ctl03_divDropDown_ctl00")))
        driver.execute_script("arguments[0].click();", select_all_button)
        time.sleep(2)

        view_report_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl04_ctl00")))
        driver.execute_script("arguments[0].click();", view_report_button)
        time.sleep(18)

        excel_dropdown_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl05_ctl04_ctl00_ButtonImg")))
        driver.execute_script("arguments[0].click();", excel_dropdown_button)
        time.sleep(3)

        excel_download_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl05_ctl04_ctl00_Menu > div:nth-child(2) > a")))
        driver.execute_script("arguments[0].click();", excel_download_button)
        logger.info("Download initiated")

        latest_file = get_latest_file_with_item_name(downloads_dir)
        driver.quit()

        if latest_file:
            logger.info(f"File downloaded successfully: {latest_file}")
            df = pd.read_excel(latest_file)
            return df
        else:
            logger.warning("No new file was downloaded")
            return None
    except Exception as e:
        logger.error(f"Error in itemlistscraper: {e}", exc_info=True)
        if 'driver' in locals():
            driver.quit()
        return None

def codedatescraper():
    try:
        codate_url = "http://hffsuk02/Reports/report/ReportsUK/Customer/CoDate2-X"
        logger.debug(f"Starting codedatescraper with URL: {codate_url}")
        options = configure_options()
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 20)

        driver.get(codate_url)
        driver.fullscreen_window()
        time.sleep(10)

        frames = driver.find_elements(By.TAG_NAME, "iframe")
        if frames:
            driver.switch_to.frame(frames[0])
            logger.debug("Navigated to frame")

        dropdown_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl05_ctl04_ctl00_ButtonImg")))
        driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_button)
        time.sleep(5)
        driver.execute_script("arguments[0].click();", dropdown_button)
        time.sleep(10)

        excel_download_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#ReportViewerControl_ctl05_ctl04_ctl00_Menu > div:nth-child(2) > a")))
        driver.execute_script("arguments[0].scrollIntoView(true);", excel_download_button)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", excel_download_button)
        time.sleep(15)
        logger.info("Download initiated")

        latest_file = None
        attempt_time = 0
        while not latest_file and attempt_time < 60:
            logger.debug(f"Attempt {attempt_time + 1}: Checking for the latest file")
            latest_file = get_latest_file_path(downloads_dir)
            time.sleep(1)
            attempt_time += 1

        driver.quit()

        if latest_file:
            logger.info(f"File downloaded successfully: {latest_file}")
            df = pd.read_excel(latest_file)
            return df
        else:
            logger.warning("No new file was downloaded")
            return None
    except Exception as e:
        logger.error(f"Error in codedatescraper: {e}", exc_info=True)
        if 'driver' in locals():
            driver.quit()
        return None

def file_analysis(codate_df, itemlist_df, cell_run_df):
    df_codate = pd.DataFrame(codate_df)
    df_itemlist = pd.DataFrame(itemlist_df)
    df_cell_run = pd.DataFrame(cell_run_df)

    print("Initial CoDate DataFrame:")
    print(df_codate.head())

    print("Initial ItemList DataFrame:")
    print(df_itemlist.head())

    print("Initial Cell Run DataFrame:")
    print(df_cell_run.head())

    if 'Item Number' not in df_codate.columns:
        raise KeyError("'Item Number' column is missing in codate DataFrame.")
    if 'Run Rate' not in df_cell_run.columns:
        raise KeyError("'Run Rate' column is missing in cell run DataFrame.")

    df_cell_run = df_cell_run.reset_index().rename(columns={'index': 'Item Number'})

    print("Cell Run DataFrame after resetting index:")
    print(df_cell_run.head())

    df_codate_cleaned = df_codate.dropna(subset=['CustID']) if 'CustID' in df_codate.columns else df_codate
    df_itemlist_cleaned = df_itemlist.dropna(subset=['Quantity'])

    print("Cleaned CoDate DataFrame:")
    print(df_codate_cleaned.head())

    print("Cleaned ItemList DataFrame:")
    print(df_itemlist_cleaned.head())

    df_itemlist_cleaned['MinutesOfJob'] = df_itemlist_cleaned['Quantity']
    df_itemlist_cleaned['TotalMinutesOfJob'] = df_itemlist_cleaned.groupby('Buyer')['MinutesOfJob'].transform('sum')

    total_quantity_per_buyer = df_itemlist_cleaned.groupby('Buyer')['Quantity'].sum().reset_index()
    total_quantity_per_part = df_itemlist_cleaned.groupby('Parent')['Quantity'].sum().reset_index()

    print("Total Quantity per Buyer:")
    print(total_quantity_per_buyer.head())

    print("Total Quantity per Part:")
    print(total_quantity_per_part.head())

    total_quantity_per_buyer_codate = df_codate_cleaned.groupby('Buyer')['WCRMins'].sum().reset_index()
    total_quantity_per_part_codate = df_codate_cleaned.groupby('Item Number')['WCRMins'].sum().reset_index()

    print("Total Quantity per Buyer CoDate:")
    print(total_quantity_per_buyer_codate.head())

    print("Total Quantity per Part CoDate:")
    print(total_quantity_per_part_codate.head())

    total_quantity_per_buyer['Alert'] = total_quantity_per_buyer['Quantity'].apply(lambda x: 'Alert' if x > 4000 else '')
    total_quantity_per_part['Alert'] = total_quantity_per_part['Quantity'].apply(lambda x: 'Alert' if x > 4000 else '')
    total_quantity_per_buyer_codate['Alert'] = total_quantity_per_buyer_codate['WCRMins'].apply(lambda x: 'Alert' if x > 4000 else '')
    total_quantity_per_part_codate['Alert'] = total_quantity_per_part_codate['WCRMins'].apply(lambda x: 'Alert' if x > 4000 else '')

    total_quantity_per_buyer = total_quantity_per_buyer.sort_values(by='Quantity', ascending=False)
    total_quantity_per_part = total_quantity_per_part.sort_values(by='Quantity', ascending=False)
    total_quantity_per_buyer_codate = total_quantity_per_buyer_codate.sort_values(by='WCRMins', ascending=False)
    total_quantity_per_part_codate = total_quantity_per_part_codate.sort_values(by='WCRMins', ascending=False)

    if 'PromShip' in df_codate_cleaned.columns:
        df_codate_cleaned['PromShip'] = pd.to_datetime(df_codate_cleaned['PromShip'], format='%d-%b-%y')
    df_codate_cleaned = df_codate_cleaned.sort_values(by='PromShip', ascending=True)
    df_codate_cleaned = df_codate_cleaned[df_codate_cleaned['Buyer'].str.startswith('8')]

    df_codate_cleaned = pd.merge(df_codate_cleaned, df_cell_run, on='Item Number', how='left')

    print("Merged CoDate DataFrame with Cell Run DataFrame:")
    print(df_codate_cleaned.head())

    df_codate_cleaned['MinutesOfJob'] = df_codate_cleaned['WCRMins'] * df_codate_cleaned['OrderQty']
    df_codate_cleaned['Daily Cell Run Rate'] = df_codate_cleaned['Buyer'].map(cell_run_rates)
    df_codate_cleaned['ExceedsHalfaDaysWork'] = df_codate_cleaned['MinutesOfJob'] > (0.5 * df_codate_cleaned['Daily Cell Run Rate'])

    df_critical = df_codate_cleaned[df_codate_cleaned['ExceedsHalfaDaysWork']]

    print("Critical Entries DataFrame:")
    print(df_critical.head())

    with pd.ExcelWriter('Item Breakdown.xlsx', engine='xlsxwriter') as writer:
        df_codate_cleaned.to_excel(writer, sheet_name='CoDate Data', index=False, na_rep='NA')
        df_itemlist_cleaned.to_excel(writer, sheet_name='Item List Data', index=False, na_rep='NA')
        df_cell_run.to_excel(writer, sheet_name='Cell Run Rates', index=False, na_rep='NA')
        df_critical.to_excel(writer, sheet_name='Critical Entries', index=False, na_rep='NA')
        total_quantity_per_buyer.to_excel(writer, sheet_name='Total Quantity per Buyer', index=False, na_rep='NA')
        total_quantity_per_part.to_excel(writer, sheet_name='Total Quantity per Part', index=False, na_rep='NA')
        total_quantity_per_buyer_codate.to_excel(writer, sheet_name='Total Quantity per Buyer CoDate', index=False, na_rep='NA')
        total_quantity_per_part_codate.to_excel(writer, sheet_name='Total Quantity per Part CoDate', index=False, na_rep='NA')
    
    os.startfile('Item Breakdown.xlsx')

    return total_quantity_per_buyer, total_quantity_per_part, total_quantity_per_buyer_codate, total_quantity_per_part_codate, df_critical

def load_data():
    attempts = 3
    codate_df, itemlist_df = None, None

    for attempt in range(attempts):
        codate_df = codedatescraper()
        if codate_df is not None:
            print(f"CoDate DataFrame loaded.")
            break
        else:
            print(f"Attempt {attempt + 1} for CoDate scraper failed.")
        time.sleep(5)

    for attempt in range(attempts):
        itemlist_df = itemlistscraper()
        if itemlist_df is not None:
            print(f"ItemList DataFrame loaded.")
            break
        else:
            print(f"Attempt {attempt + 1} for ItemList scraper failed.")
        time.sleep(5)

    if codate_df is not None and itemlist_df is not None:
        print("Data processing completed successfully. Variables stored in codate_df and itemlist_df.")
        file_analysis(codate_df, itemlist_df, df_cell_run)
    else:
        print("Data processing failed. Check the logs for more information.")
        
def create_tkinter_gui():
    root = tk.Tk()
    root.attributes('-fullscreen', True)  # Set fullscreen mode
    root.title("Run Rate Adjustments")

    main_frame = ttk.Frame(root, padding="10")
    main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    
    # Create a canvas and scrollbar
    canvas = tk.Canvas(main_frame)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

    main_frame.grid_rowconfigure(0, weight=1)
    main_frame.grid_columnconfigure(0, weight=1)

    ttk.Label(scrollable_frame, text="Adjust Run Rates for Each Cell", font=("Helvetica", 16)).grid(row=0, column=0, columnspan=2, pady=10)

    row = 1
    entries = {}
    for cell, rate in cell_run_rates.items():
        ttk.Label(scrollable_frame, text=f"Cell {cell}:", font=("Helvetica", 12)).grid(row=row, column=0, sticky=tk.E, padx=10, pady=5)
        entry = ttk.Entry(scrollable_frame, font=("Helvetica", 12))
        entry.insert(0, rate)
        entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=10, pady=5)
        entries[cell] = entry
        row += 1

    ttk.Label(scrollable_frame, text="Paste Excel Data:", font=("Helvetica", 12)).grid(row=row, column=0, pady=10, sticky=tk.W)
    row += 1
    text_widget = tk.Text(scrollable_frame, height=10, width=80, font=("Helvetica", 12))
    text_widget.grid(row=row, column=0, columnspan=2, pady=10, padx=10)

    def update_run_rates_from_paste():
        try:
            # Get the pasted text from the Text widget
            pasted_data = text_widget.get("1.0", tk.END).strip()

            # Process the pasted text to ensure correct format
            lines = pasted_data.splitlines()
            processed_data = []
            for line in lines:
                parts = line.split()
                if len(parts) >= 2:
                    cell = parts[0]
                    rate = parts[-1].replace(",", "")
                    processed_data.append(f"{cell}\t{rate}")

            processed_data = "\n".join(processed_data)

            # Convert the pasted text to a pandas DataFrame
            data = StringIO(processed_data)
            df_pasted = pd.read_csv(data, sep="\t", header=None, names=["Cell", "Run Rate"])

            # Convert Run Rate to float
            df_pasted["Run Rate"] = df_pasted["Run Rate"].astype(float)

            # Update the cell_run_rates dictionary and DataFrame
            for index, row in df_pasted.iterrows():
                cell = str(row["Cell"])
                rate = row["Run Rate"]
                cell_run_rates[cell] = rate  # Add or update the cell run rate

            # Update the GUI entries
            for cell, entry in entries.items():
                entry.delete(0, tk.END)
                entry.insert(0, cell_run_rates[cell])
            for cell, rate in cell_run_rates.items():
                if cell not in entries:
                    ttk.Label(scrollable_frame, text=f"Cell {cell}:", font=("Helvetica", 12)).grid(row=row, column=0, sticky=tk.E, padx=10, pady=5)
                    entry = ttk.Entry(scrollable_frame, font=("Helvetica", 12))
                    entry.insert(0, rate)
                    entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=10, pady=5)
                    entries[cell] = entry
                    row += 1
            
            messagebox.showinfo("Success", "Run rates updated from pasted data!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update run rates: {e}")

    row += 1
    ttk.Button(scrollable_frame, text="Update from Paste", command=update_run_rates_from_paste, style="TButton").grid(row=row, column=0, columnspan=2, pady=10)

    def save_run_rates():
        global cell_run_rates
        for cell, entry in entries.items():
            cell_run_rates[cell] = float(entry.get())
        logger.info("Run rates updated.")
        root.destroy()

    row += 1
    ttk.Button(scrollable_frame, text="Save", command=save_run_rates, style="TButton").grid(row=row, column=0, columnspan=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_tkinter_gui()
    load_data()
