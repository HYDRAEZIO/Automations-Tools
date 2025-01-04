import os
import re
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from threading import Thread

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Define base directories for LR Copy, Receipt, and Invoice
directories = {
    "LR Copy": "",
    "Receipt": "",
    "Invoice": ""
}

# Function to select directory
def select_directory(doc_type):
    dir_path = filedialog.askdirectory()
    if dir_path:
        directories[doc_type] = dir_path
        update_log(f"Selected {doc_type} directory: {dir_path}")

# Function to send files based on the provided Excel file
def send_files(file_path):
    try:
        # Read and prepare the Excel file
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip().str.lower()

        # Verify required columns
        required_columns = ['sms no', 'vch no.', 'lr no.', 'ref no:']
        for col in required_columns:
            if col not in df.columns:
                update_log(f"Missing column: {col}")
                return

        # Launch browser and load WhatsApp Web
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        driver.maximize_window()
        driver.get("https://web.whatsapp.com")
        messagebox.showinfo("QR Code Scan", "Please scan the QR code on WhatsApp Web and click OK to continue.")

        # Process each row in the Excel file
        for index, row in df.iterrows():
            phone_numbers = [num.strip() for num in str(row['sms no']).split(",")]
            vch_no = str(int(row['vch no.'])) if pd.notna(row['vch no.']) else None
            ref_no = normalize_ref_no(row['ref no:']) if pd.notna(row['ref no:']) else None

            for phone_number in phone_numbers:
                if not phone_number or not phone_number.isdigit():
                    continue

                if selected_docs['LR Copy'].get() and vch_no:
                    lr_file = search_file('LR Copy', vch_no)
                    if lr_file:
                        send_file(driver, phone_number, lr_file, 'LR Copy')

                if selected_docs['Receipt'].get() and vch_no:
                    receipt_file = search_file('Receipt', vch_no)
                    if receipt_file:
                        send_file(driver, phone_number, receipt_file, 'Receipt')

                if selected_docs['Invoice'].get() and ref_no:
                    invoice_file = search_file('Invoice', ref_no)
                    if invoice_file:
                        send_file(driver, phone_number, invoice_file, 'Invoice')
# created by Vivek Pandey @ https://github.com/HYDRAEZIO
        messagebox.showinfo("Process Completed", "All selected files have been processed.")
        driver.quit()

    except Exception as e:
        update_log(f"Error: {str(e)}")

# Function to normalize `Ref No:`
def normalize_ref_no(ref_no):
    match = re.search(r'\d+', str(ref_no))
    return match.group(0) if match else ref_no

# File search function
def search_file(doc_type, doc_name):
    dir_path = directories.get(doc_type)
    if not dir_path:
        update_log(f"No directory selected for {doc_type}")
        return None

    file_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.pdf', '.docx', '.xlsx', '.xls', '.txt']
    doc_name_normalized = re.sub(r'\s+', '', doc_name).lower()

    for root, dirs, files in os.walk(dir_path):
        for file in files:
            file_name_normalized = re.sub(r'\s+', '', os.path.splitext(file)[0]).lower()
            file_extension = os.path.splitext(file)[1].lower()
            if doc_name_normalized == file_name_normalized and file_extension in file_extensions:
                file_path = os.path.join(root, file)
                update_log(f"Found {doc_type} file: {file_path}")
                return file_path

    update_log(f"No matching file found for {doc_type} with name '{doc_name}'")
    return None

# Function to send a file via WhatsApp
def send_file(driver, phone_number, file_path, doc_type):
    try:
        if not os.path.exists(file_path):
            update_log(f"File does not exist: {file_path}")
            return

        # Open chat with the phone number
        driver.get(f"https://web.whatsapp.com/send?phone={phone_number}")
        update_log(f"Opened chat with {phone_number}")

        # Wait for and click the "Attach" button
        attach_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//button[@title="Attach"]'))
        )
        attach_btn.click()
        update_log(f"Clicked on 'Attach' button for {phone_number}")

        # Wait for the file input and upload the file
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
        )
        file_input.send_keys(file_path)
        update_log(f"Uploaded file '{os.path.basename(file_path)}' for {phone_number}")

        # Wait for the "Send" button and click
        send_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Send"]'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", send_btn)
        try:
            send_btn.click()
            update_log(f"Clicked 'Send' button for {phone_number}")
        except Exception as e:
            update_log(f"Standard click failed, attempting JavaScript click: {str(e)}")
            driver.execute_script("arguments[0].click();", send_btn)

        time.sleep(4)  # Ensure the file is sent
        update_log(f"File '{os.path.basename(file_path)}' from '{doc_type}' sent to {phone_number}")

    except Exception as e:
        update_log(f"Failed to send file '{os.path.basename(file_path)}' to {phone_number} - {str(e)}")

# Function to update the log
def update_log(message):
    log_text.insert(tk.END, f"{message}\n")
    log_text.see(tk.END)

# Function to browse and process Excel file
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        update_log(f"Selected file: {file_path}")
        Thread(target=send_files, args=(file_path,)).start()

# Initialize the Tkinter GUI
app = tk.Tk()
app.title("WhatsApp Document Sender")
app.geometry("600x550")

# Define flags for document type selection
selected_docs = {
    "LR Copy": tk.BooleanVar(app),
    "Receipt": tk.BooleanVar(app),
    "Invoice": tk.BooleanVar(app)
}

tk.Label(app, text="WhatsApp Document Sender", font=("Arial", 16)).pack(pady=10)

tk.Button(app, text="Select LR Copy Directory", command=lambda: select_directory('LR Copy')).pack(pady=5)
tk.Button(app, text="Select Receipt Directory", command=lambda: select_directory('Receipt')).pack(pady=5)
tk.Button(app, text="Select Invoice Directory", command=lambda: select_directory('Invoice')).pack(pady=5)

tk.Checkbutton(app, text="Include LR Copy", variable=selected_docs['LR Copy']).pack(pady=5)
tk.Checkbutton(app, text="Include Receipt", variable=selected_docs['Receipt']).pack(pady=5)
tk.Checkbutton(app, text="Include Invoice", variable=selected_docs['Invoice']).pack(pady=5)

tk.Button(app, text="Browse Excel File", command=browse_file).pack(pady=10)

log_text = tk.Text(app, height=15, width=70)
log_text.pack(pady=10)

app.mainloop()