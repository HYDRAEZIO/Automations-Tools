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

# Function to select directory for product files
def select_directory(product):
    dir_path = filedialog.askdirectory()
    if dir_path:
        directories[product] = dir_path
        update_log(f"Selected {product} directory: {dir_path}")

# Function to send files based on the Excel file
def send_files(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip().str.lower()

        # Verify required columns
        required_columns = ['name', 'mobile number']
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
        for _, row in df.iterrows():
            phone_number = str(row['mobile number']).strip()
            if not phone_number.isdigit():
                continue

            # Loop through product categories and send relevant files
            for product, dir_path in directories.items():
                if pd.notna(row.get(product.lower())) and row[product.lower()] == "YES":
                    product_file = search_file(product, phone_number)
                    if product_file:
                        send_file(driver, phone_number, product_file, product)

        messagebox.showinfo("Process Completed", "All selected files have been processed.")
        driver.quit()

    except Exception as e:
        update_log(f"Error: {str(e)}")
# created by Vivek Pandey @ https://github.com/HYDRAEZIO
# File search function
def search_file(product, phone_number):
    dir_path = directories.get(product)
    if not dir_path:
        update_log(f"No directory selected for {product}")
        return None

    file_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.pdf', '.docx', '.xlsx', '.xls', '.txt']
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            file_extension = os.path.splitext(file)[1].lower()
            if file_extension in file_extensions:
                return os.path.join(root, file)

    update_log(f"No matching file found for {product}")
    return None

# Function to send a file via WhatsApp
def send_file(driver, phone_number, file_path, product):
    try:
        if not os.path.exists(file_path):
            update_log(f"File does not exist: {file_path}")
            return

        # Open chat with the phone number
        driver.get(f"https://web.whatsapp.com/send?phone={phone_number}")
        update_log(f"Opened chat with {phone_number}")

        # Wait for and click the "Attach" button
        attach_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div[2]/button'))
        )
        attach_btn.click()

        # Wait for the file input and upload the file
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
        )
        file_input.send_keys(file_path)
        update_log(f"Uploaded file '{os.path.basename(file_path)}' for {phone_number}")

        # Wait for the "Send" button and click
        send_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div[3]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div'))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", send_btn)
        send_btn.click()

        time.sleep(4)  # Ensure the file is sent
        update_log(f"File '{os.path.basename(file_path)}' for '{product}' sent to {phone_number}")

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
app.title("WhatsApp Product Sender")
app.geometry("600x550")

# Define directories for each product
directories = {
    "Toys": "",
    "School Stationery": "",
    "Office Stationery": "",
    "Scissors": "",
    "Hardware": "",
    "Fans": ""
}

tk.Label(app, text="WhatsApp Product Sender", font=("Arial", 16)).pack(pady=10)

# Add buttons for selecting directories
for product in directories.keys():
    tk.Button(app, text=f"Select {product} Directory", command=lambda p=product: select_directory(p)).pack(pady=5)

tk.Button(app, text="Browse Excel File", command=browse_file).pack(pady=10)

log_text = tk.Text(app, height=15, width=70)
log_text.pack(pady=10)

app.mainloop()