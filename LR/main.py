import pandas as pd
import os
import re
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from threading import Thread
import glob

# Define the base directory for Receipt
receipt_directory = ""

# Function to select the directory for receipts
def select_receipt_directory():
    global receipt_directory
    dir_path = filedialog.askdirectory()
    if dir_path:
        receipt_directory = dir_path
        update_log(f"Selected Receipt directory: {dir_path}")

# Function to send receipts via WhatsApp using Selenium
def send_receipts(file_path):
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Normalize column names
        df.columns = df.columns.str.strip().str.lower()

        # Check for necessary columns
        if 'sms no' not in df.columns or 'vch no.' not in df.columns:
            update_log("Missing required columns: 'SMS No' or 'Vch No.'")
            return

        # Launch the browser and open WhatsApp Web
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        driver.get("https://web.whatsapp.com")

        # Wait for the user to scan the QR code
        messagebox.showinfo("QR Code Scan", "Please scan the QR code on WhatsApp Web and click OK to continue.")

        for index, row in df.iterrows():
            phone_numbers = row['sms no'].split(",")  # Handle multiple phone numbers
            vch_no = str(int(row['vch no.'])) if pd.notna(row['vch no.']) else None

            for phone_number in phone_numbers:
                phone_number = phone_number.strip()

                # Attempt to send the Receipt image using `Vch No.`
                if vch_no:
                    receipt_image = search_image_file(vch_no)
                    if receipt_image:
                        send_image(driver, phone_number, receipt_image)

        messagebox.showinfo("Process Completed", "All receipt images have been sent.")
        driver.quit()

    except Exception as e:
        update_log(f"Error: {str(e)}")

# Helper function to search for an image file in the directory, including subdirectories and handling network paths
def search_image_file(doc_name):
    if not receipt_directory:
        update_log("No directory selected for Receipt")
        return None

    # List of common image extensions (case-insensitive)
    image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.gif']

    # Normalize the document name (remove spaces, convert to lowercase)
    doc_name_normalized = re.sub(r'\s+', '', doc_name).lower()

    # Traverse the directory and its subdirectories
    for root, dirs, files in os.walk(receipt_directory):
        for file in files:
            # Normalize the filename (remove spaces, convert to lowercase)
            file_name_normalized = re.sub(r'\s+', '', os.path.splitext(file)[0]).lower()
            file_extension = os.path.splitext(file)[1].lower()

            # Check if the normalized file name matches the document name and if the extension is an image format
            if doc_name_normalized == file_name_normalized and file_extension in image_extensions:
                file_path = os.path.join(root, file)
                update_log(f"Found Receipt image: {file_path}")
                return file_path

    update_log(f"No matching image file found for Receipt with name '{doc_name}'")
    return None

# Helper function to send an image via Selenium
def send_image(driver, phone_number, image_path):
    try:
        # Open chat for the specified phone number
        driver.get(f"https://web.whatsapp.com/send?phone={phone_number}")
        time.sleep(10)  # Wait for the chat to load

        # Locate and click the attachment button
        attach_btn = driver.find_element(By.XPATH, '//div[@title="Attach"]')
        attach_btn.click()
        time.sleep(2)

        # Locate and send the image file
        file_input = driver.find_element(By.XPATH, '//input[@accept="*"]')
        file_input.send_keys(image_path)
        time.sleep(2)

        # Click the send button
        send_btn = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
        send_btn.click()
        time.sleep(5)  # Wait for the image to be sent

        update_log(f"Receipt image '{os.path.basename(image_path)}' sent to {phone_number}")
    except Exception as e:
        update_log(f"Failed to send image '{os.path.basename(image_path)}' to {phone_number} - {str(e)}")

# Function to update the log
def update_log(message):
    log_text.insert(tk.END, f"{message}\n")
    log_text.see(tk.END)

# Function to open file dialog and start processing
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        update_log(f"Selected file: {file_path}")
        # Run the file processing in a separate thread
        Thread(target=send_receipts, args=(file_path,)).start()
# created by Vivek Pandey @ https://github.com/HYDRAEZIO
# Setting up the Tkinter GUI
app = tk.Tk()
app.title("WhatsApp Receipt Image Sender")
app.geometry("600x450")

# Title Label
tk.Label(app, text="WhatsApp Receipt Image Sender", font=("Arial", 16)).pack(pady=10)

# Directory Selection Button
tk.Button(app, text="Select Receipt Image Directory", command=select_receipt_directory).pack(pady=10)

# Browse Excel File Button
tk.Button(app, text="Browse Excel File", command=browse_file).pack(pady=10)

# Log Textbox
log_text = tk.Text(app, height=15, width=70)
log_text.pack(pady=10)

# Start the GUI loop
app.mainloop()