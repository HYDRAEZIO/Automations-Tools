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

# Function to detect categories based on the Excel sheet and the main directory
def detect_categories(file_path, main_dir):
    """
    Detect categories from the Excel file and match them with directories in the main folder.
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip().str.lower()

        # Extract unique category names from the Excel file
        detected_categories = [col for col in df.columns if col not in ['name', 'mobile number']]
        update_log(f"Detected categories from Excel: {detected_categories}")

        # Match categories with subdirectories in the main directory
        for category in detected_categories:
            sub_dir = os.path.join(main_dir, category)
            if os.path.isdir(sub_dir):
                directories[category] = sub_dir
                update_log(f"Matched directory for {category}: {sub_dir}")
            else:
                directories[category] = None
                update_log(f"No directory found for category: {category}")

        return True

    except Exception as e:
        update_log(f"Error detecting categories: {str(e)}")
        return False

# Function to send files based on the Excel file
def send_files(file_path, main_dir):
    """
    Process the Excel file and send files from the detected categories.
    """
    if not detect_categories(file_path, main_dir):
        update_log("Category detection failed. Aborting process.")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip().str.lower()

        # Replace NaN values in the mobile number column with empty strings
        df['mobile number'] = df['mobile number'].fillna("").astype(str)

        # Launch browser and load WhatsApp Web
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        driver.maximize_window()
        driver.get("https://web.whatsapp.com")
        messagebox.showinfo("QR Code Scan", "Please scan the QR code on WhatsApp Web and click OK to continue.")

        # Process each category one by one
        for category, dir_path in directories.items():
            if not dir_path:
                update_log(f"Skipping category {category} as no directory is found.")
                continue

            # Get all files in the category directory
            files = [
                os.path.join(dir_path, file)
                for file in os.listdir(dir_path)
                if os.path.isfile(os.path.join(dir_path, file)) and file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif'))
            ]

            if not files:
                update_log(f"No files found in directory for {category}")
                continue

            # Process each file for all contacts
            for file_path in files:
                update_log(f"Processing file {file_path} for category {category}")
                for _, row in df.iterrows():
                    phone_numbers = row['mobile number'].strip()

                    # Split multiple numbers (assume comma, space, or semicolon as delimiters)
                    numbers_list = re.split(r'[,\s;]+', phone_numbers)

                    for phone_number in numbers_list:
                        if phone_number:  # Ensure the phone number is not empty
                            send_file(driver, phone_number, file_path, category)

        messagebox.showinfo("Process Completed", "All selected files have been processed.")
        driver.quit()

    except Exception as e:
        update_log(f"Error: {str(e)}")

# Function to send an image via WhatsApp
def send_file(driver, phone_number, file_path, category):
    """
    Send a single file via WhatsApp Web.
    """
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

        # Wait for the file input under "Photos & Videos" and upload the file
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
        )
        file_input.send_keys(file_path)
        update_log(f"Uploaded image '{os.path.basename(file_path)}' for {phone_number}")

        # Wait for the caption input field and enter the caption
        caption = f"{category}: {os.path.basename(file_path)}"
        caption_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true" and @aria-placeholder="Add a caption"]'))
        )
        caption_field.send_keys(caption)
        update_log(f"Caption added: '{caption}'")

        # Wait for the "Send" button and click it
        send_btn = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Send"]'))
        )
        send_btn.click()

        time.sleep(4)  # Ensure the image is sent
        update_log(f"Image '{os.path.basename(file_path)}' with caption '{caption}' sent to {phone_number}")

    except Exception as e:
        update_log(f"Failed to send image '{os.path.basename(file_path)}' to {phone_number} - {str(e)}")

# Function to update the log
def update_log(message):
    """
    Update the log displayed in the GUI.
    """
    log_text.insert(tk.END, f"{message}\n")
    log_text.see(tk.END)

# Function to browse and process Excel file
def browse_file():
    """
    Open a file dialog to select the main directory and Excel file.
    """
    main_dir = filedialog.askdirectory(title="Select Main Directory")
    if not main_dir:
        update_log("Main directory not selected.")
        return

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        update_log(f"Selected file: {file_path}")
        Thread(target=send_files, args=(file_path, main_dir)).start()

# Initialize the Tkinter GUI
app = tk.Tk()
app.title("WhatsApp Product Sender")
app.geometry("600x550")

# Define directories for detected categories
directories = {}

tk.Label(app, text="WhatsApp Product Sender", font=("Arial", 16)).pack(pady=10)

# Add a button to browse the Excel file and directory
tk.Button(app, text="Browse Excel File and Directory", command=browse_file).pack(pady=10)

log_text = tk.Text(app, height=15, width=70)
log_text.pack(pady=10)

app.mainloop()
