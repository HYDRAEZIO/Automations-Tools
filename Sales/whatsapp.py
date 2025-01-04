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

# Base directory for storing images
base_directory = filedialog.askdirectory(title="Select Base Directory for Images")

# Category list (initialized empty for dynamic addition)
categories = []

# Progress tracking variables
total_tasks = 0
completed_tasks = 0

# Function to create a new category
def add_category():
    category_name = category_entry.get().strip()
    if category_name and category_name not in categories:
        categories.append(category_name)
        category_dir = os.path.join(base_directory, category_name)
        os.makedirs(category_dir, exist_ok=True)
        update_log(f"Added new category: {category_name} (Directory created: {category_dir})")
        category_entry.delete(0, tk.END)
    else:
        update_log(f"Category '{category_name}' already exists or is invalid.")

# Function to get image files for a category
def get_image_files(category):
    category_dir = os.path.join(base_directory, category)
    if not os.path.exists(category_dir):
        update_log(f"Directory for {category} not found.")
        return []
    return [os.path.join(category_dir, file) for file in os.listdir(category_dir) if file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif'))]

# Function to send images to all clients for a specific category
def send_images_to_clients(driver, phone_numbers, image_files, category):
    global completed_tasks, total_tasks

    if not image_files:
        update_log(f"No images found for {category}.")
        return

    for image_file in image_files:
        for phone_number in phone_numbers:
            try:
                # Open chat for the specified phone number
                driver.get(f"https://web.whatsapp.com/send?phone={phone_number}")
                time.sleep(10)  # Wait for the chat to load

                # Locate and click the attachment button
                attach_btn = driver.find_element(By.XPATH, '//div[@title="Attach"]')
                attach_btn.click()
                time.sleep(2)

                # Locate and send the file
                file_input = driver.find_element(By.XPATH, '//input[@accept="*"]')
                file_input.send_keys(image_file)
                time.sleep(2)

                # Click the send button
                send_btn = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
                send_btn.click()
                time.sleep(5)  # Wait for the file to be sent

                update_log(f"Sent image '{os.path.basename(image_file)}' to {phone_number} for category '{category}'")

                # Update progress
                completed_tasks += 1
                update_progress()

            except Exception as e:
                update_log(f"Failed to send image '{os.path.basename(image_file)}' to {phone_number} - {str(e)}")

# Function to update progress in the GUI
def update_progress():
    progress_percent = (completed_tasks / total_tasks) * 100
    progress_var.set(f"Progress: {completed_tasks}/{total_tasks} ({progress_percent:.2f}%)")
    app.update_idletasks()

# Function to send images to all clients based on the Excel file
def send_files(file_path):
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Normalize column names
        df.columns = df.columns.str.strip().str.lower()

        # Check for necessary columns
        required_columns = ['name', 'mobile number']
        for col in required_columns:
            if col not in df.columns:
                update_log(f"Missing column: {col}")
                return

        # Extract phone numbers
        phone_numbers = df['mobile number'].dropna().astype(str).tolist()

        # Launch the browser and open WhatsApp Web
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        driver.get("https://web.whatsapp.com")

        # Wait for the user to scan the QR code
        messagebox.showinfo("QR Code Scan", "Please scan the QR code on WhatsApp Web and click OK to continue.")

        # Calculate total tasks for progress tracking
        global total_tasks, completed_tasks
        total_tasks = sum(len(get_image_files(category)) for category in categories) * len(phone_numbers)
        completed_tasks = 0

        # Iterate through each category and send images to all clients
        for category in categories:
            image_files = get_image_files(category)
            send_images_to_clients(driver, phone_numbers, image_files, category)

        messagebox.showinfo("Process Completed", "All images have been sent to clients.")
        driver.quit()

    except Exception as e:
        update_log(f"Error: {str(e)}")

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
        Thread(target=send_files, args=(file_path,)).start()

# Setting up the Tkinter GUI
app = tk.Tk()
app.title("WhatsApp Image Sender with Category Management")
app.geometry("600x650")

# Progress tracking variable
progress_var = tk.StringVar(app)
progress_var.set("Progress: 0/0 (0.00%)")

# GUI Components
tk.Label(app, text="WhatsApp Image Sender", font=("Arial", 16)).pack(pady=10)

# Add category section
tk.Label(app, text="Add New Category:", font=("Arial", 12)).pack(pady=5)
category_entry = tk.Entry(app, width=30)
category_entry.pack(pady=5)
tk.Button(app, text="Add Category", command=add_category).pack(pady=5)

# Display current categories
tk.Label(app, text="Current Categories:", font=("Arial", 12)).pack(pady=5)
category_list = tk.Text(app, height=5, width=50)
category_list.pack(pady=5)

# Button to browse Excel file
tk.Button(app, text="Browse Excel File", command=browse_file).pack(pady=10)

# Progress label
progress_label = tk.Label(app, textvariable=progress_var, font=("Arial", 12))
progress_label.pack(pady=10)

# Log text box
log_text = tk.Text(app, height=15, width=70)
log_text.pack(pady=10)

# Update the category list in the GUI
def update_category_list():
    category_list.delete(1.0, tk.END)
    for category in categories:
        category_list.insert(tk.END, f"- {category}\n")

# Hook the add_category function to update the list
def add_category_hook():
    add_category()
    update_category_list()

# Override the Add Category button with the hook function
tk.Button(app, text="Add Category", command=add_category_hook).pack(pady=5)

# Initialize with any existing categories found in the base directory
for folder in os.listdir(base_directory):
    if os.path.isdir(os.path.join(base_directory, folder)):
        categories.append(folder)

# Update the category list initially
update_category_list()

app.mainloop()