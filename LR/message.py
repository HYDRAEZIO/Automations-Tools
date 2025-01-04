import pandas as pd
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from threading import Thread

# Construct the English message with proper formatting
english_message = (
    "Dear Customer, The LR copy is now available on the Aadinath App; log in to view/download it and click on the link mentioned below to watch the access guide. \n Android : https://play.google.com/store/apps/details?id=app.sam \n Apple: https://apps.apple.com/in/app/aadinath-exports/id6476626373 \n Video Link :https://tinyurl.com/LRDEMOVIDEOs")

# Construct the Hindi message with proper formatting
hindi_message = (
   " प्रिय ग्राहक, LR कॉपी अब आदिनाथ ऐप पर उपलब्ध है; इसे देखने/डाउनलोड करने के लिए लॉगिन करें और एक्सेस गाइड देखने के लिए नीचे दिए गए लिंक पर क्लिक करें।\n Android : https://play.google.com/store/apps/details?id=app.sam \n Apple: https://apps.apple.com/in/app/aadinath-exports/id6476626373 \n वीडियो लिंक :https://tinyurl.com/LRDEMOVIDEOs )"
)

# Function to send both English and Hindi messages
def send_both_messages(driver, phone_number):
    try:
        # Send the English message
        send_message(driver, phone_number, english_message)
        time.sleep(2)  # Short delay between messages

        # Send the Hindi message
        send_message(driver, phone_number, hindi_message)
        time.sleep(2)
    except Exception as e:
        update_log(f"Failed to send messages to {phone_number} - {str(e)}")

# Helper function to send a message via Selenium
def send_message(driver, phone_number, message):
    try:
        driver.get(f"https://web.whatsapp.com/send?phone={phone_number}")
        time.sleep(10)  # Wait for the chat to load

        # XPath for the message input box
        input_box_xpath = '//div[@contenteditable="true" and @data-tab="10"]'

        # Wait for the message input box to appear
        message_box = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, input_box_xpath))
        )

        # Click the input box, type the message, and send it
        message_box.click()
        message_box.send_keys(message)
        time.sleep(1)
        message_box.send_keys(Keys.ENTER)
        update_log(f"Message sent to {phone_number}")
    except Exception as e:
        update_log(f"Failed to send message to {phone_number} - {str(e)}")

# Function to send messages to all contacts in the Excel file
def send_messages(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip().str.lower()

        # Check if necessary columns are present
        if 'a/c name' not in df.columns or 'sms mobileno' not in df.columns:
            update_log("Missing required columns: 'A/c Name' or 'SMS Mobileno'")
            return

        # Launch the browser and open WhatsApp Web
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        driver.get("https://web.whatsapp.com")
        messagebox.showinfo("QR Code Scan", "Please scan the QR code on WhatsApp Web and click OK to continue.")

        for index, row in df.iterrows():
        # Correctly format the phone number
            phone_number = f"+91{str(row['sms mobileno']).split('.')[0].strip()}"
            name = row['a/c name'].strip() if pd.notna(row['a/c name']) else "ग्राहक"

            # Log the current phone number and client name
            update_log(f"Sending messages to {phone_number} ({name})")
            send_both_messages(driver, phone_number)

        messagebox.showinfo("Process Completed", "All messages have been sent.")
        driver.quit()

    except Exception as e:
        update_log(f"Error: {str(e)}")

# Function to update the log in the GUI
def update_log(message):
    log_text.insert(tk.END, f"{message}\n")
    log_text.see(tk.END)

# Function to browse and select the Excel file
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        update_log(f"Selected file: {file_path}")
        Thread(target=send_messages, args=(file_path,)).start()

# Tkinter GUI Setup
app = tk.Tk()
app.title("WhatsApp Message Sender")
app.geometry("600x400")

# Title Label
title_label = tk.Label(app, text="WhatsApp Message Sender", font=("Arial", 16))
title_label.pack(pady=10)

# Browse Excel File Button
browse_button = tk.Button(app, text="Browse Excel File", command=browse_file, font=("Arial", 12))
browse_button.pack(pady=10)

# Log Textbox
log_text = tk.Text(app, height=15, width=70)
log_text.pack(pady=10)

# Start the Tkinter GUI loop
app.mainloop()