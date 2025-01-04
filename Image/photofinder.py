import os
import pandas as pd
import cv2
import tkinter as tk
from tkinter import filedialog, messagebox
import re

# Supported image extensions
IMAGE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.gif']

# Function to clean and normalize strings for comparison
def clean_string(s):
    # Remove non-alphanumeric characters, convert to lowercase, and split into words
    return re.findall(r'\w+', s.lower())

# Function to get all image files in a directory recursively
def get_all_images(root_dir):
    all_images = {}
    for dirpath, _, filenames in os.walk(root_dir):
        for file in filenames:
            if os.path.splitext(file)[1].lower() in IMAGE_EXTENSIONS:
                all_images[file.lower()] = os.path.join(dirpath, file)
    return all_images
# created by Vivek Pandey @ https://github.com/HYDRAEZIO
# Function to fetch images based on the document list using Prd. Code
def fetch_images_by_Prd_Code():
    # Get document file path
    doc_file = filedialog.askopenfilename(filetypes=[("All Files", "*.xlsx *.xls *.csv *.txt")])
    if not doc_file:
        return

    # Read the document based on its extension
    file_extension = os.path.splitext(doc_file)[1].lower()

    try:
        if file_extension in ['.xlsx', '.xls']:
            df = pd.read_excel(doc_file)
        elif file_extension == '.csv':
            df = pd.read_csv(doc_file)
        elif file_extension == '.txt':
            df = pd.read_csv(doc_file, delimiter="\t", header=None, names=["Prd. Code"])
        else:
            messagebox.showerror("Error", "Unsupported file format.")
            return
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read the document: {str(e)}")
        return

    # Check if required column is present
    if 'Prd. Code' not in df.columns:
        messagebox.showerror("Error", "'Prd. Code' column not found in the document.")
        return

    # Get image directory
    image_directory = filedialog.askdirectory(title="Select Image Directory")
    if not image_directory:
        return

    # Get output directory
    output_directory = filedialog.askdirectory(title="Select Output Directory")
    if not output_directory:
        return

    # Normalize paths
    image_directory = os.path.normpath(image_directory)
    output_directory = os.path.normpath(output_directory)

    # Create output directory if it doesn't exist
    os.makedirs(output_directory, exist_ok=True)

    # Get all images in the main directory and its subdirectories
    available_images = get_all_images(image_directory)
    print(f"Available images in directory: {available_images}")

    # Track if any image is successfully copied
    copied_images = 0
    missing_images = []

    # Loop through the document list and try to find matching images by Prd. Code
    for index, row in df.iterrows():
        # Extract and clean Prd. Code
        raw_Prd_Code = str(row['Prd. Code']).strip().lower()
        Prd_Code = "".join(re.findall(r'\w+', raw_Prd_Code))  # Remove special characters and spaces
        found_image = None

        # Attempt to find matching image using substring matching
        for img_name, img_path in available_images.items():
            normalized_img_name = "".join(re.findall(r'\w+', img_name.lower()))  # Normalize file name
            if Prd_Code in normalized_img_name:  # Substring match
                found_image = img_path
                break

        if found_image:
            try:
                print(f"Processing image: {found_image}")

                # Read and save the image
                image = cv2.imread(found_image)
                if image is not None:
                    output_path = os.path.join(output_directory, os.path.basename(found_image))
                    cv2.imwrite(output_path, image)
                    copied_images += 1
                    print(f"Found and saved: {os.path.basename(found_image)}")
                else:
                    print(f"Failed to read image: {os.path.basename(found_image)}")
            except Exception as e:
                print(f"Error processing image {os.path.basename(found_image)}: {e}")
        else:
            missing_images.append(raw_Prd_Code)  # Use the raw code for reporting
            print(f"Image not found for: {raw_Prd_Code}")

    # Display result messages
    if copied_images > 0:
        messagebox.showinfo("Success", f"Successfully fetched and saved {copied_images} images!")

    if missing_images:
        # Show message box with missing Prd. Code
        missing_message = "\n".join(missing_images)
        messagebox.showwarning("Missing Images", f"Could not find images for the following Prd. Codes:\n{missing_message}")
    elif copied_images == 0:
        messagebox.showwarning("No Images Found", "No images were found or copied.")

# Create GUI Window
window = tk.Tk()
window.title("Image Fetcher")

# Set window size
window.geometry("400x200")

# Add button to fetch images
btn_fetch = tk.Button(window, text="Fetch Images", command=fetch_images_by_Prd_Code, width=20, height=2)
btn_fetch.pack(pady=50)

# Run the GUI loop
window.mainloop()