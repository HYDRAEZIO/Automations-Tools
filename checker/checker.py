import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Button, Label

def select_excel_file():
    """Prompt the user to select an Excel file."""
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        excel_file_label.config(text=f"Selected: {file_path}")
        app_data["excel_file"] = file_path

def select_directory():
    """Prompt the user to select a directory."""
    dir_path = filedialog.askdirectory(title="Select Image Directory")
    if dir_path:
        directory_label.config(text=f"Selected: {dir_path}")
        app_data["image_directory"] = dir_path

def get_all_files(directory):
    """Recursively fetch all files in the directory and its subdirectories."""
    all_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            all_files.append(os.path.join(root, file))
    return all_files

def process_data():
    """Process the data to cross-check images."""
    excel_file = app_data.get("excel_file")
    image_directory = app_data.get("image_directory")
    
    if not excel_file or not image_directory:
        messagebox.showerror("Error", "Please select both an Excel file and an image directory.")
        return
    
    try:
        # Load Excel file
        data = pd.read_excel(excel_file)
        if "Product Code" not in data.columns:
            messagebox.showerror("Error", "The Excel file must have a 'Product Code' column.")
            return
        
        # Get list of all image files in subdirectories
        all_image_files = get_all_files(image_directory)
        image_files = {os.path.splitext(os.path.basename(f))[0].lower() for f in all_image_files}
        
        # Cross-check images
        data["Image Found"] = data["Product Code"].str.lower().isin(image_files)
        
        # Display the result and save to a new Excel file
        result_file = os.path.join(image_directory, "image_check_result.xlsx")
        data.to_excel(result_file, index=False)
        
        messagebox.showinfo("Success", f"Process completed. Results saved to:\n{result_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
# created by Vivek Pandey @ https://github.com/HYDRAEZIO
# GUI Setup
app_data = {}
root = tk.Tk()
root.title("Product Image Checker")

# Excel File Selection
Label(root, text="Step 1: Select Excel File").pack(pady=5)
Button(root, text="Choose Excel File", command=select_excel_file).pack()
excel_file_label = Label(root, text="No file selected")
excel_file_label.pack(pady=5)

# Directory Selection
Label(root, text="Step 2: Select Image Directory").pack(pady=5)
Button(root, text="Choose Directory", command=select_directory).pack()
directory_label = Label(root, text="No directory selected")
directory_label.pack(pady=5)

# Process Button
Button(root, text="Process Data", command=process_data).pack(pady=20)

root.mainloop()