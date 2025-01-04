import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def browse_file1():
    global file1_path
    file1_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    file1_label.config(text=os.path.basename(file1_path))

def browse_file2():
    global file2_path
    file2_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    file2_label.config(text=os.path.basename(file2_path))

def merge_files():
    if not file1_path or not file2_path:
        messagebox.showerror("Error", "Please select both Excel files.")
        return
# created by Vivek Pandey @ https://github.com/HYDRAEZIO
    try:
        # Load Excel files
        file1 = pd.ExcelFile(file1_path)
        file2 = pd.ExcelFile(file2_path)
        
        # Load the first sheets
        file1_data = file1.parse(file1.sheet_names[0])
        file2_data = file2.parse(file2.sheet_names[0])

        # Default columns for merging
        merge_column = "A/c Code"
        number_column1 = "mobile number"
        number_column2 = "SMS Mobileno"

        # Check if columns exist
        if merge_column not in file1_data.columns or merge_column not in file2_data.columns:
            messagebox.showerror("Error", f"Merge column '{merge_column}' not found. Please check the data.")
            return
        if number_column1 not in file1_data.columns or number_column2 not in file2_data.columns:
            messagebox.showerror("Error", f"Number columns '{number_column1}' or '{number_column2}' not found. Please check the data.")
            return

        # Standardize mobile numbers for validation
        file1_data[number_column1] = file1_data[number_column1].astype(str).str.replace(r'\s+', '', regex=True)
        file2_data[number_column2] = file2_data[number_column2].astype(str).str.replace(r'\s+', '', regex=True)

        # Rename columns to differentiate A/c Codes
        file1_data.rename(columns={merge_column: "A/c Code (File 1)"}, inplace=True)
        file2_data.rename(columns={merge_column: "A/c Code (File 2)"}, inplace=True)

        # Perform the merge based on A/c Code
        merged_data = pd.merge(file1_data, file2_data, how='inner', left_on="A/c Code (File 1)", right_on="A/c Code (File 2)")

        # Add a validation column for mobile number matching
        merged_data['Number Match'] = merged_data.apply(
            lambda row: "Matched" if row[number_column1] == row[number_column2] else "Not Matched", axis=1
        )

        # Save the merged file
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_file:
            merged_data.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Merged file saved as {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Initialize GUI
root = tk.Tk()
root.title("Excel Sheet Merger")
root.geometry("500x300")

# File Selection
tk.Label(root, text="Select File 1:").pack()
file1_label = tk.Label(root, text="No file selected", fg="gray")
file1_label.pack()
tk.Button(root, text="Browse", command=browse_file1).pack()

tk.Label(root, text="Select File 2:").pack()
file2_label = tk.Label(root, text="No file selected", fg="gray")
file2_label.pack()
tk.Button(root, text="Browse", command=browse_file2).pack()

# Merge Button
tk.Button(root, text="Merge Files", command=merge_files).pack(pady=20)

# Run GUI
root.mainloop()
