import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def select_stock_file():
    """Open a file dialog to select the stock Excel file."""
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        entry_stock.delete(0, tk.END)
        entry_stock.insert(0, file_path)

def select_container_files():
    """Open a file dialog to select multiple container Excel files."""
    file_paths = filedialog.askopenfilenames(
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_paths:
        entry_containers.delete(0, tk.END)
        entry_containers.insert(0, ";".join(file_paths))

def merge_and_save():
    """Merge stock file with multiple container files and save results."""
    stock_file = entry_stock.get()
    container_files = entry_containers.get().split(";")
    save_directory = filedialog.askdirectory(title="Select Directory to Save Results")

    if not stock_file or not container_files or not save_directory:
        messagebox.showerror("Error", "Please select the stock file, container files, and save directory.")
        return

    try:
        # Read the stock file
        df_stock = pd.read_excel(stock_file, sheet_name=0)
        df_stock.rename(columns={'Prd Code': 'prd_code'}, inplace=True)
        df_stock['prd_code'] = df_stock['prd_code'].astype(str).str.strip().str.lower()

        for container_file in container_files:
            # Extract container number from the file name
            container_number = os.path.splitext(os.path.basename(container_file))[0]

            # Read the container file
            df_container = pd.read_excel(container_file, sheet_name=0)
            df_container.rename(columns={'prd_code': 'prd_code'}, inplace=True)
            df_container['prd_code'] = df_container['prd_code'].astype(str).str.strip().str.lower()

            # Merge stock and container files
            df_merged = pd.merge(
                df_container,  # Container List (left table)
                df_stock,  # Stock List (right table)
                on='prd_code', 
                how='left'
            )

            # Replace NaN with 0 for unmatched rows
            df_merged.fillna(0, inplace=True)

            # Save the merged result with the container number as the file name
            output_file = os.path.join(save_directory, f"{container_number}_merged.xlsx")
            df_merged.to_excel(output_file, index=False)

        messagebox.showinfo("Success", f"All container files have been processed and saved in {save_directory}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the main window
root = tk.Tk()
root.title("Excel File Merger")

# Stock file selection
label_stock = tk.Label(root, text="Select Stock File (Excel):")
label_stock.pack(pady=5)

entry_stock = tk.Entry(root, width=50)
entry_stock.pack(pady=5)

btn_browse_stock = tk.Button(root, text="Browse", command=select_stock_file)
btn_browse_stock.pack(pady=5)

# Container files selection
label_containers = tk.Label(root, text="Select Container Files (Excel):")
label_containers.pack(pady=5)

entry_containers = tk.Entry(root, width=50)
entry_containers.pack(pady=5)

btn_browse_containers = tk.Button(root, text="Browse", command=select_container_files)
btn_browse_containers.pack(pady=5)

# Merge and save button
btn_merge_save = tk.Button(root, text="Merge and Save", command=merge_and_save, bg='lightblue')
btn_merge_save.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()