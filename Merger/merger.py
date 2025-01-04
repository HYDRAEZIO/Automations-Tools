import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def select_file(sheet_num):
    """Open a file dialog to select an Excel file."""
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if file_path:
        if sheet_num == 1:
            entry_sheet1.delete(0, tk.END)
            entry_sheet1.insert(0, file_path)
        elif sheet_num == 2:
            entry_sheet2.delete(0, tk.END)
            entry_sheet2.insert(0, file_path)

def merge_files():
    """Merge two selected Excel files based on 'Prd Code' only."""
    sheet1_path = entry_sheet1.get()
    sheet2_path = entry_sheet2.get()

    if not sheet1_path or not sheet2_path:
        messagebox.showerror("Error", "Please select both Excel files.")
        return

    try:
        # Read Sheet 1 and Sheet 2
        df_sheet1 = pd.read_excel(sheet1_path, sheet_name=0)
        df_sheet2 = pd.read_excel(sheet2_path, sheet_name=0)

        # Trim spaces from column names in both DataFrames
        df_sheet1.columns = df_sheet1.columns.str.strip()
        df_sheet2.columns = df_sheet2.columns.str.strip()

        # Rename columns for merging consistency in Sheet 1
        df_sheet1.rename(columns={'Prd. Code': 'Prd Code'}, inplace=True)

        # Standardize case and strip spaces for consistency in merge columns
        df_sheet1['Prd Code'] = df_sheet1['Prd Code'].astype(str).str.strip().str.lower()
        df_sheet2['Prd Code'] = df_sheet2['Prd Code'].astype(str).str.strip().str.lower()

        # Create backup columns for Sheet 2's 'Prd Code'
        df_sheet2.rename(columns={'Prd Code': 'Sheet2 Prd Code'}, inplace=True)

        # Merge the two sheets based on 'Prd Code' only
        df_merged = pd.merge(
            df_sheet1, df_sheet2, 
            left_on='Prd Code',
            right_on='Sheet2 Prd Code', 
            how='left'
        )

        # Ask user where to save the merged output
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
            title="Save Merged Output"
        )

        if output_file:
            df_merged.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Merged output saved to {output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the main window
root = tk.Tk()
root.title("Excel File Merger")

# Sheet 1 file selection
label_sheet1 = tk.Label(root, text="Select Sheet 1 (Excel File):")
label_sheet1.pack(pady=5)

entry_sheet1 = tk.Entry(root, width=50)
entry_sheet1.pack(pady=5)

btn_browse1 = tk.Button(root, text="Browse", command=lambda: select_file(1))
btn_browse1.pack(pady=5)

# Sheet 2 file selection
label_sheet2 = tk.Label(root, text="Select Sheet 2 (Excel File):")
label_sheet2.pack(pady=5)

entry_sheet2 = tk.Entry(root, width=50)
entry_sheet2.pack(pady=5)

btn_browse2 = tk.Button(root, text="Browse", command=lambda: select_file(2))
btn_browse2.pack(pady=5)

# Merge button
btn_merge = tk.Button(root, text="Merge Files", command=merge_files, bg='lightblue')
btn_merge.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()