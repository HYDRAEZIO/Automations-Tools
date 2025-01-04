import pandas as pd
from tkinter import filedialog, Tk, messagebox

# Function to compare two Excel sheets and fetch missing product codes
def compare_excel_sheets():
    # Hide the main Tkinter window
    root = Tk()
    root.withdraw()

    # Ask user to select the first Excel file (with product codes and names)
    file1_path = filedialog.askopenfilename(title="Select Excel File 1 (with Product Codes)",
                                            filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file1_path:
        messagebox.showerror("Error", "No Excel file selected for File 1.")
        return

    # Ask user to select the second Excel file (with product names only)
    file2_path = filedialog.askopenfilename(title="Select Excel File 2 (without Product Codes)",
                                            filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file2_path:
        messagebox.showerror("Error", "No Excel file selected for File 2.")
        return

    try:
        # Read the first Excel file into a DataFrame, treating 'Product Code' as string
        df1 = pd.read_excel(file1_path, dtype={'Product Code': str})
        
        # Check if necessary columns exist in File 1
        if 'Product Name' not in df1.columns or 'Product Code' not in df1.columns:
            messagebox.showerror("Error", "'Product Name' and 'Product Code' columns not found in File 1.")
            return

        # Read the second Excel file into a DataFrame
        df2 = pd.read_excel(file2_path)

        # Check if 'Product Name' column exists in File 2
        if 'Product Name' not in df2.columns:
            messagebox.showerror("Error", "'Product Name' column not found in File 2.")
            return

        # Ensure 'Product Code' is treated as a string and fill NaN with empty strings
        df1['Product Code'] = df1['Product Code'].fillna('').astype(str)

        # Normalize product names for case-insensitive and space-insensitive comparison
        df1['Product Name'] = df1['Product Name'].str.strip().str.lower()
        df2['Original Product Name'] = df2['Product Name']  # Preserve the original names
        df2['Product Name'] = df2['Product Name'].str.strip().str.lower()

        # Merge the two DataFrames on 'Product Name' to fetch the missing product codes
        df_merged = pd.merge(df2, df1[['Product Name', 'Product Code']], on='Product Name', how='left')

        # Check for product codes that are still missing and set them to empty strings
        df_merged['Product Code'] = df_merged['Product Code'].replace('nan', '').replace('None', '')

        # Identify renamed products
        renamed_products = df_merged[df_merged['Original Product Name'] != df_merged['Product Name']][
            ['Original Product Name', 'Product Name']]

        # Ask user to save the renamed products log
        log_file_path = filedialog.asksaveasfilename(title="Save Renamed Products Log",
                                                     defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx")])
        if log_file_path:
            renamed_products.to_excel(log_file_path, index=False)
            messagebox.showinfo("Success", f"Renamed products log saved at:\n{log_file_path}")
        else:
            messagebox.showinfo("Info", "No log file saved.")

        # Ask user to save the new Excel file
        output_file_path = filedialog.asksaveasfilename(title="Save Updated Excel File",
                                                        defaultextension=".xlsx",
                                                        filetypes=[("Excel files", "*.xlsx")])
        if output_file_path:
            df_merged.to_excel(output_file_path, index=False)
            messagebox.showinfo("Success", f"Updated Excel file saved at:\n{output_file_path}")
        else:
            messagebox.showerror("Error", "Invalid file extension. Please save as '.xlsx'.")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to process the Excel files: {str(e)}")

# Run the function to start the process
if __name__ == "__main__":
    compare_excel_sheets()