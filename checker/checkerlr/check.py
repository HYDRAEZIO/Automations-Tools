import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def load_excel():
    global excel_file
    excel_file = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if excel_file:
        excel_label.config(text=f"Excel File: {excel_file}")

def select_folder():
    global pdf_directory
    pdf_directory = filedialog.askdirectory(title="Select PDF Folder")
    if pdf_directory:
        folder_label.config(text=f"Folder: {pdf_directory}")

def cross_check():
    if not excel_file or not pdf_directory:
        messagebox.showerror("Error", "Please select both an Excel file and a PDF folder.")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(excel_file)

        # Assume numbers are in the first column and extract the number part
        df['Number'] = df.iloc[:, 0].apply(lambda x: str(x).split()[0])

        # List all PDF files in the directory
        pdf_files = {file.split('.')[0] for file in os.listdir(pdf_directory) if file.endswith('.pdf')}

        # Check each number if a corresponding PDF exists
        df['PDF Exists'] = df['Number'].apply(lambda x: 'Yes' if x in pdf_files else 'No')

        # Display results in the treeview
        for i in tree.get_children():
            tree.delete(i)
        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

        # Collect numbers for which PDFs were not found
        missing_pdfs = df[df['PDF Exists'] == 'No']['Number'].tolist()
        if missing_pdfs:
            messagebox.showinfo("Missing PDFs", "PDFs not found for the following numbers:\n" + "\n".join(missing_pdfs))
        else:
            messagebox.showinfo("Result", "All PDFs are accounted for.")

    except FileNotFoundError:
        messagebox.showerror("Error", "The specified file or directory was not found.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Setup the GUI
root = tk.Tk()
root.title("PDF Cross-Checker")
root.geometry("600x400")

# Variables
excel_file = ''
pdf_directory = ''

# Layout
load_excel_btn = tk.Button(root, text="Load Excel File", command=load_excel)
load_excel_btn.pack(pady=(20, 5))

excel_label = tk.Label(root, text="No Excel file selected")
excel_label.pack()

select_folder_btn = tk.Button(root, text="Select PDF Folder", command=select_folder)
select_folder_btn.pack(pady=5)

folder_label = tk.Label(root, text="No Folder Selected")
folder_label.pack()

cross_check_btn = tk.Button(root, text="Cross Check Files", command=cross_check)
cross_check_btn.pack(pady=(10, 20))

tree = ttk.Treeview(root, columns=('Number', 'PDF Exists'), show='headings')
tree.heading('Number', text='Number')
tree.heading('PDF Exists', text='PDF Exists')
tree.pack(expand=True, fill='both')

root.mainloop()