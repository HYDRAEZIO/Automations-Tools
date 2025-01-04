import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to load the Excel file
def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            global df
            df = pd.read_excel(file_path)
            messagebox.showinfo("Success", "File loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file.\n{e}")

# Function to process the data
def process_data():
    try:
        global df
        # Replace NaN values with 0
        df = df.fillna(0)
        # Multiply numeric columns by 4 and ensure no decimals
        df = df.applymap(lambda x: int(x * 4) if isinstance(x, (int, float)) else x)
        messagebox.showinfo("Success", "Data processed successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process data.\n{e}")

# Function to save the modified data
def save_file():
    try:
        output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                        filetypes=[("Excel files", "*.xlsx")])
        if output_file_path:
            df.to_excel(output_file_path, index=False)
            messagebox.showinfo("Success", "File saved successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save file.\n{e}")

# Main application window
app = tk.Tk()
app.title("Excel Processor")
app.geometry("300x200")

# Create buttons for each action
load_btn = tk.Button(app, text="Load Excel File", command=load_file)
load_btn.pack(pady=10)

process_btn = tk.Button(app, text="Process Data", command=process_data)
process_btn.pack(pady=10)

save_btn = tk.Button(app, text="Save File", command=save_file)
save_btn.pack(pady=10)

# Run the application
app.mainloop()