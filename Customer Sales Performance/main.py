import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading

# Function to categorize clients based on total sales
def categorize_client(total_sales):
    if total_sales < 500000:
        return 'Below 5 Lakh'
    elif total_sales < 1000000:
        return 'Below 10 Lakh'
    elif total_sales < 2000000:
        return 'Below 20 Lakh'
    elif total_sales < 3000000:
        return 'Below 30 Lakh'
    elif total_sales < 4000000:
        return 'Below 40 Lakh'
    elif total_sales < 5000000:
        return 'Below 50 Lakh'
    else:
        return 'Above 50 Lakh'

# Function to calculate the average of the best 3 months and the last 3 months
def calculate_avg_best_last(cust_sales):
    best_3_avg = sum(sorted(cust_sales, reverse=True)[:3]) / 3
    last_3_avg = sum(cust_sales[-3:]) / 3
    ratio = last_3_avg / best_3_avg if best_3_avg != 0 else float('inf')  # Avoid division by zero
    return best_3_avg, last_3_avg, ratio

# Function to create the line chart
def create_sales_chart(client_name, cust_sales, comp_sales, months, ratios, output_dir):
    plt.figure(figsize=(10, 7))

    # Plotting the customer sales data
    plt.plot(months, cust_sales, marker='o', color='blue', label='Customer Sales')

    # Add data labels on the line chart
    for idx, sales in enumerate(cust_sales):
        plt.text(idx, sales, f'{sales:.2f}', ha='center', va='bottom')

    # Set the x-axis labels to show the month name with the exact ratio in decimal form
    plt.xticks(ticks=range(len(months)), labels=ratios, rotation=45)

    # Title with added spacing
    plt.title(f'Sales Progress for {client_name} (Customer vs Company)', pad=40)

    # Grid
    plt.grid(True)

    # Save the plot as a PNG in the specified output directory
    safe_client_name = client_name.replace('/', '-')
    graph_png_path = os.path.join(output_dir, f'{safe_client_name}_sales_progress.png')
    plt.savefig(graph_png_path, format='png', bbox_inches='tight')
    plt.close()

    return graph_png_path

# Function to load large Excel files in chunks (by splitting sheets or using specified ranges)
def load_large_excel(file_path):
    # Load the Excel file in parts or chunks (if it contains multiple sheets)
    xls = pd.ExcelFile(file_path)
    all_data = pd.DataFrame()

    # Load each sheet one by one
    for sheet_name in xls.sheet_names:
        df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
        all_data = pd.concat([all_data, df_sheet], ignore_index=True)

    return all_data

# Function to generate broker-wise PDF reports
def generate_reports():
    global df, output_dir
    
    if df is None or output_dir == '':
        messagebox.showerror("Error", "Please select an Excel file and output directory.")
        return

    # Define the category order
    category_order = [
        'Below 5 Lakh',
        'Below 10 Lakh',
        'Below 20 Lakh',
        'Below 30 Lakh',
        'Below 40 Lakh',
        'Below 50 Lakh',
        'Above 50 Lakh'
    ]

    # Calculate total sales across all months for each client
    df['Total Cust Sales'] = df[['Cust Apr-2024', 'Cust May-2024', 'Cust Jun-2024', 
                                 'Cust Jul-2024', 'Cust Aug-2024', 'Cust Sep-2024']].sum(axis=1)
    df['Category'] = df['Total Cust Sales'].apply(categorize_client)
    df['Category'] = pd.Categorical(df['Category'], categories=category_order, ordered=True)

    brokers = df['Broker'].unique()

    progress_bar['maximum'] = len(brokers)
    progress_bar['value'] = 0

    for broker in brokers:
        broker_clients = df[df['Broker'] == broker].sort_values(by=['Category', 'Total Cust Sales'])

        if broker_clients.empty:
            continue

        pdf_graphs = FPDF(orientation='P', unit='mm', format='A4')
        pdf_graphs.set_auto_page_break(auto=True, margin=15)
        pdf_graphs.set_font("Arial", size=12)

        for category in category_order:
            clients = broker_clients[broker_clients['Category'] == category]

            if clients.empty:
                continue

            pdf_graphs.add_page()
            pdf_graphs.set_font('Arial', 'B', 20)
            pdf_graphs.cell(200, 20, f'Batch: {category}', ln=True, align='C')

            clients['Best 3 Avg'], clients['Last 3 Avg'], clients['Sales Ratio'] = zip(*clients.apply(
                lambda row: calculate_avg_best_last([
                    row['Cust Apr-2024'], row['Cust May-2024'], row['Cust Jun-2024'],
                    row['Cust Jul-2024'], row['Cust Aug-2024'], row['Cust Sep-2024']
                ]), axis=1))

            declining_clients = clients[clients['Sales Ratio'] > 1].sort_values(by='Sales Ratio', ascending=False)
            improving_clients = clients[clients['Sales Ratio'] <= 1].sort_values(by='Sales Ratio', ascending=True)

            sorted_clients = pd.concat([declining_clients, improving_clients])

            for i, row in sorted_clients.iterrows():
                client_name = row['A/c Name']
                cust_sales = [
                    row['Cust Apr-2024'], row['Cust May-2024'], row['Cust Jun-2024'],
                    row['Cust Jul-2024'], row['Cust Aug-2024'], row['Cust Sep-2024']
                ]
                comp_sales = [
                    row['Comp Apr-2024'], row['Comp May-2024'], row['Comp Jun-2024'],
                    row['Comp Jul-2024'], row['Comp Aug-2024'], row['Comp Sep-2024']
                ]
                months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep']

                graph_png_path = create_sales_chart(client_name, cust_sales, comp_sales, months, ratios=[
                    f'{month} ({cust/comp:.6f})' if comp != 0 else f'{month} (0)'
                    for month, cust, comp in zip(months, cust_sales, comp_sales)
                ], output_dir=output_dir)

                pdf_graphs.add_page()
                pdf_graphs.set_font('Arial', 'B', 16)
                pdf_graphs.cell(200, 10, f'Sales Progress Graph for {client_name}', ln=True, align='C')

                pdf_graphs.image(graph_png_path, x=10, y=30, w=180)
                pdf_graphs.ln(160)
                pdf_graphs.set_font('Arial', size=12)
                pdf_graphs.cell(0, 20, f"Client: {client_name}", ln=True, align='C')
                pdf_graphs.cell(0, 20, f"Best 3 Months Avg: {row['Best 3 Avg']:.2f}", ln=True, align='C')
                pdf_graphs.cell(0, 20, f"Last 3 Months Avg: {row['Last 3 Avg']:.2f}", ln=True, align='C')
                pdf_graphs.cell(0, 20, f"Sales Ratio (Last/Best): {row['Sales Ratio']:.2f}", ln=True, align='C')

                if row['Sales Ratio'] < 1:
                    pdf_graphs.set_text_color(255, 0, 0)
                    pdf_graphs.cell(0, 20, "Performance is declining!", ln=True, align='C')
                    pdf_graphs.set_text_color(0, 0, 0)

        pdf_output_path = os.path.join(output_dir, f'{broker}_client_sales_analysis.pdf')
        pdf_graphs.output(pdf_output_path)
        progress_bar['value'] += 1
        root.update_idletasks()

    messagebox.showinfo("Success", "Broker-wise PDF reports generated successfully.")

# Function to load Excel file
def load_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            df = load_large_excel(file_path)
            lbl_file.config(text=f"File Loaded: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

# Function to set output directory
def set_output_directory():
    global output_dir
    output_dir = filedialog.askdirectory()
    if output_dir:
        lbl_output_dir.config(text=f"Output Directory: {output_dir}")

# Function to start report generation in a separate thread
def start_report_generation():
    thread = threading.Thread(target=generate_reports)
    thread.start()

# Initialize Tkinter window
root = tk.Tk()
root.title("Broker-wise Sales Analysis Application")

df = None
output_dir = ''

# Create GUI elements
frm_controls = ttk.Frame(root, padding="10")
frm_controls.grid(row=0, column=0, sticky="ew")

btn_load_file = ttk.Button(frm_controls, text="Load Excel File", command=load_file)
btn_load_file.grid(row=0, column=0, padx=5, pady=5)

lbl_file = ttk.Label(frm_controls, text="No file loaded")
lbl_file.grid(row=0, column=1, padx=5, pady=5)

btn_set_output = ttk.Button(frm_controls, text="Set Output Directory", command=set_output_directory)
btn_set_output.grid(row=1, column=0, padx=5, pady=5)

lbl_output_dir = ttk.Label(frm_controls, text="No output directory set")
lbl_output_dir.grid(row=1, column=1, padx=5, pady=5)

btn_generate = ttk.Button(frm_controls, text="Generate Reports", command=start_report_generation)
btn_generate.grid(row=2, column=0, columnspan=2, pady=10)

progress_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=300)
progress_bar.grid(row=3, column=0, pady=10)

# Run the application
root.mainloop()