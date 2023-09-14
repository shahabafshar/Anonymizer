import pandas as pd
import numpy as np
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook

def get_decimal_points(value):
    """Return the number of decimal points in a number."""
    try:
        return len(str(value).split(".")[1])
    except IndexError:
        return 0

def randomize_data(df, formula_cells):
    min_rows = int(0.5 * len(df))
    max_rows = int(1.5 * len(df))
    num_rows = np.random.randint(min_rows, max_rows)

    # Random duplication or omission of rows
    rows = df.sample(n=num_rows, replace=True)

    # Identify numeric columns and randomize them
    for col in df.select_dtypes(include=[np.number]).columns:
        if col not in formula_cells:  # Preserve formula cells
            range_val = df[col].mean() * 0.05
            rows[col] = rows[col].apply(lambda x: round(x + np.random.uniform(-range_val, range_val), get_decimal_points(x)))

    return rows

def choose_and_anonymize():
    file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx;*.xls")])

    if not file_path:
        return

    # Load formulas to ensure they're preserved
    formula_cells = {}
    wb = load_workbook(file_path)
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "=" in cell.value:  # Check for formula cells
                    if sheet_name not in formula_cells:
                        formula_cells[sheet_name] = {}
                    formula_cells[sheet_name][cell.coordinate] = cell.value
    del wb  # Close the workbook

    df = pd.read_excel(file_path)
    df_anon = randomize_data(df, formula_cells)

    directory, file_name = os.path.split(file_path)
    base_name, ext = os.path.splitext(file_name)
    output_path = os.path.join(directory, base_name + "-anon" + ext)

    try:
        df_anon.to_excel(output_path, index=False)
        
        # Reload formulas to the anonymized file
        wb = load_workbook(output_path)
        for sheet_name, cells in formula_cells.items():
            sheet = wb[sheet_name]
            for cell, formula in cells.items():
                sheet[cell] = formula
        wb.save(output_path)
        del wb  # Close the workbook

        messagebox.showinfo("Success", f"Anonymized data saved to: {output_path}")
    except PermissionError:
        messagebox.showerror("Error", f"Permission denied: Could not write to '{output_path}'. Ensure the file is not open elsewhere and you have write permissions.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def main():
    window = tk.Tk()
    window.title("Excel Data Anonymizer")

    label = tk.Label(window, text="Excel Data Anonymizer", font=("Arial", 16))
    label.pack(pady=20)

    btn = tk.Button(window, text="Choose and Anonymize Excel File", command=choose_and_anonymize)
    btn.pack(pady=20)

    window.mainloop()

if __name__ == '__main__':
    main()
