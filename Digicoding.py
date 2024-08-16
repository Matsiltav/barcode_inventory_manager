import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import re

# Function to sanitize strings by removing or replacing illegal characters
def sanitize_string(value):
    if isinstance(value, str):
        # Remove non-ASCII characters and specific problematic characters
        value = re.sub(r'[^\x20-\x7E]', '', value)
        value = value.replace("↔", "_")
    return value

# Function to remove trailing zeros
def remove_trailing_zeros(value):
    if isinstance(value, str):
        value = value.rstrip('0')
    return value

# Parsing the matrix code and extracting relevant data
def parse_matrix_code(matrix_code):
    try:
        # Assuming the format based on observed patterns
        segments = matrix_code.strip('|').split('P')
        part_number = segments[1].split('K')[0].strip()
        mfr_part_number = segments[2].split('K')[0].strip() if len(segments) > 2 else ""
        lot_code = segments[-1].rstrip('Z').strip()  # Stripping 'Z' and trimming whitespace
        description = "Parsed Description"  # Placeholder description

        # Remove trailing zeros from relevant fields
        part_number = remove_trailing_zeros(part_number)
        mfr_part_number = remove_trailing_zeros(mfr_part_number)
        lot_code = remove_trailing_zeros(lot_code)
        
        # Further clean up lot code by removing any trailing characters
        lot_code = re.sub(r'[^\w\s]', '', lot_code)
        
        return part_number, mfr_part_number, description, lot_code
    except Exception as e:
        messagebox.showerror("Error", f"Error parsing matrix code: {e}")
        return None, None, None, None

# GUI for scanning and saving to Excel
class BarcodeScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Barcode Scanner")
        self.barcodes = []

        # UI Elements
        self.barcode_entry = tk.Entry(root, width=50)
        self.barcode_entry.grid(row=0, column=0, padx=10, pady=10)
        self.barcode_entry.bind('<Return>', self.add_barcode)

        self.add_button = tk.Button(root, text="Add Barcode", command=self.add_barcode)
        self.add_button.grid(row=0, column=1, padx=10, pady=10)

        self.undo_button = tk.Button(root, text="Undo Last Scan", command=self.undo_last_scan)
        self.undo_button.grid(row=1, column=0, padx=10, pady=10)

        self.delete_button = tk.Button(root, text="Delete Selected", command=self.delete_selected)
        self.delete_button.grid(row=1, column=1, padx=10, pady=10)

        self.export_button = tk.Button(root, text="Export to Excel", command=self.export_to_excel)
        self.export_button.grid(row=2, column=0, padx=10, pady=10)

        self.export_as_button = tk.Button(root, text="Export As...", command=self.export_as)
        self.export_as_button.grid(row=2, column=1, padx=10, pady=10)

        self.barcode_listbox = tk.Listbox(root, height=20, width=90)
        self.barcode_listbox.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

    def add_barcode(self, event=None):
        barcode_data = self.barcode_entry.get().strip()
        if barcode_data:
            part_number, mfr_part_number, description, lot_code = parse_matrix_code(barcode_data)
            if part_number and mfr_part_number and description and lot_code:
                sanitized_part_number = sanitize_string(part_number)
                sanitized_mfr_part_number = sanitize_string(mfr_part_number)
                sanitized_description = sanitize_string(description)
                sanitized_lot_code = sanitize_string(lot_code)

                self.barcodes.append({
                    "Part Number": sanitized_part_number,
                    "MFR Part Number": sanitized_mfr_part_number,
                    "Description": sanitized_description,
                    "Lot Code": sanitized_lot_code
                })
                self.barcode_listbox.insert(tk.END, f"{sanitized_part_number} - {sanitized_mfr_part_number} - {sanitized_description} - {sanitized_lot_code}")
                self.barcode_entry.delete(0, tk.END)
            else:
                messagebox.showerror("Error", "Failed to parse matrix code.")

    def undo_last_scan(self):
        if self.barcodes:
            self.barcodes.pop()
            self.barcode_listbox.delete(tk.END)

    def delete_selected(self):
        selected_indices = self.barcode_listbox.curselection()
        for index in reversed(selected_indices):
            self.barcodes.pop(index)
            self.barcode_listbox.delete(index)

    def export_to_excel(self):
        df = pd.DataFrame(self.barcodes)
        today_date = datetime.now().strftime("%Y-%m-%d")
        file_path = filedialog.asksaveasfilename(
            initialfile=f"Inventory_{today_date}.xlsx", 
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.save_to_excel_with_formatting(df, file_path)

    def export_as(self):
        df = pd.DataFrame(self.barcodes)
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.save_to_excel_with_formatting(df, file_path)

    def save_to_excel_with_formatting(self, df, file_path):
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Inventory Data')

                # Access the openpyxl workbook and sheet
                workbook = writer.book
                sheet = workbook.active

                # Set column widths and bold headers
                for col_num, column_title in enumerate(df.columns, 1):
                    col_letter = get_column_letter(col_num)
                    sheet.column_dimensions[col_letter].width = max(len(column_title) + 2, 15)
                    sheet[col_letter + '1'].font = Font(bold=True)

        except PermissionError as e:
            messagebox.showerror("Permission Error", f"Could not save the file: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeScannerApp(root)
    app.run()
