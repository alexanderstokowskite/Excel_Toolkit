import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Listbox, MULTIPLE
import os


class ExcelToCSVConverter:
    def __init__(self, root):
        self.root = root
        self.selected_sheets = []

    def open_excel_file(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xls *.xlsx *.xlsm")]
        )
        if not filepath:
            return None, None
        try:
            excel_file = pd.ExcelFile(filepath)
            sheet_names = excel_file.sheet_names
            return filepath, sheet_names
        except Exception as e:
            messagebox.showerror("Error", f"Could not read the Excel file: {str(e)}")
            return None, None

    def select_sheets(self, sheet_names):
        self.selected_sheets = []

        def on_ok():
            selected_indices = listbox.curselection()
            self.selected_sheets = [sheet_names[i] for i in selected_indices]
            top.destroy()

        top = Toplevel(self.root)
        top.title("Select Sheets")
        top.geometry("400x350+1050+300")
        top.configure(bg="lightgrey")

        listbox = Listbox(
            top,
            selectmode=MULTIPLE,
            bg="white",
            fg="black",
            font=("Arial", 10),
            bd=0,
            highlightthickness=0,
        )
        for name in sheet_names:
            listbox.insert(tk.END, " " * 5 + name)  # Einrückung durch Leerzeichen
        listbox.pack(
            fill=tk.BOTH, expand=True, padx=20, pady=20
        )  # Einrückung durch Padding

        ok_button = tk.Button(top, text="OK", command=on_ok, bg="lightgrey")
        ok_button.pack(pady=10)

        top.transient(self.root)
        try:
            top.grab_set()
            self.root.wait_window(top)
        except tk.TclError as e:
            messagebox.showerror("Error", f"Could not grab the window: {str(e)}")

    def save_sheets_as_csv(self, filepath, selected_sheets):
        try:
            excel_file = pd.ExcelFile(filepath)
            directory = os.path.dirname(filepath)
            for sheet in selected_sheets:
                df = pd.read_excel(excel_file, sheet_name=sheet)
                csv_filename = os.path.join(directory, f"{sheet}.csv")
                df.to_csv(csv_filename, index=False)
            messagebox.showinfo(
                "Success", "Selected sheets have been saved as CSV files."
            )
        except Exception as e:
            messagebox.showerror("Error", f"Could not save sheets as CSV: {str(e)}")

    def run(self):
        filepath, sheet_names = self.open_excel_file()
        if filepath and sheet_names:
            self.select_sheets(sheet_names)
            if self.selected_sheets:
                self.save_sheets_as_csv(filepath, self.selected_sheets)
