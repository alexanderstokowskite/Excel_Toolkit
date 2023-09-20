import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, Toplevel
from DataFrameToExcel import DataFrameToExcel as DFEClass

class ExcelProcessor:
    
    def __init__(self, master=None):
        self.file_path = None
        if master is None:
            self.top = tk.Tk()
            print("I am Master")     
        else:
            self.top = tk.Toplevel(master)
            self.top.withdraw()
            print("I am not master")
        
    def close(self): 
        self.top.withdraw()    
        
    def select_excel_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Wählen Sie eine Excel-Datei aus"
        )
        if not (self.file_path.endswith(".xlsx") or self.file_path.endswith(".xlsm")):
            print("Bitte wählen Sie eine gültige Excel-Datei (.xlsx oder .xlsm) aus!")
            return None
        return self.file_path

    def list_and_select_sheet(self):
        self.top.deiconify()  # Zeigen Sie das Hauptfenster an
        xls = pd.ExcelFile(self.file_path)
        sheet_names = xls.sheet_names

        top_level = tk.Toplevel(self.top)
        top_level.title("Arbeitsblatt auswählen")
        self.top.withdraw

        label = tk.Label(top_level, text="Bitte wählen Sie ein Arbeitsblatt aus:")
        label.pack(pady=20)

        combo = ttk.Combobox(top_level, values=sheet_names)
        combo.pack(pady=20)
        combo.set(sheet_names[0])

        self.selected_sheet_name = None

        def on_ok():
            self.selected_sheet_name = combo.get()
            top_level.destroy()

        ok_btn = tk.Button(top_level, text="OK", command=on_ok)
        ok_btn.pack(pady=20)

        top_level.transient(self.top)  # Setzen des Hauptfensters als übergeordnetes Fenster
        top_level.grab_set()  # Modal machen
        self.top.wait_window(top_level)  # Warten auf das Schließen des Toplevel-Fensters
        self.top.withdraw()  # Verstecken Sie das Hauptfenster erneut
        return self.selected_sheet_name

    def write_columns_to_txt(self):
        df = pd.read_excel(self.file_path, sheet_name=self.selected_sheet_name, nrows=0)
        columns = df.columns
        dir_path = os.path.dirname(self.file_path)
        txt_path = os.path.join(dir_path, "content.txt")
        with open(txt_path, "w") as txt_file:
            for col in columns:
                txt_file.write(f"{col}\n")
        messagebox.showinfo(
            "Erfolgreich",
            f"Spaltennamen wurden in {txt_path} geschrieben.",
            parent=self.top,
        )

    def apply_filters_from_preset(self, df):
        filter_txt_path = filedialog.askopenfilename(
            parent=self.top,
            title="Wählen Sie die Filter-Preset txt-Datei aus",
            filetypes=[("Text files", "*.txt")],
            defaultextension=".txt",
        )
        if not filter_txt_path:
            return df
        with open(filter_txt_path, "r") as txt_file:
            lines = txt_file.readlines()
        for line in lines:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            column_name, criteria = line.split(" ", 1)
            operator, values = criteria.split(" ", 1)
            df[column_name] = df[column_name].astype(str)
            values_list = [str(value).strip() for value in values.split()]
            if operator == "==":
                df = df[df[column_name].isin(values_list)]
            elif operator == "!=":
                df = df[~df[column_name].isin(values_list)]
            elif operator == "=":
                or_combined_criteria = "|".join(values_list)
                df = df[
                    df[column_name].str.contains(
                        or_combined_criteria, case=False, na=False
                    )
                ]
        return df

    def select_columns(self, df):
        select_txt_path = filedialog.askopenfilename(
            parent=self.top,
            title="Wählen Sie die select.txt-Datei aus",
            filetypes=[("Text files", "*.txt")],
        )
        if not select_txt_path:
            return df
        with open(select_txt_path, "r") as txt_file:
            selected_columns = [
                line.strip()
                for line in txt_file.readlines()
                if not line.startswith("#")
            ]
        missing_columns = [col for col in selected_columns if col not in df.columns]
        if missing_columns:
            messagebox.showerror(
                "Error",
                f"Folgende Spalten wurden im DataFrame nicht gefunden: {', '.join(missing_columns)}",
                parent=self.top,
            )
            return df
        df_selected = df[selected_columns]
        return df_selected
    
    def get_save_path(self):
        """Rückgabe des Pfads, in dem die Datei gespeichert wurde."""
        try:
            return self.select_xlsx_path
        except AttributeError:
            return None
    
    def main(self):
        
        #self.top.deiconify()
        selected_file = self.select_excel_file()
        if selected_file:
            #self.top.deiconify()
            self.selected_sheet_name = self.list_and_select_sheet()
            if self.selected_sheet_name:
                self.write_columns_to_txt()
                df = pd.read_excel(self.file_path, sheet_name=self.selected_sheet_name)
                if "rdg_id" in df.columns:
                    df["rdg_id"] = (
                        df["rdg_id"]
                        .fillna(-1)
                        .astype(int)
                        .astype(str)
                        .replace("-1", "")
                    )
                df_filtered = self.apply_filters_from_preset(df)
                df_selected = self.select_columns(df_filtered)
                dir_path = os.path.dirname(self.file_path)
                self.select_xlsx_path = os.path.join(dir_path, "select.xlsx")
              
        return df_selected
        

if __name__ == "__main__":
    processor = ExcelProcessor()
    processor.main()
