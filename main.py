import tkinter as tk
from tkinter import filedialog, ttk, messagebox, Toplevel
import pandas as pd
import os
import platform
from Claas_CSV import CSVLoader
from DataFrameToExcel import DataFrameToExcel as DFEClass

global select_xlsx_path
select_xlsx_path = None


def position_top_window():
    # Position des Hauptfensters abrufen
    x = root.winfo_x()
    y = root.winfo_y()

    # Position des top-Fensters basierend auf der Position des Hauptfensters setzen
    #top.geometry(f"300x150+{x+20}+{y+40}")

def quit_program():
    root.destroy()

def select_excel_file(root):
    file_path = filedialog.askopenfilename(title="Wählen Sie eine Excel-Datei aus")

    # Überprüfen Sie, ob die ausgewählte Datei die richtige Erweiterung hat
    if not (file_path.endswith(".xlsx") or file_path.endswith(".xlsm")):
        print("Bitte wählen Sie eine gültige Excel-Datei (.xlsx oder .xlsm) aus!")
        return None

    return file_path

def run_sequence_excel():
    
    #global root
    from Excelprocessor import ExcelProcessor
    global select_xlsx_path
    global df_selected
    processor = ExcelProcessor(master=root)
    df_selected = processor.main()
    #processor.main()
    #select_xlsx_path = processor.get_save_path()
    processor.close()
    
    app = DFEClass(df_selected, master=root)
    print("DFEClass-Instanz erstellt")
    #select_xlsx_path = app.get_save_path()
    print(f"Datei gespeichert: {select_xlsx_path}")
    
def open_file():
    
    global select_xlsx_path
    # Hier verwenden Sie select_xlsx_path
    if select_xlsx_path and os.path.isfile(select_xlsx_path):
        if platform.system() == "Windows":
            os.startfile(select_xlsx_path)
        elif platform.system() == "Darwin":  # MacOS
            os.system("open " + select_xlsx_path)
        elif platform.system() == "Linux":
            os.system("xdg-open " + select_xlsx_path)
    else:
        print(f"Datei nicht gefunden: {select_xlsx_path}")

def open_csv_loader():
    loader = CSVLoader(root)
    root.wait_window(loader.top)
    df = loader.df
    print("DataFrame received in main program")
    
    from DataFrameToExcel import DataFrameToExcel as DFEClass
    df.reset_index(inplace=True)
    app = DFEClass(df, master=root)

def check_select_xlsx_path():
    
    global select_xlsx_path
    print("check")
    print(select_xlsx_path)
    try: 
        if select_xlsx_path:
            open_file_button.place(relx=0.9, rely=0.15, anchor="se")  # Machen Sie den Button sichtbar
    except AttributeError:
        # Hier könnten Sie den Button auch unsichtbar machen, wenn Sie möchten
        pass


# Hauptfenster erstellen

root = tk.Tk()
root.title("Main Program")
root.geometry("700x500+300+300")
root.configure(bg="#4F4F4F")

# Buttons "Funktion1" und "Funktion2" hinzufügen
function1_button = tk.Button(
    root,
    text="Read from Excel",
    command=run_sequence_excel,
    fg="Green",
    width=20,
    height=2,
)
function1_button.pack(anchor="nw", padx=10, pady=10)

function2_button = tk.Button(root, text="Read CSV", command=open_csv_loader, fg="Blue", width=20, height=2)
function2_button.pack(anchor="nw", padx=10, pady=10)

function3_button = tk.Button(root, text="Chack for files", command=check_select_xlsx_path, fg="Yellow", width=20, height=2)
function3_button.pack(anchor="nw", padx=10, pady=10)

# Erstellen Sie den "Load File"-Button, aber machen Sie ihn zuerst unsichtbar
open_file_button = tk.Button(
    root,
    text="Open File",
    command=open_file,
    fg="Red",
    width=20,
    height=2,
)
#open_file_button.place(relx=0.9, rely=0.15, anchor="se")

quit_button = tk.Button(
    root, text="quit program", command=quit_program, bg="Darkgrey", width=20, height=2
)
quit_button.place(relx=0.9, rely=0.9, anchor="se")

root.after(2000, check_select_xlsx_path)
root.mainloop()
