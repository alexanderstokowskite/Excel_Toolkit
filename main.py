import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
import platform
from Class.Claas_CSV import CSVLoader
from excelselect import selector
from Class.DataFrameToExcel import DataFrameToExcel as DFTE
import subprocess

global file_path
file_path = None


def clear_window(window):
    for widget in window.winfo_children():
        widget.destroy()


def position_top_window():
    # Position des Hauptfensters abrufen
    x = root.winfo_x()
    y = root.winfo_y()

    # Position des top-Fensters basierend auf der Position des Hauptfensters setzen
    # top.geometry(f"300x150+{x+20}+{y+40}")


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
    root.withdraw()
    select_xlsx_path = selector.run_selector(root)
    df_selected = pd.read_excel(select_xlsx_path)

    # Abfrage, ob der Excel Formatter gestartet werden soll
    def on_yes():
        global file_path
        top.destroy()
        main_window()
        app = DFTE(df_selected, master=root)
        print("DFTEClass-Instanz beendet")

        # file_path = app.get_save_path()

    def on_no():
        global file_path
        top.destroy()
        select_xlsx_path = filedialog.asksaveasfilename(
            initialfile="select.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if select_xlsx_path:
            df_selected.to_excel(select_xlsx_path, sheet_name="Sheet", index=False)
            tk.messagebox.showinfo(
                "Information", "Datei wurde gespeichert unter: " + select_xlsx_path
            )
            print(select_xlsx_path)
            file_path = select_xlsx_path
            main_window()

    top = tk.Toplevel(root)
    top.title("Excel Formatter")

    label = tk.Label(top, text="Do you want to launch the Excel formatter?")
    label.pack(pady=10)

    yes_button = tk.Button(top, text="Yes", command=on_yes)
    no_button = tk.Button(top, text="No", command=on_no)
    yes_button.pack(side="left", expand=True, fill="both")
    no_button.pack(side="right", expand=True, fill="both")


def open_file(path):
    # Hier verwenden Sie path
    if path and os.path.isfile(path):
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":  # MacOS
            os.system("open " + path)
        elif platform.system() == "Linux":
            os.system("xdg-open " + path)
    else:
        print(f"Datei nicht gefunden: {path}")


def open_csv_loader():
    loader = CSVLoader(root)
    root.wait_window(loader.top)
    df = loader.df
    print("DataFrame received in main program")

    df.reset_index(inplace=True)
    app = DFTE(df, master=root)


def open_file_with_default_application():
    file_path = filedialog.askopenfilename()

    if file_path:
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":
                subprocess.call(["open", file_path])
            elif platform.system() == "Linux":
                subprocess.call(["xdg-open", file_path])
            else:
                raise EnvironmentError("Unsupported operating system")

        except EnvironmentError as e:
            print(e)
    else:
        print("Keine Datei ausgewählt")


def compare_excel_with_ui():
    import Comp_unique_Func as cu

    cu.main()


def compare_excel_line_by_line():
    import Comp_LbL_Func as clbl

    clbl.main()


def main_window():
    global file_path
    clear_window(root)
    root.deiconify()
    root.title("Function Selector")
    root.geometry("700x500+300+300")
    root.configure(bg="#4F4F4F")
    global image
    image = tk.PhotoImage(
        file="image/SWAT.png"
    )  # Bitte ersetzen Sie "image.png" durch den tatsächlichen Pfad zu Ihrem Bild
    # OSX pfad / - Windows Pfad mit \
    image_label = tk.Label(root, image=image)
    image_label.pack(side="right", fill="both", expand="yes")

    # Buttons "Funktion1" und "Funktion2" hinzufügen
    function1_button = tk.Button(
        root,
        text="Read from Excel",
        command=run_sequence_excel,
        fg="Green",
        width=22,
        height=2,
    )
    function1_button.pack(anchor="nw", padx=10, pady=10)

    function2_button = tk.Button(
        root, text="Read CSV", command=open_csv_loader, fg="Blue", width=22, height=2
    )
    function2_button.pack(anchor="nw", padx=10, pady=10)

    function3_button = tk.Button(
        root,
        text="Compare Excel with UI",
        command=compare_excel_with_ui,
        fg="orange",
        width=22,
        height=2,
    )
    function3_button.pack(anchor="nw", padx=10, pady=10)

    function4_button = tk.Button(
        root,
        text="Compare Excel Line by Line",
        command=compare_excel_line_by_line,
        fg="Red",
        width=22,
        height=2,
    )
    function4_button.pack(anchor="nw", padx=10, pady=10)

    function5_button = tk.Button(
        root,
        text="Check for files",
        command=open_file_with_default_application,
        fg="Darkgrey",
        width=22,
        height=2,
    )
    function5_button.pack(anchor="nw", padx=10, pady=10)

    # Erstellen Sie den "Load File"-Button, aber machen Sie ihn zuerst unsichtbar
    open_file_button = tk.Button(
        root,
        text="Open File",
        command=lambda: open_file(file_path),
        fg="Red",
        width=20,
        height=2,
    )

    try:
        path = file_path
        if path:
            open_file_button.place(
                relx=0.9, rely=0.15, anchor="se"
            )  # Machen Sie den Button sichtbar
    except AttributeError:
        # Hier könnten Sie den Button auch unsichtbar machen, wenn Sie möchten
        pass

    # open_file_button.place(relx=0.9, rely=0.15, anchor="se")

    quit_button = tk.Button(
        root,
        text="quit program",
        command=quit_program,
        bg="Darkgrey",
        width=20,
        height=2,
    )
    quit_button.place(relx=0.9, rely=0.9, anchor="se")


root = tk.Tk()
main_window()
root.mainloop()
