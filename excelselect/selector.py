import tkinter as tk
from tkinter import filedialog, ttk, messagebox, Toplevel
import pandas as pd
import os

# This is the module of the excle selection tool


def select_excel_file(root):
    file_path = filedialog.askopenfilename(title="Wählen Sie eine Excel-Datei aus")

    # Überprüfen Sie, ob die ausgewählte Datei die richtige Erweiterung hat
    if not (file_path.endswith(".xlsx") or file_path.endswith(".xlsm")):
        print("Bitte wählen Sie eine gültige Excel-Datei (.xlsx oder .xlsm) aus!")
        return None

    return file_path


def list_and_select_sheet(root, file_path):
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names

    label = tk.Label(root, text="Bitte wählen Sie ein Arbeitsblatt aus:")
    label.pack(pady=20)

    combo = ttk.Combobox(root, values=sheet_names)
    combo.pack(pady=20)
    combo.set(sheet_names[0])

    selected_sheet_name = None

    def on_ok():
        nonlocal selected_sheet_name
        selected_sheet_name = combo.get()
        root.quit()

    ok_btn = tk.Button(root, text="OK", command=on_ok)
    ok_btn.pack(pady=20)

    root.mainloop()
    return selected_sheet_name


def write_columns_to_txt(root, file_path, selected_sheet_name):
    df = pd.read_excel(file_path, sheet_name=selected_sheet_name, nrows=0)
    columns = df.columns

    dir_path = os.path.dirname(file_path)
    txt_path = os.path.join(dir_path, "content.txt")

    with open(txt_path, "w") as txt_file:
        for col in columns:
            txt_file.write(f"{col}\n")

    # Zeige eine Bestätigungsmeldung mit Tkinter
    messagebox.showinfo(
        "Erfolgreich", f"Spaltennamen wurden in {txt_path} geschrieben.", parent=root
    )


def apply_filters_from_preset(df, root):
    # Lasse den Benutzer die Filter-Preset txt-Datei auswählen
    filter_txt_path = filedialog.askopenfilename(
        parent=root,
        title="Wählen Sie die Filter-Preset txt-Datei aus",
        filetypes=[("Text files", "*.txt")],
        defaultextension=".txt",
    )

    # Wenn der Benutzer den Dialog abbricht, ohne eine Datei auszuwählen, kehre sofort zurück
    if not filter_txt_path:
        return df

    # Ansonsten verarbeite die Filter
    with open(filter_txt_path, "r") as txt_file:
        lines = txt_file.readlines()

    for line in lines:
        line = line.strip()
        if not line or line.startswith("#"):
            continue

        column_name, criteria = line.split(" ", 1)
        operator, values = criteria.split(" ", 1)

        # Konvertiere alle Werte in der Spalte in Strings
        df[column_name] = df[column_name].astype(str)

        values_list = [str(value).strip() for value in values.split()]

        if operator == "==":
            if len(values_list) == 1:
                df = df[df[column_name] == values_list[0]]
            else:
                df = df[df[column_name].isin(values_list)]
        elif operator == "!=":
            if len(values_list) == 1:
                df = df[df[column_name] != values_list[0]]
            else:
                df = df[~df[column_name].isin(values_list)]
        elif operator == "=":
            # Kombiniere die Kriterien mit einem | für die ODER-Verknüpfung
            or_combined_criteria = "|".join(values_list)
            df = df[
                df[column_name].str.contains(or_combined_criteria, case=False, na=False)
            ]

    return df


def select_columns(df, root):
    # Lasse den Benutzer die select.txt-Datei auswählen
    select_txt_path = filedialog.askopenfilename(
        parent=root,
        title="Wählen Sie die select.txt-Datei aus",
        filetypes=[("Text files", "*.txt")],
    )

    # Wenn der Benutzer den Dialog abbricht, ohne eine Datei auszuwählen
    if not select_txt_path:
        return df

    # Lese die Spaltennamen aus der select.txt-Datei
    with open(select_txt_path, "r") as txt_file:
        selected_columns = [
            line.strip() for line in txt_file.readlines() if not line.startswith("#")
        ]

    # Überprüfe, ob alle Spaltennamen im DataFrame vorhanden sind
    missing_columns = [col for col in selected_columns if col not in df.columns]
    if missing_columns:
        messagebox.showerror(
            "Error",
            f"Folgende Spalten wurden im DataFrame nicht gefunden: {', '.join(missing_columns)}",
            parent=root,
        )
        return df

    # Wähle die angegebenen Spalten
    df_selected = df[selected_columns]
    return df_selected


def run_selector(root):
    if "root" not in globals() or globals()["root"] is None:
        root = tk.Tk()
        root.title("Choose")
        root.withdraw()
        print("Die Variable 'root' wurde definiert.")
    else:
        print("Die Variable 'root' war bereits definiert.")
    selected_file = select_excel_file(root)
    if selected_file:
        root.deiconify()  # Show the main window
        selected_sheet_name = list_and_select_sheet(root, selected_file)
        if selected_sheet_name:
            write_columns_to_txt(root, selected_file, selected_sheet_name)

            # Lese den gesamten DataFrame
            df = pd.read_excel(selected_file, sheet_name=selected_sheet_name)
            # Nachdem Sie den DataFrame eingelesen haben
            if "rdg_id" in df.columns:
                df["rdg_id"] = (
                    df["rdg_id"].fillna(-1).astype(int).astype(str).replace("-1", "")
                )

            # Wende die Filter-Presets an, falls bereitgestellt
            df_filtered = apply_filters_from_preset(df, root)

            # Wende die Spaltenselektion an
            df_selected = select_columns(df_filtered, root)

            # Speichern Sie den gefilterten und ausgewählten DataFrame in select.xlsx
            dir_path = os.path.dirname(selected_file)
            select_xlsx_path = os.path.join(dir_path, "select.xlsx")
            df_selected.to_excel(select_xlsx_path, index=False)

            messagebox.showinfo(
                "Erfolgreich",
                f"Die gefilterten und ausgewählten Daten wurden in {select_xlsx_path} gespeichert.",
                parent=root,
            )

            root.destroy()
        return select_xlsx_path


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    run_selector(root)
