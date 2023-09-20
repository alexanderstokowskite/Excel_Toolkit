from tkinter import filedialog, Tk, Label, Entry, Button, StringVar, Toplevel, Checkbutton, BooleanVar
import pandas as pd

from tkinter import Checkbutton, IntVar

class CSVLoader:
    def __init__(self, master):
        self.top = Toplevel(master)
        self.file_path = filedialog.askopenfilename(
            title="Wählen Sie eine CSV-Datei aus"
        )
        self.df = None

        if self.file_path:
            self.split_char = StringVar(value=",")
            self.strip_char = StringVar(value='"')
            self.has_header = IntVar(value=1)

            Label(self.top, text="Split CSV at:").grid(row=0, column=0, padx=20, pady=5)
            Entry(self.top, textvariable=self.split_char).grid(
                row=0, column=1, padx=20, pady=5
            )

            Label(self.top, text="Strip additionally:").grid(
                row=1, column=0, padx=20, pady=5
            )
            Entry(self.top, textvariable=self.strip_char).grid(
                row=1, column=1, padx=20, pady=5
            )
            
            Checkbutton(self.top, text="Has Header", variable=self.has_header).grid(
                row=2, column=0, padx=20, pady=5
            )

            Button(self.top, text="Load Data", command=self.load_data).grid(
                row=3, column=0, columnspan=2, pady=20
            )

    def load_data(self):
        split_char = self.split_char.get()
        strip_char = self.strip_char.get()

        print(f"Separator: {repr(split_char)}")
        print(f"Quotechar: {repr(strip_char)}")

        self.df = pd.read_csv(self.file_path, sep=split_char, header=None)
        self.df = self.df[0].str.split(split_char, expand=True)

        if not self.has_header.get():
            # Wenn kein Header vorhanden ist, fügen Sie standardmäßige Spaltennamen hinzu
            self.df.columns = [f"Column_{i}" for i in range(self.df.shape[1])]
        else:
            # Entfernen der Anführungszeichen aus den Spaltennamen
            self.df.iloc[0] = self.df.iloc[0].str.strip(strip_char)
            self.df.columns = self.df.iloc[0]
            self.df = self.df[1:]

        self.df = self.df.apply(lambda x: x.str.strip(strip_char))
        self.df.set_index(self.df.columns[0], inplace=True)
        self.df.index.name = None
        self.df.columns.name = None

        self.df.dropna(how="all", inplace=True)
        print(self.df.head())
        self.top.destroy()




# Hauptprogramm
if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    def open_csv_loader():
        loader = CSVLoader(root)
        root.wait_window(loader.top)
        df = loader.df
        print("DataFrame received in main program")
        
        # Dateispeicherdialog anzeigen und Dateipfad zum Speichern erhalten
        #excel_path = filedialog.asksaveasfilename(title="Speichern als", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        # Exportieren der Daten in eine Excel-Datei
        #df.to_excel(excel_path)
        
        from DataFrameToExcel import DataFrameToExcel as DFEClass
        df.reset_index(inplace=True)
        app = DFEClass(df, master=root)

    Button(root, text="Open CSV Loader", command=open_csv_loader).pack(pady=20)
    root.deiconify()
    root.mainloop()