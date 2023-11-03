# In der Datei main.py
from excelselect import selector
from Class.DataFrameToExcel import DataFrameToExcel as DFTE
import pandas as pd


root = None
select_xlsx_path = selector.run_selector(root)
df = pd.read_excel(select_xlsx_path)
app = DFTE(df)
save_path = app.get_save_path()
print("Gespeicherter Pfad:", save_path)
# app = DFTE(df, show_gui=False)
#
# params = {
#    "sort_column": "project_number",
#    "title_bg_color": "000000",
#    "title_font_color": "FFFFFF",
#    "file_name": "output_alex__test_file",
#    "correct_date_format": "no",
#    "highlight_rows": "yes",
#    "highlight_column": "rdg_id",
#    "highlight_value": "663",
#    "highlight_color": "00FFFF",
#    "file_path": None,
#    "date_columns": ["plan_g2_date", "plan_g3_date", "plan_g4_date", "plan_g5_date"],
# }
#
# optional auto setting of file path
# directory_path = os.path.dirname(file_path)
# output_file_path = os.path.join(directory_path, params["file_name"] + ".xlsx")
# params["file_path"] = output_file_path

# app.save_to_excel(params=params)
# save_path = app.get_save_path()
# print("Der Aufruf war erfolgreich")
# print("Gespeicherter Pfad:", save_path)
