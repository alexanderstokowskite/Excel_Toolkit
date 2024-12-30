import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from sklearn.metrics import confusion_matrix
import numpy as np


def ask_file_path():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="Select a file",
        filetypes=(("Excel files", "*.xlsx"), ("CSV files", "*.csv")),
    )
    root.destroy()
    return file_path


def load_data(file_path):
    if file_path.endswith(".csv"):
        return pd.read_csv(file_path)
    elif file_path.endswith(".xlsx"):
        df = pd.ExcelFile(file_path)
        sheet_name = ask_sheet(df.sheet_names)
        return pd.read_excel(file_path, sheet_name=sheet_name)


def ask_sheet(sheets):
    root = tk.Tk()
    root.title("Sheet Selection")
    ttk.Label(root, text="Select a sheet:").pack(pady=20)
    var = tk.StringVar(root)
    var.set(sheets[0])
    popupMenu = ttk.OptionMenu(root, var, *sheets)
    popupMenu.pack()
    button = ttk.Button(root, text="OK", command=root.destroy)
    button.pack(pady=20)
    root.mainloop()
    return var.get()


def select_columns(df):
    root = tk.Tk()
    root.title("Column Selection")
    ttk.Label(root, text="Select the Prediction column:").pack(pady=10)
    pred_var = tk.StringVar(root)
    pred_var.set(df.columns[0])
    pred_menu = ttk.OptionMenu(root, pred_var, *df.columns)
    pred_menu.pack()

    ttk.Label(root, text="Select the Actual column:").pack(pady=10)
    real_var = tk.StringVar(root)
    real_var.set(df.columns[0])
    real_menu = ttk.OptionMenu(root, real_var, *df.columns)
    real_menu.pack()

    button = ttk.Button(root, text="Continue", command=root.destroy)
    button.pack(pady=20)
    root.mainloop()
    return pred_var.get(), real_var.get()


def ask_use_matches_only():
    root = tk.Tk()
    root.title("Match Entries Only")
    ttk.Label(root, text="Use matching entries only?").pack(pady=20)
    var = tk.StringVar(root)
    var.set("No")
    yes_no_menu = ttk.OptionMenu(root, var, "Yes", "No")
    yes_no_menu.pack()
    button = ttk.Button(root, text="OK", command=root.destroy)
    button.pack(pady=20)
    root.mainloop()
    return var.get() == "Yes"


def plot_confusion_matrix(df, pred_col, real_col, use_matches_only):
    if use_matches_only:
        df = df[df[pred_col].isin(df[real_col]) & df[real_col].isin(df[pred_col])]
    labels = sorted(pd.unique(df[[pred_col, real_col]].values.ravel()))
    conf_mat = confusion_matrix(df[real_col], df[pred_col], labels=labels)
    conf_mat_normalized = conf_mat.astype("float") / conf_mat.sum(axis=1)[:, np.newaxis]

    fig, ax = plt.subplots(nrows=1, ncols=2, figsize=(18, 7))

    # Plotting the actual count confusion matrix
    sns.heatmap(
        conf_mat,
        annot=True,
        fmt="d",
        cmap="Blues",
        ax=ax[0],
        annot_kws={"size": 10},
        xticklabels=labels,
        yticklabels=labels,
    )
    ax[0].set_xlabel("Predicted")
    ax[0].set_ylabel("Actual")
    ax[0].set_title("Confusion Matrix (Counts)")

    # Plotting the accuracy percentage confusion matrix
    sns.heatmap(
        conf_mat_normalized,
        annot=True,
        fmt=".2%",
        cmap="Blues",
        ax=ax[1],
        annot_kws={"size": 10},
        xticklabels=labels,
        yticklabels=labels,
    )
    ax[1].set_xlabel("Predicted")
    ax[1].set_ylabel("Actual")
    ax[1].set_title("Confusion Matrix (Accuracy per Cell)")

    plt.show()


def main():
    file_path = ask_file_path()
    if file_path:
        df = load_data(file_path)
        pred_col, real_col = select_columns(df)
        use_matches_only = ask_use_matches_only()
        plot_confusion_matrix(df, pred_col, real_col, use_matches_only)


if __name__ == "__main__":
    main()
