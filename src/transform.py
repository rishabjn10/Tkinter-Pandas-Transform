import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd


def add_range_column(df):
    """
    Adds a new column 'Range' to the dataframe which is the difference between 'High' and 'Low' columns.

    Parameters:
    df (pd.DataFrame): The input dataframe containing 'High' and 'Low' columns.

    Returns:
    pd.DataFrame: The dataframe with the new 'Range' column.
    """
    df["Range"] = df["High"] - df["Low"]
    return df


def convert_date_column(df):
    """
    Converts the 'Date' column of the dataframe into datetime objects.

    Parameters:
    df (pd.DataFrame): The input dataframe containing 'Date' column.

    Returns:
    pd.DataFrame: The dataframe with 'Date' column converted to datetime objects.
    """
    df["Date"] = pd.to_datetime(df["Date"])
    return df


def add_next_month_column(df):
    """
    Adds a new column 'Next Month' to the dataframe which is the date one month after the 'Date' column.

    Parameters:
    df (pd.DataFrame): The input dataframe containing 'Date' column.

    Returns:
    pd.DataFrame: The dataframe with the new 'Next Month' column.
    """
    df["Next Month"] = df["Date"] + pd.DateOffset(months=1)
    return df


def transform_dataframe(df):
    """
    Applies all transformations to the dataframe: adds 'Range' column, converts 'Date' column,
    and adds 'Next Month' column.

    Parameters:
    df (pd.DataFrame): The input dataframe containing 'Date', 'High', and 'Low' columns.

    Returns:
    pd.DataFrame: The transformed dataframe with 'Range' and 'Next Month' columns added.
    """
    df = add_range_column(df)
    df = convert_date_column(df)
    df = add_next_month_column(df)
    return df


def load_data_from_excel(file_path):
    """
    Loads data from an Excel file into a DataFrame.

    Parameters:
    file_path (str): The path to the Excel file.

    Returns:
    pd.DataFrame: The loaded dataframe.
    """
    return pd.read_excel(file_path)


def save_data_to_excel(df, file_path):
    """
    Saves the DataFrame to an Excel file.

    Parameters:
    df (pd.DataFrame): The dataframe to be saved.
    file_path (str): The path to the Excel file.
    """
    df.to_excel(file_path, index=False)


class DataFrameTransformerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DataFrame Transformer")

        self.df = None

        self.import_button = tk.Button(root, text="Import", command=self.import_file)
        self.import_button.pack(pady=10)

        self.transform_button = tk.Button(
            root, text="Transform", command=self.transform_data
        )
        self.transform_button.pack(pady=10)

        self.export_button = tk.Button(root, text="Export", command=self.export_file)
        self.export_button.pack(pady=10)

    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.df = pd.read_excel(file_path)
            messagebox.showinfo("Info", "File imported successfully!")

    def transform_data(self):
        if self.df is not None:
            self.df = transform_dataframe(self.df)
            messagebox.showinfo("Info", "Transformation is complete!")
        else:
            messagebox.showwarning("Warning", "Please import a file first!")

    def export_file(self):
        if self.df is not None:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
            )
            if file_path:
                self.df.to_excel(file_path, index=False)
                messagebox.showinfo("Info", "File exported successfully!")
        else:
            messagebox.showwarning(
                "Warning", "Please import and transform a file first!"
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = DataFrameTransformerApp(root)
    root.mainloop()
