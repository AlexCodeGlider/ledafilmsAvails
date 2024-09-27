import os
import sys
import pandas as pd
import tkinter
from tkinter import messagebox
import json

def get_app_dir():
    """Get the directory where the executable or script is located."""
    if getattr(sys, 'frozen', False):
        # Running in a bundle (PyInstaller executable)
        application_path = os.path.dirname(sys.executable)
    else:
        # Running in a normal Python environment
        application_path = os.path.dirname(os.path.abspath(__file__))
    return application_path

def check_file(data_file):
    # Check if the files exists
    if not os.path.isfile(data_file):
        messagebox.showerror("Error", f"Required file not found: {data_file}")
        sys.exit(1)

def clean_date(date_string, start_date=True):
    try:
        formated_date = pd.to_datetime(date_string)
        if pd.isna(formated_date):
            return pd.Timestamp('2100-12-31 00:00:00')
        return formated_date
    except pd.errors.OutOfBoundsDatetime:
        if start_date:
            return pd.Timestamp('1974-11-01 00:00:00')
        return pd.Timestamp('2100-12-31 00:00:00')
    
def tidy_split(df, column, sep=',', keep=False):
    """
    Split the values of a column and expand so the new DataFrame has one split
    value per row. Filters rows where the column is missing.

    Params
    ------
    df : pandas.DataFrame
        dataframe with the column to split and expand
    column : str
        the column to split and expand
    sep : str
        the string used to split the column's values
    keep : bool
        whether to retain the presplit value as it's own row

    Returns
    -------
    pandas.DataFrame
        Returns a dataframe with the same columns as `df`.
    """
    indexes = list()
    new_values = list()
    df = df.dropna(subset=[column])
    for i, presplit in enumerate(df[column].astype(str)):
        values = presplit.split(sep)
        if keep and len(values) > 1:
            indexes.append(i)
            new_values.append(presplit)
        for value in values:
            indexes.append(i)
            new_values.append(value.strip())
    new_df = df.iloc[indexes, :].copy()
    new_df[column] = new_values
    return new_df

def getCategories(df, df_name, max_unique=12):
    category_cols = [
        col for col in df.columns if (len(df[col].unique()) <= max_unique) and (df[col].dtype not in [bool, int])
    ]
    categories = {
        col:list(df[col].unique()) for col in category_cols
    }
    with open(df_name+"_enums.json", "w") as outfile:
        json.dump(categories, outfile)