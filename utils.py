import os
import sys
import pandas as pd
import tkinter
from tkinter import messagebox
import json
import csv

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

def avails_region(cell):
    latam_countries = [
        'Costa Rica',
        'El Salvador',
        'Guatemala',
        'Honduras',
        'Nicaragua',
        'Panama',
        'Mexico',
        'Argentina',
        'Bolivia',
        'Brazil',
        'Chile',
        'Colombia',
        'Ecuador',
        'Paraguay',
        'Peru',
        'Uruguay',
        'Venezuela'
    ]

    latam_ex_mex = [
        'Costa Rica',
        'El Salvador',
        'Guatemala',
        'Honduras',
        'Nicaragua',
        'Panama',
        'Argentina',
        'Bolivia',
        'Brazil',
        'Chile',
        'Colombia',
        'Ecuador',
        'Paraguay',
        'Peru',
        'Uruguay',
        'Venezuela'
    ]

    latam_ex_brz = [
        'Costa Rica',
        'El Salvador',
        'Guatemala',
        'Honduras',
        'Nicaragua',
        'Panama',
        'Mexico',
        'Argentina',
        'Bolivia',
        'Chile',
        'Colombia',
        'Ecuador',
        'Paraguay',
        'Peru',
        'Uruguay',
        'Venezuela'
    ]

    latam_ex_mex_brz = [
        'Costa Rica',
        'El Salvador',
        'Guatemala',
        'Honduras',
        'Nicaragua',
        'Panama',
        'Argentina',
        'Bolivia',
        'Chile',
        'Colombia',
        'Ecuador',
        'Paraguay',
        'Peru',
        'Uruguay',
        'Venezuela'
    ]

    caribbean = [
        'Anguilla',
        'Antigua and Barbuda',
        'Aruba',
        'Bahamas',
        'Barbados',
        'Bermuda',
        'Bonaire',
        'British Virgin Islands',
        'Cayman Islands',
        'Cuba',
        'Curaçao',
        'Dominica',
        'Dominican Republic',
        'Grenada',
        'Guadeloupe',
        'Haiti',
        'Jamaica',
        'Martinique',
        'Montserrat',
        'Puerto Rico',
        'Saint Barthélemy',
        'Saint Kitts and Nevis',
        'Saint Lucia',
        'Saint Vincent and the Grenadines',
        'Sint Eustatius',
        'Sint Maarten',
        'Trinidad and Tobago',
        'Turks and Caicos Islands',
        'United States Virgin Islands',
    ]

    dependencies = ['Guyana', 'French Guiana', 'Suriname', 'Belize']

    world = set(['Moldova', 'Malaysia', 'Qatar', 'Luxembourg', 'Portugal', 'Kenya', 'United Kingdom', 'Ghana', 'Andorra', 'South Korea', 'Latvia', 'Mayotte', 'Comoros', 'Turkmenistan', 'South Africa', 'Ukraine', 'Singapore', 'Kazakhstan', 'India', 'Slovenia', 'Bahrain', 'Indonesia', 'Lebanon', 'Tajikistan', 'Cambodia', 'Syria', 'Cameroon', 'Burundi', 'Tonga', 'Lithuania', 'Gabon', 'Bosnia and Herzegovina', 'Romania', 'Finland', 'Spain', 'Mauritania', 'Croatia', 'Djibouti', 'Bulgaria', 'Tanzania', 'South Sudan', 'Greece', 'Sudan', 'Brunei', 'Vietnam', 'North Korea', 'Philippines', 'Pakistan', 'Papua New Guinea', 'Uzbekistan', 'Sweden', 'Azerbaijan', 'New Zealand', 'Liberia', 'Vanuatu', 'Yemen', 'Russia', 'Czech Republic', 'Taiwan', 'Mongolia', 'Senegal', 'Botswana', 'Georgia', 'Serbia', 'United Arab Emirates', 'Iran', 'France', 'Saudi Arabia', 'Liechtenstein', 'Uganda', 'Zambia', 'Montenegro', 'Cyprus', 'Bermuda', 'Puerto Rico', 'Australia', 'Central African Republic', 'Iceland', 'Burkina Faso', 'Germany', 'Algeria', 'Denmark', 'Malta', 'Benin', 'Hungary', 'Solomon Islands', 'Bangladesh', 'Mauritius', 'Nepal', 'Norway', 'Lesotho', 'Belgium', 'Kyrgyzstan', 'New Caledonia', 'Fiji', 'Italy', 'Malawi', 'Bahamas', 'Seychelles', 'Madagascar', 'Sri Lanka', 'Kosovo', 'Israel', 'Laos', 'Togo', 'Canada', 'Guinea', 'Zimbabwe', 'French Polynesia', 'Albania', 'China', 'Mali', 'Ethiopia', 'Morocco', 'Namibia', 'Egypt', 'Japan', 'Bhutan', 'Belarus', 'Sierra Leone', 'Equatorial Guinea', 'Jordan', 'Estonia', 'Armenia', 'Turkey', 'Chad', 'Rwanda', 'Guinea-Bissau', 'Hong Kong', 'Switzerland', 'Nigeria', 'Kuwait', 'Monaco', 'Poland', 'Eritrea', 'Afghanistan', 'Iraq', 'Tuvalu', 'Mozambique', 'Ireland', 'Kiribati', 'Niger', 'Angola', 'Tunisia', 'Slovakia', 'Somalia', 'Libya', 'Thailand', 'Austria', 'Oman'])

    cell = cell - set(caribbean) - set(dependencies)
    if (
        set(latam_countries).issubset(cell) and \
        len(set(world).intersection(cell)) <= 16
    ):
        return 'All Latam'
    elif (
        set(latam_ex_mex).issubset(cell) and \
        len(set(world).intersection(cell)) <= 16
    ):
        return 'Latam excluding Mexico'
    elif (
        set(latam_ex_brz).issubset(cell) and \
        len(set(world).intersection(cell)) <= 16
    ):
        return 'Latam excluding Brazil'
    elif (
        set(latam_ex_mex_brz).issubset(cell) and \
        len(set(world).intersection(cell)) <= 16
    ):
        return 'Latam excluding Mexico and Brazil'
    elif (
        len(set(world).intersection(cell)) >= 50 and \
        len(set(latam_countries).intersection(cell)) >= 16
    ):
        return 'Worldwide'
    elif (
        len(set(world).intersection(cell)) >= 50 and \
        len(set(latam_countries).intersection(cell)) < 16
    ):
        return 'Worldwide excluding Latam'
    else:
        return cell

def clean_str(cell):
    # Define the characters to remove
    chars_to_remove = '[]{}\"\''

    # Create a translation table that maps the unwanted characters to None
    translation_table = str.maketrans('', '', chars_to_remove)

    # Apply the translation table to the input string
    new_cell = str(cell).translate(translation_table)
    new_cell = new_cell.replace('None', '')
    new_cell = new_cell.replace('nan', '')
    new_cell = new_cell.replace('NaT', '')
    return new_cell

# read the first row of a csv file into a list
def read_csv_header(csv_file):
    with open(csv_file, 'r') as f:
        reader = csv.reader(f)
        return next(reader)
    
# Function to apply on each row
def max_date(row):
    today = pd.to_datetime('today')
    if today < row[('start_date', 'Sales', 'License')]:
        # Exclude 'sales_start_date' and 'sales_end_date', and return max date
        return max(row.drop([('start_date', 'Sales', 'License'), ('end_date', 'Sales', 'License')]))
    else:
        # Return max date of the entire row
        return max(row)
    
def non_exclusive_end_date(row):
    sales_dates = [row.iloc[2], 
                row.iloc[3]]
    valid_dates = [
        date for date in sales_dates if row.iloc[0] <= date <= row.iloc[1]
    ]
    
    if valid_dates:
        return min(valid_dates)
    else:
        return pd.NaT