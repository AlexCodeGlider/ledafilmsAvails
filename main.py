import os
import sys
import tkinter
from tkinter import messagebox
import pandas as pd
from utils import get_app_dir

def main():
    app_dir = get_app_dir()
    
    # Path to the external file
    data_file = os.path.join(app_dir, 'data', 'Contract Summary.xlsx')
    
    # Check if the file exists
    if not os.path.isfile(data_file):
        messagebox.showerror("Error", f"Required file not found: {data_file}")
        sys.exit(1)
    
    # Read and process the data file
    df = pd.read_excel(data_file)
    print(df.head())
    
    # Display completion message
    root = tkinter.Tk()
    root.withdraw()
    messagebox.showinfo("Task Completed", "Data processing complete! All tasks have been successfully executed.")

if __name__ == '__main__':
    main()
