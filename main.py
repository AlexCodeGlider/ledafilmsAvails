import os
import sys
import tkinter
from tkinter import messagebox
import pandas as pd
from utils import get_app_dir
from dataProcess import process_data

def main():
    try:
        process_data()
    except Exception as e:
        # Display error message
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("Error", "An error occurred while processing the data. Please check the log file for more information.")
        # Write error message to log file
        app_dir = get_app_dir()
        log_file = os.path.join(app_dir, 'error.log')
        with open(log_file, 'w') as f:
            f.write(str(e))
        sys.exit(1)
    
    # Display completion message
    root = tkinter.Tk()
    root.withdraw()
    messagebox.showinfo("Task Completed", "Data processing complete! All tasks have been successfully executed.")

if __name__ == '__main__':
    main()
