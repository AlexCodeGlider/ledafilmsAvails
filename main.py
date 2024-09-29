import os
import sys
import tkinter
from tkinter import messagebox
import pandas as pd
from utils import get_app_dir
from data_process import process_data
from avails import avails_process

def main():
    # Display a message box to state that the program is running
    root = tkinter.Tk()
    root.withdraw()
    messagebox.showinfo("Data Processing", "Data processing started. Click OK to continue.")

    # Process the data
    try:
        process_data()
        # Display completion message
        print("Data processing complete! Please wait while Avails are being generated.")
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

    # Process Avails
    try:
        avails_process()
    except Exception as e:
        # Display error message
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("Error", "An error occurred while processing the Avails. Please check the log file for more information.")
        # Write error message to log file
        app_dir = get_app_dir()
        log_file = os.path.join(app_dir, 'error.log')
        with open(log_file, 'w') as f:
            f.write(str(e))
    
    # Display completion message
    root = tkinter.Tk()
    root.withdraw()
    messagebox.showinfo("Task Completed", "Avails processing complete! Check the avails folder for the output files. Click OK to exit.")

if __name__ == '__main__':
    main()
