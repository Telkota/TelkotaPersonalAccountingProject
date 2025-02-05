import tkinter as tk
from tkinter import filedialog, messagebox
from logic import *

def open_file_dialog(filetype):
    file_path = filedialog.askopenfilename(filetypes=[filetype])
    return file_path

def save_file_dialog(default_extension, filetype):
    file_path = filedialog.asksaveasfilename(defaultextension=default_extension, filetypes=[filetype])
    return file_path

def create_initial_popup():
    def submit_files():
        global xlsx_path, csv_path, xlsx_path
        xlsx_path = xlsx_entry.get() if xlsx_entry.get() else None
        csv_path = csv_entry.get() if csv_entry.get() else None

        print(xlsx_path, csv_path, new_xlsx_path)

        if xlsx_path and csv_path:
            print("xlsx and csv paths found - continuing")
            initial_popup.destroy()
            create_main_gui()
        elif new_xlsx_path and csv_path:
            print(new_xlsx_path)
            print("New xlsx detected - continuing")
            create_new_doc(new_xlsx_path)
            xlsx_path = new_xlsx_path
            initial_popup.destroy()
            create_main_gui()
        else:
            print("One or more of the fields were left empty")
            messagebox.showerror("Error", "Please provide both an .xlsx file and a .csv file.\nYou can create a new .xlsx file if you don't have one")
    
    def enable_xlsx_field():
        xlsx_entry.configure(state="normal")
        new_xlsx_name.delete(0, tk.END)
        new_xlsx_name.configure(state="disabled")
        browse_button.configure(state="normal")

    def disable_xlsx_field():
        #Clear out the excel field and disable it
        xlsx_entry.delete(0, tk.END)
        xlsx_entry.configure(state="disabled")
        new_xlsx_name.configure(state="normal")
        
        #Enable the new excel field
        global new_xlsx_path
        new_xlsx_path = save_file_dialog(".xlsx", ("Excel Files", "*.xlsx"))
        if new_xlsx_path:
            new_xlsx_name.insert(0, new_xlsx_path)
        else:
            # Enable the fields back if no file is made
            enable_xlsx_field()


    initial_popup = tk.Tk()
    initial_popup.title("Provide Files")
    initial_popup.geometry("300x400")       # Default window size

    # Excel file group
    tk.Label(initial_popup, text="Excel (.xlsx) File:").pack(pady=10)
    xlsx_entry = tk.Entry(initial_popup)
    xlsx_entry.pack(pady=5)
    browse_button = tk.Button(initial_popup, text="Browse", command=lambda: 
              xlsx_entry.insert(0, open_file_dialog(("Excel Files", "*.xlsx"))))
    browse_button.pack(pady=5)
    
    #Creating new excel file
    tk.Label(initial_popup, text="Create a new file:").pack(pady=10)
    new_xlsx_name = tk.Entry(initial_popup, state="disabled")   #Initially disabled
    new_xlsx_name.pack(pady=5)
    tk.Button(initial_popup, text="Create new Excel file", command=disable_xlsx_field).pack(pady=5)

    # CSV file group
    tk.Label(initial_popup, text="CSV File:").pack(pady=10)
    csv_entry = tk.Entry(initial_popup)
    csv_entry.pack(pady=5)
    tk.Button(initial_popup, text="Browse", command=lambda: 
              csv_entry.insert(0, open_file_dialog(("CSV Files", "*.csv")))).pack(pady=10)
    
    tk.Button(initial_popup, text="Submit", command=submit_files).pack(pady=20)

    initial_popup.mainloop()

def create_main_gui():
    #Main GUI and logic
    print("Main GUI not implemented yet - Exiting")
    exit()
    
xlsx_path = ""
csv_path = ""
new_xlsx_path = ""

create_initial_popup()
