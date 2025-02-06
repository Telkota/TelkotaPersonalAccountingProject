import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from logic import *
import os
from datetime import datetime

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
    def add_to_completed():
        selected_item = pending_tree.selection()
        if selected_item:
            item = pending_tree.item(selected_item)
            values = item["values"]
            # Get the user comment and category
            comment = comment_entry.get("1.0", tk.END).strip()
            category = category_var.get()

            if not comment or category == "Velg Kategori":
                messagebox.showerror("Error", "Please add a comment and select a category")
                return
            
            # Add the item to the completed tree
            completed_tree.insert("", "end", values=(values[0], values[1], category, comment))
            # Remove the item from the entry field
            pending_tree.delete(selected_item)

            # Clear the comment field
            comment_entry.delete("1.0", tk.END)
            category_var.set("Velg Kategori")

            # Check if all transactions are completed
            if not pending_tree.get_children():
                submit_button.configure(state="normal")
    
    def add_category():
        new_category = simpledialog.askstring("New Category", "Enter the new category name:")
        if new_category:
            # Lower and Capitalize the input
            sanitized_input = new_category.lower().capitalize()
            add_new_category(xlsx_path, sanitized_input)
            categories = get_excel_categories(xlsx_path)
            category_menu["values"] = categories
            category_var.set(sanitized_input)
    
    def submit_all():
        transactions = []
        for item in completed_tree.get_children():
            values = completed_tree.item(item)["values"]
            transaction = {
                "Dato": values[4],
                "Beløp": values[1],
                "Kategori": values[2],
                "Beskrivelse": values[3]
            }
            transactions.append(transaction)
        
        save_document(xlsx_path, transactions)
        messagebox.showinfo("Success", "All transactions has been saved successfully")
        os.startfile(xlsx_path)
        exit()      # Exit after the user has clicked ok
    
    #Main GUI
    main_gui = tk.Tk()
    main_gui.title("Personal Accounting")
    main_gui.geometry("800x900")
    main_gui.configure(bg="#D3D3D3")

    # Grid configuration for center-alignment
    main_gui.grid_rowconfigure(0, weight=1)
    main_gui.grid_rowconfigure(1, weight=1)
    main_gui.grid_rowconfigure(2, weight=1)
    main_gui.grid_rowconfigure(3, weight=1)
    main_gui.grid_rowconfigure(4, weight=1)
    main_gui.grid_rowconfigure(5, weight=1)
    main_gui.grid_rowconfigure(6, weight=1)
    main_gui.grid_rowconfigure(7, weight=1)
    main_gui.grid_rowconfigure(8, weight=1)
    main_gui.grid_columnconfigure(0, weight=1)
    main_gui.grid_columnconfigure(2, weight=1)
    main_gui.grid_columnconfigure(3, weight=1)


    # Treeview for pending transactions
    tk.Label(main_gui, text="Pending", font=("Arial", 14), bg="#D3D3D3").grid(row=0, column=0, columnspan=3, sticky="nsew")
    pending_tree = ttk.Treeview(main_gui, columns=("Dato", "Beløp", "Beskrivelse"), show="headings")
    pending_tree.heading("Dato", text="Dato")
    pending_tree.heading("Beløp", text="Beløp")
    pending_tree.heading("Beskrivelse", text="Beskrivelse")
    pending_tree.column("Dato", width=50)
    pending_tree.column("Beløp", width=50)
    pending_tree.column("Beskrivelse", width=400)
    pending_tree.grid(row=1, column=0, columnspan=3, padx=40, pady=20, sticky="nsew")

    pending_data = filter_csv(csv_path)
    for item in pending_data:
        date_obj = item["Dato"]
        formatted_date = date_obj.strftime("%d.%m")
        values = (formatted_date, item["Beløp"], item["Beskrivelse"], date_obj)
        pending_tree.insert("", "end", values=values)

    # Treeview for completed transactions
    tk.Label(main_gui, text="Completed", font=("Arial", 14), bg="#D3D3D3").grid(row=2, column=0, columnspan=3, sticky="nsew")
    completed_tree = ttk.Treeview(main_gui, columns=("Dato", "Beløp", "Kategori", "Kommentar"), show="headings")
    completed_tree.heading("Dato", text="Dato")
    completed_tree.heading("Beløp", text="Beløp")
    completed_tree.heading("Kategori", text="Kategori")
    completed_tree.heading("Kommentar", text="Kommentar")
    completed_tree.column("Dato", width=50)
    completed_tree.column("Beløp", width=50)
    completed_tree.column("Kategori", width=100)
    completed_tree.column("Kommentar", width=300)
    completed_tree.grid(row=3, column=0, columnspan=3, padx=40, pady=20, sticky="nsew")

    # Comment entry field
    tk.Label(main_gui, text="Kommentar:", font=("Arial", 12), bg="#D3D3D3").grid(row=4, column=0, columnspan=3, pady=5, sticky="nsew")
    comment_entry = tk.Text(main_gui, width=40, height=4)
    comment_entry.grid(row=5, column=0, columnspan=2,padx=40, pady=5, sticky="nsew")

    # Category dropdown menu
    tk.Label(main_gui, text="Kategori:", font=("Arial", 12), bg="#D3D3D3").grid(row=6, column=0, pady=5, sticky="nsew")
    categories = get_excel_categories(xlsx_path)
    category_var = tk.StringVar(value="Velg Kategori")
    category_menu = ttk.Combobox(main_gui, textvariable=category_var, values=categories)
    category_menu.grid(row=7, column=0, rowspan=2, padx=40, pady=10, sticky="nsew")

    #Add new category button
    add_category_button = tk.Button(main_gui, text="New Category", command=add_category)
    add_category_button.grid(row=7, column=1, rowspan=2, padx=20, pady=10, sticky="nsew")

    # Add to completed button
    add_button = tk.Button(main_gui, text="Add", command=add_to_completed)
    add_button.grid(row=5, column=2, padx=20, pady=10, sticky="nsew")

    # Submit button - Submits the completed list if there are no more transactions
    submit_button = tk.Button(main_gui, text="Submit all", state="disabled", command=submit_all)
    submit_button.grid(row=7, column=2, rowspan=2, padx=20, pady=10, sticky="nsew")

    main_gui.mainloop()
    
xlsx_path = ""
csv_path = ""
new_xlsx_path = ""

create_initial_popup()
