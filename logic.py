import openpyxl as op
from openpyxl.styles import NamedStyle, Font
from datetime import datetime
import csv

#For testing for now
ods_filename = "test.ods"
xlsx_filename = "test.xlsx"

def check_for_overview_sheet(filename):
    """
    Checks if the given xlsx file contains a sheet with a certain name.
    Make changes to the function to adapt it to your own usage.
    The file needs to be within the same folder as the script.
    
    Arguments:
        filename: The name of the file within the same directory as the script.
    
    Returns:
        True if the specified sheet exists, False otherwise.
    """
    try:
        doc = op.load_workbook(filename)
        return "Oversikt" in doc.sheetnames

    except Exception as e:
        print(f"Error loading document: {e}")
        return False
    
def filter_csv(filename):
    """
    Opens up a CSV file and stores the columns specified by the code.
    Tweak the code to your own liking. 
    
    Arguments:
        Filename: The name of the file within the same directory as the script.
    
    Returns:
        A list of dictionary objects containing the information within the CSV.
    """

    filtered_transactions = []

    with open(filename, newline="",) as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            #print(row)
            # Check if there is a value in "Beløp ut" for converting to float
            if row["Beløp ut"]:
                # Remove the - from the value and convert to a float
                amount = float(row["Beløp ut"].replace("-", ""))
            # Check if there is a value in "Beløp inn" for converting to float
            elif row["Beløp inn"]:
                amount = float(row["Beløp inn"])
            else:
                #If there is no Amount out or in, then there is no more lines to process.
                #Cuts out the extra information in my particular CSV
                break

            #Convert the date into a datetime obj
            try:
                date_obj = datetime.strptime(row["Utført dato"], "%d.%m.%Y")
            except ValueError:
                print(f"Invalid date format: {row['Utført dato']}")
                continue

            # Write the transaction to a dictionary object to store in the list
            transaction = {
                "Dato": date_obj,
                "Beskrivelse": row["Beskrivelse"],
                "Beløp": amount
            }

            filtered_transactions.append(transaction)
    sorted_transactions = sorted(filtered_transactions, key=lambda x: x["Dato"])    
    print(sorted_transactions[:5])
    return sorted_transactions

def process_transactions(transactions):
    """
    Takes in a list of transactions and let's the user add a category and comment to the transaction.
    
    Arguments:
        transactions: Should be a list of dictionaries returned from the function 'filter_csv'
    
    Returns:
        A new list of dictionaries with Date, Amount, Comment and Category
    """
    #For now, add it to "Annet" on category and "Test" for comment to see if it works
    new_transactions = []
    for entry in transactions:
        #formatted_date = f"{entry["Dato"].day}.{entry["Dato"].month}"
        new_entry = {
            "Dato": entry["Dato"],
            "Beløp": entry["Beløp"],
            "Beskrivelse": "Test",
            "Kategori": "Annet"
        }
        new_transactions.append(new_entry)
    print(new_transactions[:5])
    return new_transactions

def save_document(transactions):
    """
    Takes in a list of transactions with Date, Amount, Comment and Category to append to a xlsx file
    
    Arguments:
        transactions: Needs to be a list of dictionaries with Date, Amount, Comment and Category - Use process_transactions()
        
    Returns:
        Nothing - The function will try to append the transactions into the document and try to save it.
    """
    try:
        workbook = op.load_workbook(xlsx_filename)
    except FileNotFoundError:
        print("File not found")
        return
    
    # Check if there is a style named "date_style" in the file already
    if "date_style" not in workbook.named_styles:
        # Add the style to get the desired format and font
        date_style = NamedStyle(name="date_style", number_format="DD.MM.YYYY")
        date_style.font = Font(name="Arial")
        workbook.add_named_style(date_style)

    for entry in transactions:
        category = entry["Kategori"]
        try:
            #Get data from the corresponding category sheet.
            worksheet = workbook[category]
        except KeyError:
            worksheet = workbook.create_sheet(title=category)

        #Find the next available row to append data
        next_row = worksheet.max_row + 1

        date_cell = worksheet.cell(row=next_row, column=1, value=entry["Dato"])
        date_cell.style = "date_style"
        worksheet.cell(row=next_row, column=2, value=entry["Beløp"])
        worksheet.cell(row=next_row, column=3, value=entry["Beskrivelse"])
    
    workbook.save(xlsx_filename)

# Test to see if it works
csv_filename = "transaksjoner_test.csv"
filtrert = filter_csv(csv_filename)
klargjort = process_transactions(filtrert)
save_document(klargjort)
