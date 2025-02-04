from pyexcel_ods3 import get_data, save_data
from datetime import datetime
import csv

#For testing for now
ods_filename = "test.ods"

def check_for_overview_sheet(filename):
    """
    Checks if the given ODS file contains a sheet with a certain name.
    Make changes to the function to adapt it to your own usage.
    The file needs to be within the same folder as the script.
    
    Arguments:
        filename: The name of the file within the same directory as the script.
    
    Returns:
        True if the specified sheet exists, False otherwise.
    """
    try:
        doc = get_data(filename)
        return "Oversikt" in doc

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
            if row["Beløp inn"]:
                amount = float(row["Beløp inn"])

            #Check for empty rows
            if not any(row[key] for key in row):
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
        formatted_date = f"{entry["Dato"].day}.{entry["Dato"].month}"
        new_entry = {
            "Dato": formatted_date,
            "Beløp": entry["Beløp"],
            "Beskrivelse": "Test",
            "Kategori": "Annet"
        }
        new_transactions.append(new_entry)
    print(new_transactions[:5])
    return new_transactions

def save_document(transactions):
    """
    Takes in a list of transactions with Date, Amount, Comment and Category to append to a ODS document
    
    Arguments:
        transactions: Needs to be a list of dictionaries with Date, Amount, Comment and Category - Use process_transactions()
        
    Returns:
        Nothing - The function will try to append the transactions into the document and try to save it.
    """

    for entry in transactions:
        category = entry["Kategori"]
        try:
            #Get data from the corresponding category sheet.
            existing_data = get_data(ods_filename)[category]
        except KeyError:
            continue

        existing_data.append([entry["Dato"], entry["Beløp"], entry["Beskrivelse"]])

        sheet_data = {category: existing_data}

        save_data(ods_filename, sheet_name=category, data=sheet_data)

# Test to see if it works
csv_filename = "transaksjoner_test.csv"
filtrert = filter_csv(csv_filename)
klargjort = process_transactions(filtrert)
save_document(klargjort)
