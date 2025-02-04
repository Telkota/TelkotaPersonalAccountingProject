from pyexcel_ods3 import get_data
import csv

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
                row["Beløp ut"] = float(row["Beløp ut"].replace("-", ""))
            
            # Check if there is a value in "Beløp inn" for converting to float
            if row["Beløp inn"]:
                row["Beløp inn"] = float(row["Beløp inn"])

            # Write the transaction to a dictionary object to store in the list
            transaction = {
                "Dato": row["Utført dato"],
                "Beskrivelse": row["Beskrivelse"],
                "Beløp inn": row["Beløp inn"],
                "Beløp ut": row["Beløp ut"]
            }

            filtered_transactions.append(transaction)

    return filtered_transactions
    
# Test to see if it works
filename = "transaksjoner_test.csv"
transaksjoner = filter_csv(filename)
print(transaksjoner[0:5])
