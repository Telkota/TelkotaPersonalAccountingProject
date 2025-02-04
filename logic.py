from pyexcel_ods3 import get_data

def check_for_overview_sheet(filename):
    """
    Checks if the given ODS file contains a sheet with a certain name.
    Make changes to the function to adapt it to your own usage.
    The file needs to be within the same folder as the script.
    
    Arguments:
        filename: The path to the ODS file.
    
    Returns:
        True if the specified sheet exists, False otherwise.
    """
    try:
        doc = get_data(filename)
        return "Oversikt" in doc

    except Exception as e:
        print(f"Error loading document: {e}")
        return False
    
# Test to see if it works
filename = "test.ods"
if check_for_overview_sheet(filename):
    print("The file contains the correct sheet.")
else:
    print("The file doesn't conain the correct sheet, or something went wrong. Exiting program")
    exit()
