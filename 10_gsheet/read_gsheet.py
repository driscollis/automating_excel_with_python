# read_gsheet.py

import gspread
from oauth2client.service_account import ServiceAccountCredentials


def authenticate(credentials):
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials, scope)
    client = gspread.authorize(creds)
    return client


def main():
    client = authenticate("pyspread.json")
    try:
        workbook = client.open("test")
        print("Test Sheet opened")
    except:
        print("Error opening test spreadsheet")
        return
    sheet = workbook.sheet1
    print("Worksheets: " + str(workbook.worksheets()))
    print(f"All values on row 1: {sheet.row_values(1)}")
    print(f"All values in column 1: {sheet.col_values(1)}")
    print(f"All values in worksheet: {sheet.get_all_values()}")
    # Get all values from worksheet as a list of dictionaries
    values = sheet.get_all_records()
    print(f"All records: {values=}")


if __name__ == "__main__":
    main()
