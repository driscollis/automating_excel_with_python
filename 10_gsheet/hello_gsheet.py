# hello_gsheet.py

import gspread
from oauth2client.service_account import ServiceAccountCredentials


def authenticate(credentials):
    """
    Authenticate with Google and get the client object
    """
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials, scope)
    client = gspread.authorize(creds)
    return client


def main():
    """
    Create a Google Sheet
    """
    # Pass in the file location for the JSON file you created
    client = authenticate("pyspread.json")
    try:
        workbook = client.open("test")
        print("Test Sheet opened")
    except:
        workbook = client.create("test")
        print("Test Sheet created")
    sheet = workbook.sheet1
    sheet.update("A1", [[1, 2], [3, 4]])
    workbook.share("YOUR_EMAIL_ADDRESS", perm_type="user", role="writer")


if __name__ == "__main__":
    main()
