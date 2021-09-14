# print_titles.py

from openpyxl import Workbook


def print_titles(path):
    workbook = Workbook()
    sheet = workbook.active
    
    # Add 101 rows of data
    sheet.append(["Book", "Kindle", "Paperback"])
    for row in range(100):
        sheet.append(["Python 101", 9.99, 15.99])
        
    # Set the first three columns as the title columns
    sheet.print_title_cols = "A:C"
    # Set the first row as the title row
    sheet.print_title_rows = "1:1"

    workbook.save(path)


if __name__ == "__main__":
    print_titles("print_titles.xlsx")
