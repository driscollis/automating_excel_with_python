# print_area.py

from openpyxl import Workbook


def print_area(path):
    workbook = Workbook()
    sheet = workbook.active

    # Add 101 rows of data
    sheet.append(["Book", "Kindle", "Paperback"])
    for row in range(100):
        sheet.append(["Python 101", 9.99, 15.99])

    # Only print the data in A1 - B20
    sheet.print_area = "A1:B20"

    workbook.save(path)


if __name__ == "__main__":
    print_area("print_area.xlsx")
