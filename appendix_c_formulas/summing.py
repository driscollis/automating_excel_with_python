# summing.py

from openpyxl import Workbook


def main(filename):
    workbook = Workbook()
    sheet = workbook.active

    # Add data to spreadsheet
    data_rows = [
        ["Book", "Kindle", "Paperback"],
        [1, 9.99, 15.99],
        [2, 9.99, 25.99],
        [3, 9.99, 25.99],
        [4, 4.99, 29.99],
        [5, 14.99, 39.99],
    ]

    for row in data_rows:
        sheet.append(row)

    # Sum up columns
    sheet["A7"] = "Totals"
    sheet["B7"] = "=SUM(B2:B6)"
    sheet["C7"] = "=SUM(C2:C6)"
    workbook.save(filename)


if __name__ == "__main__":
    main("summing.xlsx")