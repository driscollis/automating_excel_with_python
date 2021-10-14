# formatting_spreadsheet.py

import xlsxwriter


def format_data(path):
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet(name="Hello")
    data = [["Python 101", 1000, 9.99],
            ["Jupyter Notebook 101", 400, 14.99],
            ["ReportLab: PDF Processing", 250, 24.99]
    ]

    # Add formatting objects
    bold = workbook.add_format({'bold': True})
    money = workbook.add_format({'num_format': '$#,##0.00'})

    # Add headers
    sheet.write("A1", "Book Title", bold)
    sheet.write("B1", "Copies Sold", bold)
    sheet.write("C1", "Book Price", bold)

    row = 1
    col = 0
    for book, sales, amount in data:
        sheet.write(row, col, book)
        sheet.write(row, col + 1, sales)
        sheet.write(row, col + 2, amount, money)
        row += 1

    workbook.close()


if __name__ == "__main__":
    format_data("formatting.xlsx")