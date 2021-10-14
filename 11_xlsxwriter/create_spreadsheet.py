# create_spreadsheet.py

import xlsxwriter


def create_workbook(path):
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet(name="Hello")
    data = [["Python 101", 1000],
            ["Jupyter Notebook 101", 400],
            ["ReportLab: PDF Processing", 250]
    ]

    row = 0
    col = 0
    for book, sales in data:
        sheet.write(row, col, book)
        sheet.write(row, col + 1, sales)
        row += 1

    sheet2 = workbook.add_worksheet(name="World")
    workbook.close()


if __name__ == "__main__":
    create_workbook("hello.xlsx")