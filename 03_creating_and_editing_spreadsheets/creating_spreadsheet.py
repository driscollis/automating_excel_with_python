# creating_spreadsheet.py

from openpyxl import Workbook


def create_workbook(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Hello"
    sheet2 = workbook.create_sheet(title="World")
    workbook.save(path)


if __name__ == "__main__":
    create_workbook("hello.xlsx")