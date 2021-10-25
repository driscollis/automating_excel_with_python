# creating_spreadsheet.py

from openpyxl import Workbook


def create_workbook(path):
    workbook = Workbook()
    workbook.save(path)


if __name__ == "__main__":
    create_workbook("hello.xlsx")