# insert_demo.py

from openpyxl import Workbook


def inserting_cols_rows(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Hello"
    sheet["A2"] = "from"
    sheet["A3"] = "OpenPyXL"
    # insert a column before A
    sheet.insert_cols(idx=1)
    # insert 2 rows starting on the second row
    sheet.insert_rows(idx=2, amount=2)
    workbook.save(path)


if __name__ == "__main__":
    inserting_cols_rows("inserting.xlsx")
