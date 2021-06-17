# delete_demo.py

from openpyxl import Workbook


def deleting_cols_rows(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Hello"
    sheet["B1"] = "from"
    sheet["C1"] = "OpenPyXL"
    sheet["A2"] = "row 2"
    sheet["A3"] = "row 3"
    sheet["A4"] = "row 4"
    # Delete column A
    sheet.delete_cols(idx=1)
    # delete 2 rows starting on the second row
    sheet.delete_rows(idx=2, amount=2)
    workbook.save(path)


if __name__ == "__main__":
    deleting_cols_rows("deleting.xlsx")
