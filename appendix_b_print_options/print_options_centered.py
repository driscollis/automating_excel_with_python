# print_options_center.py

from openpyxl import Workbook


def center_data(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Hello"
    sheet["A2"] = "from"
    sheet["A3"] = "OpenPyXL"

    sheet.print_options.horizontalCentered = True
    sheet.print_options.verticalCentered = True

    workbook.save(path)


if __name__ == "__main__":
    center_data("print_options_center.xlsx")
