# print_options_header.py

from openpyxl import Workbook


def add_header(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "Hello"
    sheet["A2"] = "from"
    sheet["A3"] = "OpenPyXL"

    sheet.oddHeader.right.text = "Page &[Page] of &N"
    sheet.oddHeader.right.size = 14
    sheet.oddHeader.right.font = "Tahoma,Bold"
    sheet.oddHeader.right.color = "CC3366"

    workbook.save(path)


if __name__ == "__main__":
    add_header("print_options.xlsx")
