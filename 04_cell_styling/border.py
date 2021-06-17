# border.py

from openpyxl import Workbook
from openpyxl.styles import Border, Side


def border(path):
    pink = "00FF00FF"
    green = "00008000"
    thin = Side(border_style="thin", color=pink)
    double = Side(border_style="double", color=green)

    workbook = Workbook()
    sheet = workbook.active

    sheet["A1"] = "Hello"
    sheet["A1"].border = Border(top=double, left=thin, right=thin, bottom=double)
    sheet["A2"] = "from"
    sheet["A3"] = "OpenPyXL"
    sheet["A3"].border = Border(top=thin, left=double, right=double, bottom=thin)
    workbook.save(path)


if __name__ == "__main__":
    border("border.xlsx")