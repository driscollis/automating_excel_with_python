# named_style.py

from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, NamedStyle


def named_style(path):
    workbook = Workbook()
    sheet = workbook.active

    red = "00FF0000"
    font = Font(bold=True, size=22)
    thick = Side(style="thick", color=red)
    border = Border(left=thick, right=thick, top=thick, bottom=thick)
    named_style = NamedStyle(name="highlight", font=font, border=border)

    sheet["A1"].value = "Hello"
    sheet["A1"].style = named_style

    sheet["A2"].value = "from"
    sheet["A3"].value = "OpenPyXL"

    workbook.save(path)


if __name__ == "__main__":
    named_style("named_style.xlsx")