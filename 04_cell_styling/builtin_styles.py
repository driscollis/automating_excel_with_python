# builtin_styles.py

from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, GradientFill, Alignment


def builtin_styles(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"].value = "Hello"
    sheet["A1"].style = "Title"
    
    sheet["A2"].value = "from"
    sheet["A2"].style = "Headline 1"
    
    sheet["A3"].value = "OpenPyXL"
    sheet["A3"].style = "Headline 2"
    
    workbook.save(path)


if __name__ == "__main__":
    builtin_styles("builtin_styles.xlsx")