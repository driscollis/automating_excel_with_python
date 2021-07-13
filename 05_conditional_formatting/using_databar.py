# using_databar.py

from openpyxl import load_workbook
from openpyxl.formatting.rule import DataBar, FormatObject, Rule


def applying_databar(path, output_path):
    workbook = load_workbook(filename=path)
    sheet = workbook.active

    first = FormatObject(type='num', val=1)
    last = FormatObject(type="num", val=5)
    green = "00AA00"
    data_bar = DataBar(cfvo=[first, last], color=green,
                       showValue=None, minLength=None, maxLength=None)

    rule = Rule(type='dataBar', dataBar=data_bar)
    sheet.conditional_formatting.add("B1:B100", rule)
    workbook.save(output_path)


if __name__ == "__main__":
    applying_databar("ratings.xlsx",
                     output_path="databar.xlsx")