# using_iconset.py

from openpyxl import load_workbook
from openpyxl.formatting.rule import IconSet, FormatObject, Rule


def applying_iconset(path, output_path):
    workbook = load_workbook(filename=path)
    sheet = workbook.active

    first = FormatObject(type='num', val=0)
    mid = FormatObject(type="num", val=3)
    last = FormatObject(type="num", val=5)

    iconset = IconSet(iconSet='3TrafficLights1', cfvo=[first, mid, last],
                      showValue=None, percent=None, reverse=None)

    rule = Rule(type="iconSet", iconSet=iconset)

    sheet.conditional_formatting.add("B1:B100", rule)
    workbook.save(output_path)


if __name__ == "__main__":
    applying_iconset("ratings.xlsx",
                     output_path="iconset.xlsx")