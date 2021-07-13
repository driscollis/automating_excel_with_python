# using_iconsetrule.py

from openpyxl import load_workbook
from openpyxl.formatting.rule import IconSetRule


def applying_iconsetrule(path, output_path):
    workbook = load_workbook(filename=path)
    sheet = workbook.active

    icon_set_rule = IconSetRule("5Arrows", "num", [1, 2, 3, 4, 5])
    sheet.conditional_formatting.add("B1:B100", icon_set_rule)
    workbook.save(output_path)


if __name__ == "__main__":
    applying_iconsetrule("ratings.xlsx",
                         output_path="iconsetrule.xlsx")