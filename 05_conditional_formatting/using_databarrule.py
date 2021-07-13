# using_databarrule.py

from openpyxl import load_workbook
from openpyxl.formatting.rule import DataBarRule


def applying_databar_rule(path, output_path):
    workbook = load_workbook(filename=path)
    sheet = workbook.active

    red = "AA0000"
    data_bar_rule = DataBarRule(start_type="num",
                                start_value=1,
                                end_type="num",
                                end_value="5",
                                color=red)

    sheet.conditional_formatting.add("B1:B100", data_bar_rule)
    workbook.save(output_path)


if __name__ == "__main__":
    applying_databar_rule("ratings.xlsx",
                          output_path="databarrule.xlsx")