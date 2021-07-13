# using_colorscalerule.py

from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScaleRule


def applying_colorscalerule(path, output_path):
    workbook = load_workbook(filename=path)
    sheet = workbook.active

    red = "AA0000"
    yellow = "00FFFF00"
    green = "00AA00"

    color_scale_rule = ColorScaleRule(start_type="num",
                                      start_value=1,
                                      start_color=red,
                                      mid_type="num",
                                      mid_value=3,
                                      mid_color=yellow,
                                      end_type="num",
                                      end_value=5,
                                      end_color=green)

    sheet.conditional_formatting.add("B1:B100", color_scale_rule)
    workbook.save(output_path)


if __name__ == "__main__":
    applying_colorscalerule("ratings.xlsx",
                        output_path="colorscalerule.xlsx")