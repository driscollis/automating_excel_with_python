# using_colorscale.py

from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScale, FormatObject, Rule
from openpyxl.styles import Color


def applying_colorscale(path, output_path):
    workbook = load_workbook(filename=path)
    sheet = workbook.active

    first = FormatObject(type='min')
    last = FormatObject(type="max")

    colors = [Color('AA0000'),   # red
              Color('00AA00')]   # green
    color_scale = ColorScale(cfvo=[first, last], color=colors)

    rule = Rule(type="colorScale", colorScale=color_scale)

    sheet.conditional_formatting.add("B1:B100", rule)
    workbook.save(output_path)


if __name__ == "__main__":
    applying_colorscale("ratings.xlsx",
                        output_path="colorscale.xlsx")