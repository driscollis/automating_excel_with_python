# style_merged_cell.py

from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, GradientFill, Alignment


def merge_style(path):
    workbook = Workbook()
    sheet = workbook.active
    cell_range = "A2:G4"
    sheet.merge_cells(cell_range)
    top_left_cell = sheet["A2"]

    light_purple = "00CC99FF"
    green = "00008000"
    thin = Side(border_style="thin", color=light_purple)
    double = Side(border_style="double", color=green)

    top_left_cell.value = "Hello from PyOpenXL"
    for column in sheet[cell_range]:
        for cell in column:
            cell.border = Border(top=double, left=thin, right=thin,
                                bottom=double)
    top_left_cell.fill = GradientFill(stop=("000000", "FFFFFF"))
    top_left_cell.font  = Font(b=True, color="FF0000", size=16)
    top_left_cell.alignment = Alignment(horizontal="center",
                                        vertical="center")
    workbook.save(path)


if __name__ == "__main__":
    merge_style("merged_style.xlsx")