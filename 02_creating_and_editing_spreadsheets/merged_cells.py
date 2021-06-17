# merged_cells.py

from openpyxl import Workbook
from openpyxl.styles import Alignment


def create_merged_cells(path, value):
    workbook = Workbook()
    sheet = workbook.active
    sheet.merge_cells("A2:E2")
    top_left_cell = sheet["A2"]
    top_left_cell.alignment = Alignment(horizontal="center",
                                        vertical="center")
    sheet["A2"] = value
    workbook.save(path)


if __name__ == "__main__":
    create_merged_cells("merged.xlsx", "Hello World")