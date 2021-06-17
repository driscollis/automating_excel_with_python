# folding.py

import openpyxl


def folding(path, rows=None, cols=None, hidden=True):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    if rows:
        begin_row, end_row = rows
        sheet.row_dimensions.group(begin_row, end_row, hidden=hidden)

    if cols:
        begin_col, end_col = cols
        sheet.column_dimensions.group(begin_col, end_col, hidden=hidden)

    workbook.save(path)


if __name__ == "__main__":
    folding("folded.xlsx", rows=(1, 5), cols=("C", "F"))
