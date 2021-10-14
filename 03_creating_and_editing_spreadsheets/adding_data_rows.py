# adding_data_rows.py

from openpyxl import Workbook


def create_workbook(path):
    workbook = Workbook()
    sheet = workbook.active
    data = [[1, 2, 3],
            ["a", "b", "c"],
            [44, 55, 66]]
    for row in data:
        sheet.append(row)
    workbook.save(path)


if __name__ == "__main__":
    create_workbook("write_rows.xlsx")
