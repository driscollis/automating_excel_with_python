# freezing_panes.py

from openpyxl import Workbook


def freeze(path, row_to_freeze):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Freeze"
    sheet.freeze_panes = row_to_freeze
    headers = ["Name", "Address", "State", "Zip"]
    sheet["A1"] = headers[0]
    sheet["B1"] = headers[1]
    sheet["C1"] = headers[2]
    sheet["D1"] = headers[3]
    data = [dict(zip(headers, ("Mike", "123 Storm Dr", "IA", "50000"))),
            dict(zip(headers, ("Ted", "555 Tornado Alley", "OK", "90000")))]
    row = 2
    for d in data:
        sheet[f'A{row}'] = d["Name"]
        sheet[f'B{row}'] = d["Address"]
        sheet[f'C{row}'] = d["State"]
        sheet[f'D{row}'] = d["Zip"]
        row += 1
    workbook.save(path)


if __name__ == "__main__":
    freeze("freeze.xlsx", row_to_freeze="A2")