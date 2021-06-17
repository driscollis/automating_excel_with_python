# editing_demo.py

from openpyxl import load_workbook


def edit(path, data):
    workbook = load_workbook(filename=path)
    sheet = workbook.active
    for cell in data:
        current_value = sheet[cell].value
        sheet[cell] = data[cell]
        print(f'Changing {cell} from {current_value} to {data[cell]}')
    workbook.save(path)

if __name__ == "__main__":
    data = {"B1": "Hi", "B5": "Python"}
    edit("inserting.xlsx", data)