# chartsheet_demo.py

from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference


def main(filename):
    wb = Workbook()
    ws = wb.active
    chart_sheet = wb.create_chartsheet()

    rows = [
        ["Python", 50],
        ["C++", 35],
        ["Java", 10],
        ["R", 5]
    ]

    for row in rows:
        ws.append(row)

    chart = PieChart()
    labels = Reference(ws, min_col=1, min_row=1, max_row=4)
    data = Reference(ws, min_col=2, min_row=1, max_row=5)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = "Programming Languages"

    chart_sheet.add_chart(chart)

    wb.save(filename)

if __name__ == "__main__":
    main("chartsheet.xlsx")