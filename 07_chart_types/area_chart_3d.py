# area_chart_3d.py

from openpyxl import Workbook
from openpyxl.chart import AreaChart3D, Reference

def main(filename):
    wb = Workbook()
    sheet = wb.active

    rows = [
        ["Book", "Kindle", "Paperback"],
        [2, 30, 40],
        [3, 25, 40],
        [4 ,30, 50],
        [5 ,10, 30],
        [6,  5, 25],
        [7 ,10, 50],
    ]

    for row in rows:
        sheet.append(row)

    chart = AreaChart3D()
    chart.title = "Area Chart 3D"
    chart.x_axis.title = "Books"
    chart.y_axis.title = "Copies Sold"
    chart.legend = None

    cats = Reference(sheet, min_col=1, min_row=1, max_row=7)
    data = Reference(sheet, min_col=2, min_row=1, max_col=3, max_row=7)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    sheet.add_chart(chart, "E2")

    wb.save(filename)

if __name__ == "__main__":
    main("area_chart_3d.xlsx")