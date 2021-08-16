# radar_chart.py

from openpyxl import Workbook
from openpyxl.chart import RadarChart, Reference


def main(filename):
    workbook = Workbook()
    sheet = workbook.active

    data = [
        ['Month', "Bulbs", "Seeds", "Flowers", "Trees & shrubs"],
        ['Jan', 0, 2500, 500, 0,],
        ['Feb', 0, 5500, 750, 1500],
        ['Mar', 0, 9000, 1500, 2500],
        ['Apr', 0, 6500, 2000, 4000],
        ['May', 0, 3500, 5500, 3500],
        ['Jun', 0, 0, 7500, 1500],
        ['Jul', 0, 0, 8500, 800],
        ['Aug', 1500, 0, 7000, 550],
        ['Sep', 5000, 0, 3500, 2500],
        ['Oct', 8500, 0, 2500, 6000],
        ['Nov', 3500, 0, 500, 5500],
        ['Dec', 500, 0, 100, 3000 ],
    ]


    for row in data:
        sheet.append(row)

    chart = RadarChart()
    chart.type = "filled"
    labels = Reference(sheet, min_col=1, min_row=2, max_row=13)
    data = Reference(sheet, min_col=2, max_col=5, min_row=1, max_row=13)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.style = 26
    chart.title = "Garden Centre Sales"
    chart.y_axis.delete = True

    sheet.add_chart(chart, "G2")

    workbook.save(filename)


if __name__ == "__main__":
    main("radar_chart.xlsx")