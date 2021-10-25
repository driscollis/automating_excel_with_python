# doughnut_chart.py

from copy import deepcopy

from openpyxl import Workbook
from openpyxl.chart import DoughnutChart, Reference, Series
from openpyxl.chart.series import DataPoint


def main(filename):
    workbook = Workbook()
    sheet = workbook.active

    data = [
        ['Books', 2019, 2020],
        ['Python 101', 40, 50],
        ['Python 201', 2, 10],
        ['ReportLab', 20, 30],
        ['Jupyter', 30, 40],
    ]

    for row in data:
        sheet.append(row)

    chart = DoughnutChart()
    labels = Reference(sheet, min_col=1, min_row=2, max_row=5)
    data = Reference(sheet, min_col=2, min_row=1, max_row=5)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = "Books sold by title"
    chart.style = 26

    # Cut the first slice out of the doughnut
    slices = [DataPoint(idx=i) for i in range(4)]
    plain, jam, lime, chocolate = slices
    chart.series[0].data_points = slices
    plain.graphicalProperties.solidFill = "FAE1D0"
    jam.graphicalProperties.solidFill = "BB2244"
    lime.graphicalProperties.solidFill = "22DD22"
    chocolate.graphicalProperties.solidFill = "61210B"
    chocolate.explosion = 10

    sheet.add_chart(chart, "E1")

    chart2 = deepcopy(chart)
    chart2.title = None
    data = Reference(sheet, min_col=3, min_row=1, max_row=5)
    series2 = Series(data, title_from_data=True)
    series2.data_points = slices
    chart2.series.append(series2)

    sheet.add_chart(chart2, "E17")

    workbook.save(filename)


if __name__ == "__main__":
    main("doughnut_chart.xlsx")