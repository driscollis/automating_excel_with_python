# bubble_chart.py

from openpyxl import Workbook
from openpyxl.chart import Series, Reference, BubbleChart


def main():
    workbook = Workbook()
    sheet = workbook.active

    rows = [
        ("Number of Products", "Sales in USD", "Market share"),
        (14, 12200, 15),
        (20, 60000, 33),
        (18, 24400, 10),
        (22, 32000, 42),
        (),
        (12, 8200, 18),
        (15, 50000, 30),
        (19, 22400, 15),
        (25, 25000, 50),
    ]

    for row in rows:
        sheet.append(row)

    chart = BubbleChart()
    chart.style = 18

    # add the first series of data
    xvalues = Reference(sheet, min_col=1, min_row=2, max_row=5)
    yvalues = Reference(sheet, min_col=2, min_row=2, max_row=5)
    size = Reference(sheet, min_col=3, min_row=2, max_row=5)
    series = Series(values=yvalues, xvalues=xvalues, zvalues=size, title="2013")
    chart.series.append(series)

    # add the second series of data
    xvalues = Reference(sheet, min_col=1, min_row=7, max_row=10)
    yvalues = Reference(sheet, min_col=2, min_row=7, max_row=10)
    size = Reference(sheet, min_col=3, min_row=7, max_row=10)
    series = Series(values=yvalues, xvalues=xvalues, zvalues=size, title="2014")
    chart.series.append(series)

    # place the chart starting in cell E1
    sheet.add_chart(chart, "E1")
    workbook.save("bubble_chart.xlsx")

if __name__ == "__main__":
    main()