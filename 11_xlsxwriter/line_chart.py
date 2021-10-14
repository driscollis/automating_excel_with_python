# line_chart.py

import xlsxwriter


def chart(path):
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet(name="Hello")

    data = [15, 65, 50, 20, 5, 50]
    sheet.write_column("A1", data)

    # Create the chart object
    chart = workbook.add_chart({"type": "line"})
    chart.add_series({"values": "=Hello!$A$1:$A$6"})

    # Add the chart to the spreadsheet
    sheet.insert_chart("C1", chart)
    workbook.close()


if __name__ == "__main__":
    chart("line_chart.xlsx")
