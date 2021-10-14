# sparklines.py

import xlsxwriter


def sparklines(path):
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet(name="Sparky")
    spark_styles = [{'range': 'Sparky!A1:E1',
                     'markers': True},
                    {'range': 'Sparky!A2:E2',
                    'type': 'column',
                    'style': 12},
                    {'range': 'Sparky!A3:E3',
                     'type': 'win_loss',
                     'negative_points': True}
                    ]

    data = [[-5, 5, 3, -2, 0,],
            [50, 40, 44, 20, 35],
            [0, 1, -1, 0, 1]]
    for row in range(len(data)):
        sheet.write_row(f"A{row+1}", data[row])
        # Add sparklines
        sheet.add_sparkline(f"F{row+1}", spark_styles[row])

    workbook.close()


if __name__ == "__main__":
    sparklines("sparklines.xlsx")
