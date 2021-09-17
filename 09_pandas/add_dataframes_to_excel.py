# add_dataframes_to_excel.py

import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows


def add_dataframes_to_excel(path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    df = pd.DataFrame(
        [[100, 433, 10], [34, 10, 0], [75, 125, 5]],
        index=["Python 101", "Python 201", "wxPython"],
        columns=["Amazon", "Leanpub", "Gumroad"],
    )

    for row in dataframe_to_rows(df):
        sheet.append(row)
    workbook.save(path)


if __name__ == "__main__":
    add_dataframes_to_excel("df_to_excel.xlsx")
