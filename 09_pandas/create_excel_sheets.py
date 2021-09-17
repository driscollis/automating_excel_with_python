# create_excel_sheets.py

import pandas as pd


def create_multiple_sheets(path):
    df = pd.DataFrame(
        [[100, 433, 10], [34, 10, 0], [75, 125, 5]],
        index=["Python 101", "Python 201", "wxPython"],
        columns=["Amazon", "Leanpub", "Gumroad"],
    )
    df2 = pd.DataFrame(
        [[150, 233, 5], [5, 15, 0], [10, 120, 5]],
        index=[
            "Jupyter Notebook",
            "Python Interview",
            "Pillow: Image Processing with Python",
        ],
        columns=["Amazon", "Leanpub", "Gumroad"],
    )
    with pd.ExcelWriter(path) as writer:
        df.to_excel(writer, sheet_name="Books")
        df2.to_excel(writer, sheet_name="More Books")


if __name__ == "__main__":
    create_multiple_sheets("pandas_to_excel_sheets.xlsx")
