# data_validation.py

import xlsxwriter


def validate(path):
    workbook = xlsxwriter.Workbook(path)
    sheet = workbook.add_worksheet()

    sheet.set_column('A:A', 34)
    sheet.set_column('B:B', 15)

    header_format = workbook.add_format(
        {
            "border": 1,
            "bg_color": "#33f3ff",
            "bold": True,
            "text_wrap": True,
            "valign": "vcenter",
            "indent": 1,
        }
    )

    sheet.write("A1", "Data Validation Example", header_format)
    sheet.write("B1", "Enter Values Here", header_format)

    sheet.write("A3", "Enter an integer between 1 and 15")
    sheet.data_validation(
        "B3",
        {"validate": "integer", "criteria": "between",
         "minimum": 1, "maximum": 15},
    )
    workbook.close()

if __name__ == "__main__":
    validate("validation.xlsx")
