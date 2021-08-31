# reading_comments.py

from openpyxl import load_workbook
from openpyxl.comments import Comment


def main(filename, cell):
    workbook = load_workbook(filename=filename)
    sheet = workbook.active
    comment = sheet[cell].comment
    print(comment)


if __name__ == "__main__":
    main("comments.xlsx", "A1")