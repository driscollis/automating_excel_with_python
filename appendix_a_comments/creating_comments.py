# creating_comments.py

from openpyxl import Workbook
from openpyxl.comments import Comment


def main(filename, cell):
    workbook = Workbook()
    sheet = workbook.active

    comment = Comment(text="Comment written by OpenPyXL",
                      author="OpenPyXL")
    sheet[cell].comment = comment
    comment.width = 300
    comment.height = 30

    print(comment.text)
    print(f"Comment author: {comment.author}")
    workbook.save(filename)


if __name__ == "__main__":
    main("automated_comments.xlsx", "B1")