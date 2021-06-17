# insert_image.py

from openpyxl import Workbook
from openpyxl.drawing.image import Image


def insert_image(path, image_path):
    workbook = Workbook()
    sheet = workbook.active
    img = Image("logo.png")
    sheet.add_image(img, "B1")
    workbook.save(path)


if __name__ == "__main__":
    insert_image("logo.xlsx", "logo.png")