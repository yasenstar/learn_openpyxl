from openpyxl import Workbook
from openpyxl.drawing.image import Image

workbook = Workbook()
sheet = workbook.active

img = Image("my_image.jpg")

sheet.add_image(img, "B2")

workbook.save("image_sample.xlsx")