from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.units import pixels_to_EMU

workbook = Workbook()
sheet = workbook.active

img = Image("my_image.jpg")
drawing = sheet.add_image(img, "B2")

# Resizing the image (in EMUs)
img.width= 200
img.height = 150

# Reposition the image (in EMUs) ?
img.left= pixels_to_EMU(500)
img.top = pixels_to_EMU(50)

workbook.save("drawing_sample.xlsx")