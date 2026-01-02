# openpyxl - 6. Images and Drawings

- [openpyxl - 6. Images and Drawings](#openpyxl---6-images-and-drawings)
  - [6.1 Adding Images to Worksheets](#61-adding-images-to-worksheets)
  - [6.2 Working with Drawing Objects](#62-working-with-drawing-objects)
  - [6.3 Resizing and Positioning Images](#63-resizing-and-positioning-images)
  - [6.4 Image Formats](#64-image-formats)

## 6.1 Adding Images to Worksheets

```python
from openpyxl import Workbook
from openpyxl.drawing.image import Image

workbook = Workbook()
sheet = workbook.active

img = Image("my_image.jpg")

sheet.add_image(img, "B2")

workbook.save("image_sample.xlsx")
```

![add-image](img/add-image.png)

Image Source: https://clipart.com/

Source Code Reference:

- `drawing/image.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/drawing/image.py

## 6.2 Working with Drawing Objects

```python
```

## 6.3 Resizing and Positioning Images

```python
```

## 6.4 Image Formats

```python
```

---

Last Updated at: 12/28/2025, 8:17:59 AM 