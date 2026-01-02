# openpyxl - 6. Images and Drawings

- [openpyxl - 6. Images and Drawings](#openpyxl---6-images-and-drawings)
  - [6.1 Adding Images to Worksheets](#61-adding-images-to-worksheets)
  - [6.2 Working with Drawing Objects](#62-working-with-drawing-objects)
  - [6.3 Resizing and Positioning Images](#63-resizing-and-positioning-images)
  - [6.4 Image Formats](#64-image-formats)

## 6.1 Adding Images to Worksheets

Adding images to worksheets uses the `openpyxl.drawing.image` module, you need to indicate the path of your image file, support `gif`, `jpeg` and `png` formats.

If you don't input the location of the image in worksheet, it's by default from `anchor = "A1"`, check source code for detail.

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

Extended information: treat `xlsx` as `zip`, you can extract the excel file after renaming its extension to `zip`, then following are the structure and we can see our image file is saved under the sub-folder:

| Level 1 | Level 2 | Level 3 |
| --- | --- | --- |
| ![xls-1](img/xlsx_1.png) | ![xls-2](img/xlsx_2.png) | ![xls-3](img/xlsx_3.png) |

## 6.2 Working with Drawing Objects

Images and other drawing objects in openpyxl are represented as `Drawing` objects.

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