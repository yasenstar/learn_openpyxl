# openpyxl - 4. Styles and Formatting

- [openpyxl - 4. Styles and Formatting](#openpyxl---4-styles-and-formatting)
  - [4.1 Fonts](#41-fonts)
  - [4.2 Fill Colors](#42-fill-colors)
  - [4.3 Borders](#43-borders)
  - [4.4 Alignment](#44-alignment)
  - [4.5 Number Formats](#45-number-formats)
  - [4.6 Conditional Formatting](#46-conditional-formatting)
    - [4.6.1 Builtin Formats](#461-builtin-formats)
    - [4.6.2 Standard Conditional Formats](#462-standard-conditional-formats)
    - [4.6.3 Formatting Entire Rows (Range)](#463-formatting-entire-rows-range)
  - [4.7 Styles and Themes](#47-styles-and-themes)
  - [4.8 Applying Styles to Cells and Ranges](#48-applying-styles-to-cells-and-ranges)

## 4.1 Fonts

You can control font attributes like name, size, bold, italic, color, etc., using the `Font` class from `openpyxl.styles`, as below:

```python
from openpyxl import Workbook
from openpyxl.styles import Font

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "Styled Text"

my_font = Font(
    name = "Arial",
    size = 20,
    bold = True,
    italic = True,
    underline = "double",
    strikethrough = True,
    shadow = True,
    color = "0000FF"
)

sheet["A1"].font = my_font

workbook.save("font_styles.xlsx")
```

Source code reference:

- `Font` class: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/fonts.py?ref_type=heads#L32

Default Font Styles:

```python
DEFAULT_FONT = Font(name="Calibri", sz=11, family=2, b=False, i=False,
                    color=Color(theme=1), scheme="minor")
```

Referece on `underline`:

```python
    u = NestedNoneSet(values=('single', 'double', 'singleAccounting',
                             'doubleAccounting'))
    underline = Alias("u")
```

Note: the color can be specified using RGB hex codes (e.g., "FF0000" for red) or named colors(?), from documentation (https://openpyxl.pages.heptapod.net/openpyxl/styles.html#colours), you can find `aRGB colours` and `Indexed Colours`. (need test on the named colors)

## 4.2 Fill Colors

Cell background colors are controlled with the `PatternFill` class:

```python
from openpyxl import Workbook
from openpyxl.styles import PatternFill

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "Yellow Fill"

# Solid Fill Style
my_fill1 = PatternFill(
    start_color = "FFFF00",
    end_color = "FFFF00",
    fill_type = "solid"
)

sheet["A1"].fill = my_fill1

sheet["B2"] = "Light Up Blue"

# Solid Fill Style
my_fill2 = PatternFill(
    start_color = "0000FF",
    end_color = "FFFFFF",
    fill_type = "lightUp"
)

sheet["B2"].fill = my_fill2

workbook.save("fill_styles.xlsx")
```

Source code reference:

- `PatternFill(Fill)` is inherit from class `Fill`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/fills.py?ref_type=heads#L68
- `fill_type`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/fills.py?ref_type=heads#L42

Fill Color Naming Conversion Mapping with Excel:

| Fill Pattern Full Name | Short Name | Name in Excel |
| --- | --- | --- |
| FILL_SOLID | solid | Solid |
| FILL_PATTERN_DARKDOWN | darkDown | Reverse Diagonal Stripe |
| FILL_PATTERN_DARKGRAY | darkGray | 75% Gray |
| FILL_PATTERN_DARKGRID | drakGrid | Diagonal Crosshatch |
| FILL_PATTERN_DARKHORIZONTAL | darkHorizontal | Horizontal Stripe |
| FILL_PATTERN_DARKTRELLIS | darkTrellis | Thick Diagonal Crosshatch |
| FILL_PATTERN_DARKUP | darkUp | Diagonal Stripe |
| FILL_PATTERN_DARKVERTICAL | darkVertical | Vertical Stripe |
| FILL_PATTERN_GRAY0625 | gray0625 | 6.25% Gray |
| FILL_PATTERN_GRAY125 | gray125 | 12.5% Gray |
| FILL_PATTERN_LIGHTDOWN | lightDown | Thin Reverse Diagonal Stripe |
| FILL_PATTERN_LIGHTGRAY | lightGray | 25% Gray |
| FILL_PATTERN_LIGHTGRID | lightGrid | Thin Horizontal Crosshatch |
| FILL_PATTERN_LIGHTHORIZONTAL | lightHorizontal | Thin Horizontal Stripe |
| FILL_PATTERN_LIGHTTRELLIS | lightTrellis | Thing Diagonal Crosshatch |
| FILL_PATTERN_LIGHTUP | lightUp | Thin Diagonal Stripe |
| FILL_PATTERN_LIGHTVERTICAL | lightVertical | Thin Vertical Stripe |
| FILL_PATTERN_MEDIUMGRAY | mediumGray | 50% Gray |

## 4.3 Borders

Borders are defined using the `Border`, `Side` classes:

```python
from openpyxl import Workbook
from openpyxl.styles import Border, Side

workbook = Workbook()
sheet = workbook.active

sheet["B2"] = "Thin Border"

sheet["D2"] = "Thick Border"

thin_border = Border(
    left = Side(style = 'thin'),
    right = Side(style = 'thin'),
    top = Side(style = 'thin'),
    bottom = Side(style = 'dashDotDot', color = "0000FF")
)

thick_border = Border(
    left = Side(style = 'thick', color = "FF0000"),
    right = Side(style = 'thick'),
    top = Side(style = 'thick'),
    bottom = Side(style = 'thick', color = "0000FF")
)

sheet["B2"].border = thin_border

sheet["D2"].border = thick_border

workbook.save("border_styles.xlsx")
```

Soruce code reference:

- class `Border` in `borders.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/borders.py?ref_type=heads#L54
  - `border_style = Alias('style')`
- class `Side` in `borders.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/borders.py?ref_type=heads#L33
  - `Side` `style=NoneSet()`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/borders.py?ref_type=heads#L41

Valid Border Side Styles:

```python
style = NoneSet(values=('dashDot','dashDotDot', 'dashed','dotted',
  'double','hair', 'medium', 'mediumDashDot', 'mediumDashDotDot',
  'mediumDashed', 'slantDashDot', 'thick', 'thin'))
```

## 4.4 Alignment

Cell content alignment is controlled with the `Alignment` class:

```python
```

Source code reference:

- class `Alignment` in `alignment.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/alignment.py?ref_type=heads#L16

```python
horizontal_alignments = (
    "general", "left", "center", "right", "fill", "justify", "centerContinuous",
    "distributed", )
vertical_aligments = (
    "top", "center", "bottom", "justify", "distributed",
)
```

## 4.5 Number Formats

Number formate were covered in the previous section (["Cells and Cell Values"](../3_cells_and_cell_values/README.md#35-number-formatting)).

## 4.6 Conditional Formatting

Conditional formatting involves applying styles based on cell values or formulas.

openpyxl provides support for this, but it's more complex; here in openpyxl documentation (https://openpyxl.pages.heptapod.net/openpyxl/formatting.html).

### 4.6.1 Builtin Formats

### 4.6.2 Standard Conditional Formats

### 4.6.3 Formatting Entire Rows (Range)

Source code reference:

- Class `DifferentialStyle` in `differential.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/differential.py?ref_type=heads#L19
- `rule.py`:
  - class `ColorScaleRule`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/formatting/rule.py?ref_type=heads#L214
  - class `FormulaRule`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/formatting/rule.py?ref_type=heads#L243
  - class `CellIsRule`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/formatting/rule.py?ref_type=heads#L253

## 4.7 Styles and Themes

openpyxl allows working with styles and themes, but the specifics are advanced and are best explored in the library’s documentation. Themes govern the overall look and feel, while styles provide more fine-grained control over individual elements.

## 4.8 Applying Styles to Cells and Ranges

Styles are applied to cells using the appropriate style properties (e.g., cell.font, cell.fill, cell.alignment, cell.number_format, cell.border). To apply styles to ranges, iterate through the cells in the range and apply the styles to each cell individually, or explore using Conditional Formatting which can apply styles to ranges based on conditions. Creating and applying a custom Style object can also be helpful for consistently applying multiple formatting elements.

```python
```

---

Last Updated at: 12/31/2025, 2:16:09 PM  