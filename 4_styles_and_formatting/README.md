# openpyxl - 4. Styles and Formatting

- [openpyxl - 4. Styles and Formatting](#openpyxl---4-styles-and-formatting)
  - [4.1 Styles and Formatting](#41-styles-and-formatting)
    - [4.1.1 Fonts](#411-fonts)
    - [4.1.2 Fill Colors](#412-fill-colors)
    - [4.1.3 Borders](#413-borders)
    - [4.1.4 Alignment](#414-alignment)
  - [4.2 Number Formats](#42-number-formats)
  - [4.3 Conditional Formatting](#43-conditional-formatting)
  - [4.4 Styles and Themes](#44-styles-and-themes)
  - [4.5 Applying Styles to Cells and Ranges](#45-applying-styles-to-cells-and-ranges)

## 4.1 Styles and Formatting

### 4.1.1 Fonts

You can control font attributes like name, size, bold, italic, color, etc., using the `Font` class from `openpyxl.styles`, as below:

```python
```

Source code reference:

- `Font` class: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/fonts.py?ref_type=heads#L32

```python
DEFAULT_FONT = Font(name="Calibri", sz=11, family=2, b=False, i=False,
                    color=Color(theme=1), scheme="minor")
```

Note: the color can be specified using RGB hex codes (e.g., "FF0000" for red) or named colors(?), from documentation (https://openpyxl.pages.heptapod.net/openpyxl/styles.html#colours), you can find `aRGB colours` and `Indexed Colours`. (need test on the named colors)

### 4.1.2 Fill Colors

Cell background colors are controlled with the `PatternFill` class:

```python
```

Source code reference:

- `PatternFill(Fill)` is inherit from class `Fill`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/fills.py?ref_type=heads#L68

### 4.1.3 Borders

### 4.1.4 Alignment

## 4.2 Number Formats

## 4.3 Conditional Formatting

## 4.4 Styles and Themes

## 4.5 Applying Styles to Cells and Ranges

---

Last Updated at: 12/28/2025, 8:16:08 AM 