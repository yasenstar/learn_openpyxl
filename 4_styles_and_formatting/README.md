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
- `fill_type`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/fills.py?ref_type=heads#L42

### 4.1.3 Borders

Borders are defined using the `Border`, `Side` classes:

```python
```

Soruce code reference:

- class `Border` in `borders.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/borders.py?ref_type=heads#L54
  - `border_style = Alias('style')`
- class `Side` in `borders.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/borders.py?ref_type=heads#L33
  - `Side` `style=NoneSet()`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/borders.py?ref_type=heads#L41

### 4.1.4 Alignment

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

## 4.2 Number Formats

Number formate were covered in the previous section (["Cells and Cell Values"](../3_cells_and_cell_values/README.md#35-number-formatting)).

Source code reference:

- Class `DifferentialStyle` in `differential.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/styles/differential.py?ref_type=heads#L19
- `rule.py`:
  - class `ColorScaleRule`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/formatting/rule.py?ref_type=heads#L214
  - class `FormulaRule`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/formatting/rule.py?ref_type=heads#L243
  - class `CellIsRule`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/formatting/rule.py?ref_type=heads#L253

## 4.3 Conditional Formatting

Conditional formatting involves applying styles based on cell values or formulas.

openpyxl provides support for this, but it's more complex; here in openpyxl documentation (https://openpyxl.pages.heptapod.net/openpyxl/formatting.html).

## 4.4 Styles and Themes

openpyxl allows working with styles and themes, but the specifics are advanced and are best explored in the library’s documentation. Themes govern the overall look and feel, while styles provide more fine-grained control over individual elements.

## 4.5 Applying Styles to Cells and Ranges

Styles are applied to cells using the appropriate style properties (e.g., cell.font, cell.fill, cell.alignment, cell.number_format, cell.border). To apply styles to ranges, iterate through the cells in the range and apply the styles to each cell individually, or explore using Conditional Formatting which can apply styles to ranges based on conditions. Creating and applying a custom Style object can also be helpful for consistently applying multiple formatting elements.

```python
```

---

Last Updated at: 12/28/2025, 8:16:08 AM 