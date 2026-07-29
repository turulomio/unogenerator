# UnoGenerator Helpers Specification

Detailed reference of all helper functions in `unogenerator.helpers`.

## Table of Contents
1. [Totals Generation (Basic)](#1-totals-generation-basic)
2. [Totals Generation (Titled)](#2-totals-generation-titled)
3. [Advanced Totals (Cross-Calculations)](#3-advanced-totals-cross-calculations)
4. [Data Block Helpers](#4-data-block-helpers)
5. [Complete Sheet Helpers](#5-complete-sheet-helpers)
6. [Utility Helpers](#6-utility-helpers)

---

## 1. Totals Generation (Basic)

### `row_totals`
Generates a horizontal row of formulas.
- `doc` (ODS): The ODS document object.
- `coord` (Coord or str): Starting coordinate.
- `list_of_totals` (list): List of formula keys (e.g. `["#SUM", "#AVG"]`).
- `color` (int, default=ColorsNamed.GrayLight): Background color.
- `styles` (list or str, default=None): Cell style(s).
- `row_from` (str, default="2"): Start row for calculation range.
- `row_to` (str, default=None): End row for calculation range.

### `column_totals`
Generates a vertical column of formulas.
- `doc` (ODS): The ODS document object.
- `coord` (Coord or str): Starting coordinate.
- `list_of_totals` (list): List of formula keys.
- `color` (int, default=ColorsNamed.GrayLight): Background color.
- `styles` (list or str, default=None): Cell style(s).
- `column_from` (str, default="B"): Start column for calculation range.
- `column_to` (str, default=None): End column for calculation range.

---

## 2. Totals Generation (Titled)

### `row_title_values_total`
Creates a row with a title, a sequence of values, and their sum.
- `doc` (ODS): The ODS document object.
- `coord` (Coord or str): Starting coordinate.
- `title` (str): Title for the row.
- `values` (list): List of numerical values.
- `style_title` (str, default="Bold"): Style for the title.
- `color_title` (int, default=ColorsNamed.Orange): Title color.
- `style_values` (str, default=None): Style for the values.
- `color_values` (int, default=ColorsNamed.White): Values color.
- `style_total` (str, default=None): Style for the total.
- `color_total` (int, default=ColorsNamed.GrayLight): Total color.

### `column_title_values_total`
Creates a column with a title, a sequence of values, and their sum.
- `doc` (ODS): The ODS document object.
- `coord` (Coord or str): Starting coordinate.
- `title` (str): Title for the column.
- `values` (list): List of numerical values.
- `style_title` (str, default="Bold"): Style for the title.
- `color_title` (int, default=ColorsNamed.Orange): Title color.
- `style_values` (str, default=None): Style for the values.
- `color_values` (int, default=ColorsNamed.White): Values color.
- `style_total` (str, default=None): Style for the total.
- `color_total` (int, default=ColorsNamed.GrayLight): Total color.

---

## 3. Advanced Totals (Cross-Calculations)

### `cross_totals_from_range`
Generates optimized vertical and horizontal totals for a data range.
- `doc` (ODS): The ODS document object.
- `range_of_data` (Range or str): The data range to calculate from.
- `key` (str, default="#SUM"): Formula alias, function name, or template (e.g. `"=SUM({}*1.21)"`).
- `column_of_totals` (bool, default=True): If True, adds column at the right.
- `row_of_totals` (bool, default=True): If True, adds row at the bottom.
- `vertical_total_title_style` (str, default="BoldCenter"): Style for right label.
- `horizontal_total_title_style` (str, default="BoldCenter"): Style for bottom label.
- `showing` (bool, default=False): Legacy. If True, adds extra "Sum of totals" block.
- `label_column` (str, default="Total"): Text for the vertical total label.
- `label_row` (str, default="Total"): Text for the horizontal total label.
- `skip_columns` (int, default=0): Number of columns to skip for bottom row totals.

---

## 4. Data Block Helpers

### `block_from_lod`
Inserts data from a List of Ordered Dictionaries.
- `doc` (ODS): The ODS document object.
- `coord_start` (Coord or str): Starting coordinate.
- `lod_` (list): The list of dictionaries.
- `keys` (list, default=None): Specific keys to write.
- `columns_header` (int, default=0): Number of columns to color as headers.
- `color_row_header` (int, default=ColorsNamed.Orange): Color for the keys row.
- `color_column_header` (int, default=ColorsNamed.Green): Color for the header columns.
- `color` (int, default=ColorsNamed.White): Background for data cells.
- `styles` (list or str, default=None): Styles for data cells.
- `column_of_totals` (bool, default=False): Generate totals on the right.
- `row_of_totals` (bool, default=False): Generate totals at the bottom.
- `key` (str, default="#SUM"): Formula key (see `cross_totals_from_range`).
- `title` (str, default=None): Merged title for the block.
- `word_wrap` (bool, default=True): Enable text wrapping.

### `block_from_lol`
Inserts data from a List of Lists.
- `doc` (ODS): The ODS document object.
- `coord_start` (Coord or str): Starting coordinate.
- `lor` (list): The list of lists.
- `headers` (list, default=None): Column header names.
- `colors` (list or int, default=ColorsNamed.White): Column colors.
- `styles` (list or str, default=None): Column styles.
- `column_of_totals` (bool, default=False): Generate totals on the right.
- `row_of_totals` (bool, default=False): Generate totals at the bottom.
- `key` (str, default="#SUM"): Formula key.
- `title` (str, default=None): Merged title for the block.
- `word_wrap` (bool, default=True): Enable text wrapping.

### `block_from_lod_with_headers`
LOD writer with hierarchical sub-headers.
- `doc` (ODS): The ODS document object.
- `lod_` (list): List of dictionaries.
- `coord` (Coord or str): Starting coordinate.
- `subtitles` (list, default=[]): Groups of columns. List of `[title, first_key]`.
- `titulo` (str, default=None): Main title.
- `column_of_totals` (bool, default=False): Generate totals on the right.
- `row_of_totals` (bool, default=False): Generate totals at the bottom.
- `freezeandselect` (Coord or str, default=None): Auto-freeze coordinate.
- `key` (str, default="#SUM"): Formula key.
### `photos_from_lod_ods`
Creates a photo catalog table in an ODS spreadsheet from a List of Dictionaries. Binary image blobs (`bytes` or `bytearray`) are automatically detected and anchored to cells, with row heights and column widths auto-adjusted to fit the image dimensions.
- `doc` (ODS): The ODS document object.
- `coord_start` (Coord or str): Starting coordinate.
- `lod_photos` (list): List of dictionaries containing data and photo blobs.
- `headers` (list, default=None): Optional list of header labels. If `None` (default), no header row is written.
- `keys` (list, default=None): Specific dictionary keys to include/order. If `None`, auto-detects all keys.
- `default_width` (float, default=2.5): Default image width in cm.
- `default_height` (float, default=2.5): Default image height in cm.
- `title` (str, default=None): Optional merged title for the block.
- `color_row_header` (int, default=ColorsNamed.Orange): Color for header row.
- `styles` (list or str, default=None): Style(s) for data cells.
- `word_wrap` (bool, default=True): Enable text wrapping.

---


## 5. Complete Sheet Helpers

### `sheet_from_lod`
Creates a new sheet and populates it from an LOD.
- `doc` (ODS): The ODS document object.
- `sheetname` (str): Name of the new sheet.
- `lod_` (list): List of dictionaries.
- `column_of_totals` (bool, default=False): Right totals.
- `row_of_totals` (bool, default=False): Bottom totals.
- `freezeandselect` (str, default=None): Coordination to freeze.
- `title` (str, default=None): Main title.
- `word_wrap` (bool, default=True): Text wrap.
- `styles` (list or str, default=None): Data styles.
- `**kwargs_columnswidth`: Extra params for column width calculation.

### `sheet_from_lol`
Creates a new sheet and populates it from an LOL.
- `doc` (ODS): The ODS document object.
- `sheetname` (str): Name of the new sheet.
- `lor` (list): List of lists.
- `headers` (list): Header names.
- `column_of_totals` (bool, default=False): Right totals.
- `row_of_totals` (bool, default=False): Bottom totals.
- `freezeandselect` (str, default=None): Coordination to freeze.
- `titulo` (str, default=None): Main title.
- `word_wrap` (bool, default=True): Text wrap.
- `**kwargs_columnswidth`: Extra params for column width calculation.

### `sheet_split_with_big_lol`
Creates multiple sheets for massive datasets.
- `doc` (ODS): The ODS document object.
- `sheet_name` (str): Base name for sheets.
- `lor` (list): Massive list of lists.
- `headers` (list): Header names.
- `headers_colors` (int, default=ColorsNamed.Orange): Header background.
- `coord_to_freeze` (Coord or str, default="A2"): Freeze coordinate.
- `max_rows` (int, default=1048575): Max rows per sheet.
- `word_wrap` (bool, default=True): Text wrap.

---

## 6. Utility Helpers

### `sheet_stylenames`
Generates a reference sheet with available document styles.
- `doc` (ODS): The ODS document object.
