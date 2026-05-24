from unogenerator.commons import ColorsNamed, Coord, Range, guess_object_style, generate_formula_total_string
from unogenerator import ODS, types
from pydicts import lod, lol
from collections import OrderedDict
from gettext import translation
from logging import debug
import logging
from math import ceil
from importlib.resources import files

logger = logging.getLogger(__name__) # Get logger for this module

try:
    t=translation('unogenerator', files("unogenerator") / 'locale')
    _=t.gettext
except:
    _=str

def row_totals(doc, coord, list_of_totals, color=ColorsNamed.GrayLight, styles=None, row_from="2", row_to=None):
    """
    Generates a row of totals starting from the given coordinate using bulk insertion.

    Args:
        doc (ODS): The ODS document object.
        coord (Coord or str): Coordinate where the totals row will start.
        list_of_totals (list): List of formulas or keys (e.g., ["Total", "#SUM", "#AVG"]).
        color (int, optional): Background color for the cells. Defaults to ColorsNamed.GrayLight.
        styles (list or str, optional): List of styles or a single style. If None, guesses from the adjacent cell.
        row_from (str, optional): The row number where the formula range begins. Defaults to "2".
        row_to (str, optional): The row number where the formula range ends. If None, defaults to the row above `coord`.
    """
    coord=Coord.assertCoord(coord)
    formulas = []
    for letter, total in enumerate(list_of_totals):
        coord_total=coord.addColumnCopy(letter)
        coord_total_from=Coord(coord_total.letter+str(row_from))
        if row_to is None:
            coord_total_to=coord_total.addRowCopy(-1)# row above
        else:
            coord_total_to=Coord(coord_total.letter+str(row_to))
        formulas.append(generate_formula_total_string(total, coord_total_from, coord_total_to))

    if styles is None:
        # Guess style from the first data cell
        first_coord_from = Coord(coord.letter + str(row_from))
        styles = guess_object_style(doc.getValue(first_coord_from), doc.default_cell_style)
    
    doc.addRowWithStyle(coord, formulas, colors=color, styles=styles)


def column_totals(doc, coord, list_of_totals, color=ColorsNamed.GrayLight, styles=None, column_from="B", column_to=None):
    """
    Generates a column of totals starting from the given coordinate using bulk insertion.

    Args:
        doc (ODS): The ODS document object.
        coord (Coord or str): Starting coordinate for the totals column.
        list_of_totals (list): List of formulas or keys (e.g., ["Total", "#SUM", "#AVG", "#MEDIAN"]).
        color (int, optional): Background color for the cells. Defaults to ColorsNamed.GrayLight.
        styles (list or str, optional): List of styles or a single style. If None, guesses from the adjacent cell.
        column_from (str, optional): The column letter where the formula range begins. Defaults to "B".
        column_to (str, optional): The column letter where the formula range ends. If None, defaults to one column before `coord`.
    """
    coord=Coord.assertCoord(coord)
    formulas = []
    for number, total in enumerate(list_of_totals):
        coord_total=coord.addRowCopy(number)
        coord_total_from=Coord(str(column_from) + coord_total.number)
        if column_to is None:
            coord_total_to=coord_total.addColumnCopy(-1)# row above
        else:
            coord_total_to=Coord(str(column_to) + coord_total.number)
        formulas.append(generate_formula_total_string(total, coord_total_from, coord_total_to))

    if styles is None:
        # Guess style from the first data cell
        first_coord_from = Coord(str(column_from) + coord.number)
        styles = guess_object_style(doc.getValue(first_coord_from), doc.default_cell_style)

    doc.addColumnWithStyle(coord, formulas, colors=color, styles=styles)

    
def row_title_values_total( doc, coord, title, values, 
        style_title=None, color_title=ColorsNamed.Orange, 
        style_values=None, color_values=ColorsNamed.White, 
        style_total=None, color_total=ColorsNamed.GrayLight
    ):
    """
        Parameters:
            - values: list: Only one row
    """
    coord=Coord.assertCoord(coord)

    if style_title is None:
        style_title="Bold"

    if style_total is None and len(values)>0:
        style_total=guess_object_style(values[0])


    i=0
    if title is not None:
        doc.addCellWithStyle(coord,title,color_title,style_title)
        i=i+1

    doc.addRowWithStyle(coord.addColumnCopy(i),values,colors=color_values,styles=style_values)
    doc.addCellWithStyle(coord.addColumnCopy(i+len(values)),f"=sum({coord.addColumnCopy(i).string()}:{coord.addColumnCopy(i+len(values)-1).string()})",color_total,style_total)
        

def column_title_values_total( doc, coord, title, values, 
        style_title=None, color_title=ColorsNamed.Orange, 
        style_values=None, color_values=ColorsNamed.White, 
        style_total=None, color_total=ColorsNamed.GrayLight
    ):
    """
        Parameters:
            - values: list: Only one column
    """
    coord=Coord.assertCoord(coord)

    if style_title is None:
        style_title="Bold"

    if style_total is None and len(values)>0:
        style_total=guess_object_style(values[0])


    i=0
    if title is not None:
        doc.addCellWithStyle(coord,title,color_title,style_title)
        i=i+1

    doc.addColumnWithStyle(coord.addRowCopy(i),values,colors=color_values,styles=style_values)
    doc.addCellWithStyle(coord.addRowCopy(i+len(values)),f"=sum({coord.addRowCopy(i).string()}:{coord.addRowCopy(i+len(values)-1).string()})",color_total,style_total)
        

def cross_totals_from_range(
        doc, 
        range_of_data, 
        key="#SUM", 
        column_of_totals=True, 
        row_of_totals=True, 
        vertical_total_title_style="BoldCenter", 
        horizontal_total_title_style="BoldCenter", 
        showing=False,
        label_column="Total",
        label_row="Total",
        skip_columns=0
    ):
    """
    Generates vertical and horizontal totals directly from a data range.

    Args:
        doc (ODS): The ODS document object.
        range_of_data (Range or str): The range containing the data values.
        key (str, optional): Formula key or template. Supported options:
                  1. Predefined aliases: '#SUM', '#AVG', '#MEDIAN'.
                  2. Standard function names: e.g., 'SUM', 'COUNT', 'PRODUCT', 'MAX'.
                  3. Custom templates: e.g., '=SUM({}*1.21)', '={}/100'. The '{}' 
                     placeholder will be replaced by the range string (e.g., 'A1:B10').
                  Defaults to "#SUM".
        column_of_totals (bool, optional): Whether to generate a column of totals to the right. Defaults to True.
        row_of_totals (bool, optional): Whether to generate a row of totals at the bottom. Defaults to True.
        vertical_total_title_style (str, optional): Style for the vertical total title. Defaults to "BoldCenter".
        horizontal_total_title_style (str, optional): Style for the horizontal total title. Defaults to "BoldCenter".
        showing (bool, optional): Legacy parameter. If True, adds an extra 'Sum of totals' block. Defaults to False.
        label_column (str, optional): Label for the column of totals. Defaults to "Total". Set to None to omit.
        label_row (str, optional): Label for the row of totals. Defaults to "Total". Set to None to omit.
        skip_columns (int, optional): Number of columns to skip from the left for row totals. Defaults to 0.

    Returns:
        Range: The new range including the generated totals and labels.
    """
    data_range = Range.assertRange(range_of_data)
    data_rows = data_range.numRows()
    data_columns = data_range.numColumns()
    
    # Guessed style for data (to match totals formatting)
    style_data = guess_object_style(doc.getValue(data_range.c_start), doc.default_cell_style)
    
    final_start_coord = data_range.c_start.copy()
    final_end_coord = data_range.c_end.copy()

    # 1. Add row of totals (Bottom)
    if row_of_totals:
        # Start totals row after skip_columns
        coord_row_totals = data_range.c_start.addRowCopy(data_rows).addColumnCopy(skip_columns)
        num_totals = data_columns - skip_columns
        if num_totals > 0:
            row_totals(doc, coord_row_totals, [key] * num_totals, styles=style_data, row_from=data_range.c_start.number, row_to=data_range.c_end.number)
            final_end_coord.addRow(1)
            
            # Label for row totals
            if label_row:
                if skip_columns > 0:
                    # Place label in the skipped columns area of the totals row
                    coord_label_start = data_range.c_start.addRowCopy(data_rows)
                    coord_label_end = coord_label_start.addColumnCopy(skip_columns - 1)
                    range_label = Range.from_coords(coord_label_start, coord_label_end)
                    doc.addCellMergedWithStyle(range_label, _(label_row), ColorsNamed.GrayLight, horizontal_total_title_style)
                elif data_range.c_start.letterIndex() > 0:
                    # Place label to the left of the range
                    coord_label_row = data_range.c_start.addColumnCopy(-1).addRowCopy(data_rows)
                    doc.addCellWithStyle(coord_label_row, _(label_row), ColorsNamed.GrayLight, horizontal_total_title_style)
                    if final_start_coord.letterIndex() > coord_label_row.letterIndex():
                        final_start_coord = coord_label_row.copy()

    # 2. Add column of totals (Right)
    if column_of_totals:
        coord_col_totals = data_range.c_start.addColumnCopy(data_columns)
        # If we also added a row of totals, we include it in the column totals (cross total)
        total_items = data_rows + (1 if row_of_totals else 0)
        column_totals(doc, coord_col_totals, [key] * total_items, styles=style_data, column_from=data_range.c_start.addColumnCopy(skip_columns).letter, column_to=data_range.c_end.letter)
        final_end_coord.addColumn(1)

        # Label for column totals
        if label_column and data_range.c_start.numberIndex() > 0:
            coord_label_column = data_range.c_start.addRowCopy(-1).addColumnCopy(data_columns)
            doc.addCellWithStyle(coord_label_column, _(label_column), ColorsNamed.GrayLight, vertical_total_title_style)
            if final_start_coord.numberIndex() > coord_label_column.numberIndex():
                final_start_coord = coord_label_column.copy()

    # 3. Handle legacy 'showing' parameter (extra cells)
    if showing:
        if column_of_totals:
            coord_sum_totals = data_range.c_start.addRowCopy(data_rows + 1).addColumnCopy(data_columns)
            doc.addCellWithStyle(coord_sum_totals, generate_formula_total_string(key, data_range.c_start.addColumnCopy(data_columns), final_end_coord), ColorsNamed.GrayLight, style_data)
            if coord_sum_totals.letterIndex() > 0:
                 doc.addCellWithStyle(coord_sum_totals.addColumnCopy(-1), _("Sum of totals"), ColorsNamed.GrayDark, style_data)
            if final_end_coord.numberIndex() < coord_sum_totals.numberIndex():
                final_end_coord = coord_sum_totals.copy()
        elif row_of_totals:
            coord_sum_totals = data_range.c_start.addColumnCopy(data_columns + 1).addRowCopy(data_rows)
            doc.addCellWithStyle(coord_sum_totals, generate_formula_total_string(key, data_range.c_start.addRowCopy(data_rows), final_end_coord), ColorsNamed.GrayLight, style_data)
            if coord_sum_totals.numberIndex() > 0:
                doc.addCellWithStyle(coord_sum_totals.addRowCopy(-1), _("Sum of totals"), ColorsNamed.GrayDark, style_data)
            if final_end_coord.letterIndex() < coord_sum_totals.letterIndex():
                final_end_coord = coord_sum_totals.copy()

    return Range.from_coords(final_start_coord, final_end_coord)


def block_from_lod(doc, coord_start,  lod_, keys=None, columns_header=0,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None, column_of_totals=False, row_of_totals=False, key="#SUM", title=None, word_wrap=True):
    """
    Write cells from a list of ordered dictionaries.

    Args:
        doc (ODS): The ODS document object.
        coord_start (Coord or str): Starting coordinate.
        lod_ (list): List of ordered dictionaries.
        keys (list, optional): List of keys to write. If None, writes all keys. Defaults to None.
        columns_header (int, optional): Number of columns to apply color_column_header. Defaults to 0.
        color_row_header (int, optional): Color for row headers. Defaults to ColorsNamed.Orange.
        color_column_header (int, optional): Color for column headers. Defaults to ColorsNamed.Green.
        color (int, optional): Color for data cells. Defaults to ColorsNamed.White.
        styles (list or str, optional): Styles for data cells. Defaults to None.
        column_of_totals (bool, optional): Whether to generate a column of totals to the right. Defaults to False.
        row_of_totals (bool, optional): Whether to generate a row of totals at the bottom. Defaults to False.
        key (str, optional): Formula key or template. Supported options:
                  1. Predefined aliases: '#SUM', '#AVG', '#MEDIAN'.
                  2. Standard function names: e.g., 'SUM', 'COUNT', 'PRODUCT', 'MAX'.
                  3. Custom templates: e.g., '=SUM({}*1.21)', '={}/100'. The '{}' 
                     placeholder will be replaced by the range string (e.g., 'A1:B10').
                  Defaults to "#SUM".
        title (str, optional): Title for the block. Defaults to None.   
        word_wrap (bool, optional): Enable word wrap and optimal height. Defaults to True.

    Returns:
        Range: The range of the data including headers and totals.
    """
    coord_start=Coord.assertCoord(coord_start)
    c=coord_start.copy()
    
    # 1. Headers
    if keys is None:
        keys=lod.lod_keys(lod_)

    # 2. Title
    if title is not None:
        if len(lod_)==0:
            doc.addCellWithStyle(c, title, ColorsNamed.Red, "BoldCenter")
        else:
            add_of_totals=1 if column_of_totals else 0
            range_title=Range.from_coords(c.copy(), c.addColumnCopy(len(keys)-1+add_of_totals))
            doc.addCellMergedWithStyle(range_title, title, ColorsNamed.Red, "BoldCenter", word_wrap=word_wrap)
        c.addRow(1)

    # 3. Handle empty lod
    if len(lod_)==0:
        doc.addCellWithStyle(c, _("No data to show"), ColorsNamed.White, "BoldCenter")
        return Range.from_coords(coord_start, c)

    # 4. Write column headers
    doc.addRowWithStyle(c, keys, color_row_header, "BoldCenter", word_wrap=word_wrap)
    
    # 5. Generate colors per column
    colors=[]
    for i in range(len(keys)):
        if i < columns_header:
            colors.append(color_column_header)
        else:
            colors.append(color)
   
    # 6. Write data rows
    lol_data=lod.lod2lol(lod_, keys)
    range_data = doc.addListOfRowsWithStyle(c.addRowCopy(1), lol_data, colors, styles, word_wrap=word_wrap)

    # 7. Generate totals
    if (column_of_totals or row_of_totals) and range_data:
        # Default to skipping the first column if columns_header is not specified, 
        # as it's typically a label/ID column in a LOD.
        skip = columns_header if columns_header > 0 else 1 if len(keys) > 1 else 0
        final_range = cross_totals_from_range(doc, range_data, key, column_of_totals, row_of_totals, skip_columns=skip)
        return Range.from_coords(coord_start, final_range.c_end)
    
    return Range.from_coords(coord_start, range_data.c_end if range_data else c)


def block_from_lol(doc, coord_start, lor, headers=None, colors=ColorsNamed.White, styles=None, column_of_totals=False, row_of_totals=False, key="#SUM", title=None, word_wrap=True):
    """
    Writes cells from a list of lists (lor) with optional headers and totals.

    Args:
        doc (ODS): The ODS document object.
        coord_start (Coord or str): Starting coordinate.
        lor (list): List of lists (data rows).
        headers (list, optional): List of header strings. Defaults to None.
        colors (list or int, optional): Column colors. Defaults to ColorsNamed.White.
        styles (list or str, optional): Column styles. Defaults to None.
        column_of_totals (bool, optional): Whether to generate a column of totals to the right. Defaults to False.
        row_of_totals (bool, optional): Whether to generate a row of totals at the bottom. Defaults to False.
        key (str, optional): Formula key or template. Defaults to "#SUM".
        title (str, optional): Main title for the block. Defaults to None.
        word_wrap (bool, optional): Enable word wrap and optimal height. Defaults to True.

    Returns:
        Range: The range of the data including headers and totals.
    """
    coord_start = Coord.assertCoord(coord_start)
    c = coord_start.copy()

    # 1. Title
    if title is not None:
        if not lor and not headers:
            doc.addCellWithStyle(c, title, ColorsNamed.Red, "BoldCenter")
        else:
            num_cols = len(headers) if headers else len(lor[0]) if lor else 1
            add_of_totals = 1 if column_of_totals else 0
            range_title = Range.from_coords(c.copy(), c.addColumnCopy(num_cols - 1 + add_of_totals))
            doc.addCellMergedWithStyle(range_title, title, ColorsNamed.Red, "BoldCenter", word_wrap=word_wrap)
        c.addRow(1)

    # 2. Handle empty data
    if not lor and not headers:
        doc.addCellWithStyle(c, _("No data to show"), ColorsNamed.White, "BoldCenter")
        return Range.from_coords(coord_start, c)

    # 3. Write column headers
    if headers:
        doc.addRowWithStyle(c, headers, ColorsNamed.Orange, "BoldCenter", word_wrap=word_wrap)
        c.addRow(1)

    # 4. Write data rows
    range_data = doc.addListOfRowsWithStyle(c, lor, colors, styles, word_wrap=word_wrap)

    # 5. Generate totals
    if (column_of_totals or row_of_totals) and range_data:
        # Default to skipping the first column if it looks like a label column
        skip = 1 if (headers and len(headers) > 1) or (lor and len(lor[0]) > 1) else 0
        final_range = cross_totals_from_range(doc, range_data, key, column_of_totals, row_of_totals, skip_columns=skip)
        return Range.from_coords(coord_start, final_range.c_end)

    return Range.from_coords(coord_start, range_data.c_end if range_data else c)


def sheet_stylenames(doc):
    """
    Creates a new sheet called "Internal style names" listing all available styles 
    organized into four columns: CellStyles, PageStyles, GraphicStyles, and TableStyles.
    Useful for identifying available style names for use with other methods.

    Args:
        doc (ODS): The ODS document object.
    """
    cell_styles = doc.dict_stylenames.get("CellStyles", [])
    page_styles = doc.dict_stylenames.get("PageStyles", [])
    graphic_styles = doc.dict_stylenames.get("GraphicStyles", [])
    table_styles = doc.dict_stylenames.get("TableStyles", [])
    
    max_len = max(len(cell_styles), len(page_styles), len(graphic_styles), len(table_styles))
    
    lod_ = []
    for i in range(max_len):
        lod_.append(OrderedDict({
            "CellStyles": cell_styles[i] if i < len(cell_styles) else "",
            "PageStyles": page_styles[i] if i < len(page_styles) else "",
            "GraphicStyles": graphic_styles[i] if i < len(graphic_styles) else "",
            "TableStyles": table_styles[i] if i < len(table_styles) else ""
        }))
    
    sheet_from_lod(doc, "Internal style names", lod_, freezeandselect="A2", columns_width_mode=types.ColumnsWidthMode.FROM_LOD)

def sheet_from_lol(doc, sheetname, lor, headers, column_of_totals=False, row_of_totals=False, freezeandselect=None, titulo=None, word_wrap=True, **kwargs_columnswidth):
    """
    Creates a sheet from a list of lists (lor) with headers and optional totals.

    Args:
        doc (ODS): The ODS document object.
        sheetname (str): The name for the new sheet.
        lor (list): The list of lists (data rows).
        headers (list): The list of header strings.
        column_of_totals (bool, optional): Whether to generate a column of totals to the right. Defaults to False.
        row_of_totals (bool, optional): Whether to generate a row of totals at the bottom. Defaults to False.
        freezeandselect (str, optional): Coordinate to freeze panes at. Defaults to None.
        titulo (str, optional): An optional title to merge across the top of the sheet. Defaults to None.
        word_wrap (bool, optional): Enable word wrap and optimal height. Defaults to True.
        **kwargs_columnswidth: Keyword arguments for setColumnsWidth.
    """
    columns_width_mode = kwargs_columnswidth.get("columns_width_mode", types.ColumnsWidthMode.FROM_LOL)
    char_to_cm = kwargs_columnswidth.get("char_to_cm", 0.22)
    padding_cm = kwargs_columnswidth.get("padding_cm", 0.5)
    min_width_cm = kwargs_columnswidth.get("min_width_cm", 2.0)
    max_width_cm = kwargs_columnswidth.get("max_width_cm", 15.0)
    value = kwargs_columnswidth.get("value")

    doc.createSheet(sheetname)

    range_block = block_from_lol(
        doc, "A1", lor, headers, 
        column_of_totals=column_of_totals, 
        row_of_totals=row_of_totals, 
        title=titulo, 
        word_wrap=word_wrap
    )

    if value is None:
        if columns_width_mode == types.ColumnsWidthMode.FROM_SHEET_CELLS:
            value = doc
        else:
            value = [headers] + lor if headers else lor

    doc.setColumnsWidth(value, columns_width_mode, char_to_cm, padding_cm, min_width_cm, max_width_cm)

    if freezeandselect:
        doc.freezeAndSelect(freezeandselect, freezeandselect, freezeandselect)
    
    return range_block

def sheet_split_with_big_lol(doc, sheet_name, lor, headers, headers_colors=ColorsNamed.Orange, coord_to_freeze="A2",  max_rows=1048575, word_wrap=True):
    """
    Splits a large list of rows across multiple sheets if it exceeds LibreOffice Calc's row limits.

    If the number of rows is lower than `max_rows`, it generates a single normal sheet.
    A one-line header is added to each generated sheet.

    Args:
        doc (ODS): The ODS document object.
        sheet_name (str): The root name of the generated sheet(s).
        lor (list): List of rows containing the data.
        headers (list): List of strings representing the column headers.
        headers_colors (int, optional): Color for the sheet headers. Defaults to ColorsNamed.Orange.
        columns_width (int, list, or None, optional): Defines column widths. If None, sets automatically.
            If int, applies to all columns. If list, specifies width per column. Defaults to None.
        coord_to_freeze (Coord or str, optional): Coordinate to freeze panes at. Defaults to "A2".
        max_rows (int, optional): Maximum number of rows per sheet (Calc limit is 1,048,576). Defaults to 1048575.
        word_wrap (bool, optional): Enable word wrap and optimal height. Defaults to True.
    """
    ceil_=ceil(len(lor)/max_rows)
    for num_sheet in range(ceil_):
        #Sets name and headers
        if ceil_>1:
            name=_("{0} ({1} of {2})").format(sheet_name, num_sheet+1,  ceil_)
            logger.debug(_("More than {0} rows. Spliting {1} of {2} sheets").format(max_rows, num_sheet+1,  ceil_))
        else:
            name=sheet_name
        doc.createSheet(name)
        doc.addRowWithStyle("A1", headers, headers_colors, "BoldCenter", word_wrap=word_wrap)
        
        #Splits data
        from_=max_rows*num_sheet
        to_=max_rows*(num_sheet+1) if len(lor)>=max_rows*(num_sheet+1) else len(lor)
        doc.addListOfRowsWithStyle("A2", lor[from_:to_], word_wrap=word_wrap)

        #Sets width of columns
    
        doc.setColumnsWidth(lor[from_:to_], types.ColumnsWidthMode.FROM_LOL)

        doc.freezeAndSelect(Coord.assertCoord(coord_to_freeze))
    

def sheet_from_lod(doc, sheetname, lod_,  column_of_totals=False, row_of_totals=False, freezeandselect=None, title=None, word_wrap=True, styles=None, **kwargs_columnswidth):
    """
        kwargs son los parametros de la funcion setColumnsWidth
    """
    columns_width_mode=kwargs_columnswidth.get("columns_width_mode", types.ColumnsWidthMode.FROM_LOD)
    char_to_cm=kwargs_columnswidth.get("char_to_cm", 0.22)
    padding_cm=kwargs_columnswidth.get("padding_cm", 0.5)
    min_width_cm=kwargs_columnswidth.get("min_width_cm", 2.0)
    max_width_cm=kwargs_columnswidth.get("max_width_cm", 15.0)
    value = kwargs_columnswidth.get("value")

    doc.createSheet(sheetname)
              
    range_final=block_from_lod(doc, "A1", lod_, column_of_totals=column_of_totals, row_of_totals=row_of_totals, word_wrap=word_wrap, styles=styles, title=title )
    
    if value is None:
        if columns_width_mode == types.ColumnsWidthMode.FROM_SHEET_CELLS:
            value = doc
        else:
            value = lod_
        
    doc.setColumnsWidth(value, columns_width_mode, char_to_cm, padding_cm, min_width_cm, max_width_cm)
    if freezeandselect:
        doc.freezeAndSelect(freezeandselect,freezeandselect, freezeandselect)
    return range_final




def block_from_lod_with_headers(doc, lod_, coord, subtitles=[], titulo=None, column_of_totals=False, row_of_totals=False, freezeandselect=None, key="#SUM", word_wrap=True):
    """
    Writes data from a list of ordered dictionaries with custom header groups, and optional totals.

    Args:
        doc (ODS): The ODS document object.
        lod_ (list): List of ordered dictionaries containing the data.
        coord (Coord or str): Starting coordinate.
        subtitles (list): List of lists [title, first_key] defining header groups.
        titulo (str, optional): Main title for the entire block. Defaults to None.
        column_of_totals (bool, optional): Whether to generate column totals. Defaults to False.
        row_of_totals (bool, optional): Whether to generate row totals. Defaults to False.
        freezeandselect (str or Coord, optional): Coordinate to freeze and select. Defaults to None.
        key (str, optional): Formula key for totals (e.g., "#SUM"). Defaults to "#SUM".
        word_wrap (bool, optional): Enable word wrap and optimal height. Defaults to True.

    Returns:
        Range: The data range (excluding headers).
    """
    if len(lod_)==0:
        doc.addCell(coord, "Sin datos que consigar")
        return

    coord=Coord.assertCoord(coord)
    coord=Coord(coord.string())# To avoid carry internal coord movements
    
    keys=lod.lod_keys(lod_)
         

    #Añado en la lista un campo nuevo de indice de inicio y indice de final
    for i in range(len(subtitles)):
        subtitles[i].append(keys.index(subtitles[i][1])) #Añade el indice de inicio
        if i== len(subtitles)-1:#Ultimo titulo
            subtitles[i].append(len(keys)-1)
        else:# hay mas titulos
            subtitles[i].append(keys.index(subtitles[i+1][1])-1)
    
    # Crea titulo principal
    if titulo is not None:
        if column_of_totals:
            add_of_totals=1
        else:
            add_of_totals=0
        doc.addCellMergedWithStyle(Range.from_coords(coord,coord.addColumnCopy(len(keys)-1+add_of_totals)), titulo, ColorsNamed.Red, "BoldCenter", word_wrap=word_wrap)
        coord.addRow(1)
     
         
    # Crea titulos
    for title, key_start, index_start, index_end in subtitles:
        c_start=coord.addColumnCopy(index_start)
        c_fin=coord.addColumnCopy(index_end)
        doc.addCellMergedWithStyle(Range.from_coords(c_start,c_fin), title, ColorsNamed.Orange, "BoldCenter", word_wrap=word_wrap)


    #Imprime listas de diccionarios
    range_=block_from_lod(doc, coord.addRowCopy(1),lod_,color_row_header=ColorsNamed.Yellow, word_wrap=word_wrap)

    if column_of_totals or row_of_totals:
        cross_totals_from_range(doc, range_, key, column_of_totals, row_of_totals)
        
    if freezeandselect:
        doc.freezeAndSelect(freezeandselect, freezeandselect, freezeandselect)

    return range_
