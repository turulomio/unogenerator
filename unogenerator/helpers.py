from unogenerator.commons import ColorsNamed, Coord, Range, guess_object_style, generate_formula_total_string
from unogenerator import ODS, types
from pydicts import lod
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
    Generates a row of totals starting from the given coordinate.

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
    for letter, total in enumerate(list_of_totals):
        coord_total=coord.addColumnCopy(letter)
        coord_total_from=Coord(coord_total.letter+row_from)
        if row_to is None:
            coord_total_to=coord_total.addRowCopy(-1)# row above
        else:
            coord_total_to=Coord(coord_total.letter+row_to)

        if styles is None:
            style=guess_object_style(doc.getValue(coord_total_from))
        elif styles.__class__.__name__ != "list":
            style=styles
        else:
            style=styles[letter]

        doc.addCellWithStyle(coord_total, generate_formula_total_string(total, coord_total_from, coord_total_to), color, style)


def column_totals(doc, coord, list_of_totals, color=ColorsNamed.GrayLight, styles=None, column_from="B", column_to=None):
    """
    Generates a column of totals starting from the given coordinate.

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
    for number, total in enumerate(list_of_totals):
        coord_total=coord.addRowCopy(number)
        coord_total_from=Coord(column_from + coord_total.number)
        if column_to is None:
            coord_total_to=coord_total.addColumnCopy(-1)# row above
        else:
            coord_total_to=Coord(column_to + coord_total.number)

        if styles is None:
            style=guess_object_style(doc.getValue(coord_total_from))
        elif styles.__class__.__name__ != "list":
            style=styles
        else:
            style=styles[number]

        doc.addCellWithStyle(coord_total, generate_formula_total_string(total, coord_total_from, coord_total_to), color, style)
        
def row_title_values_total( doc, coord, title, values, 
        style_title=None, color_title=ColorsNamed.Orange, 
        style_values=None, color_values=ColorsNamed.White, 
        style_total=None, color_total=ColorsNamed.GrayLight
    ):
    """
    Creates a column containing a title, a list of values, and a total sum at the bottom.

    Args:
        doc (ODS): The ODS document object.
        coord (Coord or str): Starting coordinate.
        title (str): Title to be placed at the starting coordinate.
        values (list): List of values to be placed below the title.
        style_title (str, optional): Style for the title cell. Defaults to "BoldCenter".
        color_title (int, optional): Background color for the title. Defaults to ColorsNamed.Orange.
        style_values (list or str, optional): Styles for the value cells. Defaults to None.
        color_values (list or int, optional): Colors for the value cells. Defaults to ColorsNamed.White.
        style_total (str, optional): Style for the total cell. Defaults to None.
        color_total (int, optional): Background color for the total cell. Defaults to ColorsNamed.GrayLight.
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
    doc.addCellWithStyle(coord.addColumnCopy(i+len(values)),f"=sum({coord.addColumnCopy(i).string()}:{coord.addColumnCopy(i+len(values)-1).string()}",color_total,style_total)

def column_title_values_total(doc, coord, title, values,
        style_title=None, color_title=ColorsNamed.Orange, 
        style_values=None, color_values=ColorsNamed.White, 
        style_total=None, color_total=ColorsNamed.GrayLight
    ):
    """
    Creates a row containing a title, a list of values, and a total sum at the end.

    Args:
        doc (ODS): The ODS document object.
        coord (Coord or str): Starting coordinate.
        title (str): Title to be placed at the starting coordinate.
        values (list): List of values to be placed after the title.
        style_title (str, optional): Style for the title cell. Defaults to "Bold".
        color_title (int, optional): Background color for the title. Defaults to ColorsNamed.Orange.
        style_values (list or str, optional): Styles for the value cells. Defaults to None.
        color_values (list or int, optional): Colors for the value cells. Defaults to ColorsNamed.White.
        style_total (str, optional): Style for the total cell. Defaults to None.
        color_total (int, optional): Background color for the total cell. Defaults to ColorsNamed.GrayLight.
    """
    coord=Coord.assertCoord(coord)

    if style_title is None:
        style_title="BoldCenter"

    if style_total is None and len(values)>0:
        style_total=guess_object_style(values[0])


    i=0
    if title is not None:
        doc.addCellWithStyle(coord,title,color_title,style_title)
        i=i+1

    doc.addColumnWithStyle(coord.addRowCopy(i),values,colors=color_values,styles=style_values)
    doc.addCellWithStyle(coord.addRowCopy(i+len(values)),f"=sum({coord.addRowCopy(i).string()}:{coord.addRowCopy(i+len(values)-1).string()}",color_total,style_total)
        

def cross_totals_from_range (
        doc, 
        range_of_data, 
        key="#SUM", 
        column_of_totals=True, 
        row_of_totals=True, 
        vertical_total_title_style="BoldCenter", 
        horizontal_total_title_style="BoldCenter", 
        showing=False
    ):
    """
    Generates vertical and horizontal totals directly from a data range.

    Calculates sums (or specified formulas) for the given range and adds "Total" labels.

    Args:
        doc (ODS): The ODS document object.
        range_of_data (Range or str): The range containing the data values.
        key (str, optional): Formula key to apply (e.g., "#SUM"). Defaults to "#SUM".
        column_of_totals (bool, optional): Whether to generate column totals. Defaults to True.
        row_of_totals (bool, optional): Whether to generate row totals. Defaults to True.
        vertical_total_title_style (str, optional): Style for the vertical total title. Defaults to "BoldCenter".
        horizontal_total_title_style (str, optional): Style for the horizontal total title. Defaults to "BoldCenter".
        showing (bool, optional): If True, shows a 'Sum of totals' cell when either column_of_totals or row_of_totals is True. Defaults to False.

    Returns:
        Range: The original data range.
    """
    range=Range.assertRange(range_of_data)
    data_rows=range.numRows()
    data_columns=range.numColumns()
    coord_horizontal_title=range.c_start.addColumnCopy(-1).addRowCopy(data_rows) 
    coord_vertical_title=range.c_start.addRowCopy(-1).addColumnCopy(data_columns)
    style_data=guess_object_style(doc.getValue(range.c_end))
    
    if column_of_totals==True and row_of_totals==True:
        doc.addCellWithStyle(coord_horizontal_title, _("Total"), ColorsNamed.GrayLight, horizontal_total_title_style)
        row_totals(doc, coord_horizontal_title.addColumnCopy(1), [key]*data_columns,styles=style_data, row_from=range.c_start.number)
        doc.addCellWithStyle(coord_vertical_title, _("Total"), ColorsNamed.GrayLight, vertical_total_title_style)
        column_totals(doc, coord_vertical_title.addRowCopy(1),[key]*(data_rows+1), styles=style_data, column_from=range.c_start.letter)
    elif column_of_totals==True:
        doc.addCellWithStyle(coord_vertical_title, _("Total"), ColorsNamed.GrayLight, vertical_total_title_style)
        column_totals(doc, coord_vertical_title.addRowCopy(1),[key]*(data_rows+0), styles=style_data, column_from=range.c_start.letter)
        if showing is True:
            coord_sum_totals=coord_vertical_title.addRowCopy(data_rows+1)
            doc.addCellWithStyle(coord_sum_totals, generate_formula_total_string(key, range.c_start.addColumnCopy(data_columns+1), range.c_end.addColumnCopy(1)), ColorsNamed.GrayLight, style_data)
            doc.addCellWithStyle(coord_sum_totals.addColumnCopy(-1), _("Sum of totals"), ColorsNamed.GrayDark, style_data)
    elif row_of_totals==True:
        doc.addCellWithStyle(coord_horizontal_title, _("Total"), ColorsNamed.GrayLight, horizontal_total_title_style)
        row_totals(doc, coord_horizontal_title.addColumnCopy(1),[key]*(data_columns+0), styles=style_data, row_from=range.c_start.number) #1 menos por la esquina
        if showing is True:
            coord_sum_totals=coord_horizontal_title.addColumnCopy(data_columns+1)
            doc.addCellWithStyle(coord_sum_totals, generate_formula_total_string(key, range.c_start.addRowCopy(data_rows+1), range.c_end.addRowCopy(1)), ColorsNamed.GrayLight, style_data)
            doc.addCellWithStyle(coord_sum_totals.addRowCopy(-1), _("Sum of totals"), ColorsNamed.GrayDark, style_data)

    return range_of_data


def block_from_lod(doc, coord_start,  lod_, keys=None, columns_header=0,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None, column_of_totals=False, row_of_totals=False, key="#SUM", title=None, word_wrap=False):
    """
    Write cells from a list of ordered dictionaries.
    Params:
        doc (ODS): The ODS document object.
        coord_start (Coord or str): Starting coordinate.
        lod_ (list): List of ordered dictionaries.
        keys (list, optional): List of keys to write. If None, writes all keys. Defaults to None.
        columns_header (int, optional): Number of columns to apply color_header. Defaults to 0.
        color_row_header (int, optional): Color for row headers. Defaults to ColorsNamed.Orange.
        color_column_header (int, optional): Color for column headers. Defaults to ColorsNamed.Green.
        color (int, optional): Color for data cells. Defaults to ColorsNamed.White.
        styles (list or str, optional): Styles for data cells. Defaults to None.
        column_of_totals: Add a total at the bottom of the block
        row_of_totals: Add a total at the right of the block
        key (str, optional): Formula key for totals. Defaults to "#SUM".
        title (str, optional): Title for the block. Defaults to None.   
        word_wrap (bool, optional): Enable word wrap and optimal height. Defaults to False.
    Returns:
        Range: The range of the data without headers. Useful to set totals.
    """
    # Check is a Coord object and makes a copy to avoid internal coord movements
    coord_start=Coord.assertCoord(coord_start)
    c=coord_start.copy()
    
    #Prepara el titulo
    if title is not None:
        if len(lod_)==0:
            doc.addCellWithStyle(c, title, ColorsNamed.Red, "BoldCenter")
        else:
            add_of_totals=1 if column_of_totals else 0
            if keys is None:
                range_title=Range.from_coords("A1", Coord("A1").addColumn(len(lod_[0].keys())-1+add_of_totals))
            else:
                range_title=Range.from_coords("A1", Coord("A1").addColumn(len(keys)-1)+add_of_totals)
            doc.addCellMergedWithStyle(range_title, title, ColorsNamed.Red, "BoldCenter", word_wrap=word_wrap)
        c.addRow(1)

    # Empty lod
    if len(lod_)==0:
        doc.addCellWithStyle(c, _("No data to show"), ColorsNamed.White, "BoldCenter")
        return None

        
    #Headers
    if keys is None:
        keys=lod.lod_keys(lod_)
    doc.addRowWithStyle(c, keys, color_row_header, "BoldCenter", word_wrap=word_wrap)
    
    #Generate list of colors
    colors=[]
    for i in range(len(keys)):
        if i <= columns_header-1:
            colors.append(color_column_header)
        else:
            colors.append(color)
   
    #Generate list of rows
    lol_=lod.lod2lol(lod_, keys)
    range_block= doc.addListOfRowsWithStyle(c.addRow(), lol_, colors, styles, word_wrap=word_wrap)

    # Generate totals 
    if column_of_totals or row_of_totals:
        if column_of_totals:
            if columns_header==0:
                columns_header=1
            range_block.c_start.addColumn(columns_header) ## Adds to skip columns headers
        return cross_totals_from_range (doc, range_block, key, column_of_totals, row_of_totals)
    return range_block

def sheet_stylenames(doc):
    """
    Creates a new sheet called "Internal style names" listing all ODS styles in four columns:
    CellStyles, PageStyles, GraphicStyles, and TableStyles.

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

def sheet_from_lol(doc, sheetname, lor, headers, column_of_totals=False, row_of_totals=False, freezeandselect=None, titulo=None, word_wrap=False, **kwargs_columnswidth):
    """
    Creates a sheet from a list of lists (lol) with headers and optional totals.

    Args:
        doc (ODS): The ODS document object.
        sheetname (str): The name for the new sheet.
        lor (list): The list of lists (data rows).
        headers (list): The list of header strings.
        column_of_totals (bool, optional): Whether to generate column totals. Defaults to False.
        row_of_totals (bool, optional): Whether to generate row totals. Defaults to False.
        freezeandselect (str, optional): Coordinate to freeze panes at. Defaults to None.
        titulo (str, optional): An optional title to merge across the top of the sheet. Defaults to None.
        word_wrap (bool, optional): Enable word wrap and optimal height. Defaults to False.
        **kwargs_columnswidth: Keyword arguments for setColumnsWidth.
    """
    columns_width_mode = kwargs_columnswidth.get("columns_width_mode", types.ColumnsWidthMode.FROM_LOL)
    char_to_cm = kwargs_columnswidth.get("char_to_cm", 0.22)
    padding_cm = kwargs_columnswidth.get("padding_cm", 0.5)
    min_width_cm = kwargs_columnswidth.get("min_width_cm", 2.0)
    max_width_cm = kwargs_columnswidth.get("max_width_cm", 15.0)

    doc.createSheet(sheetname)

    if not lor and not headers:
        if titulo:
            doc.addCellMergedWithStyle("A1:D1", titulo, ColorsNamed.Red, "BoldCenter")
        else:
            doc.addCellMergedWithStyle("A1:D1", "No hay datos", ColorsNamed.Red, "BoldCenter")
        return

    c_start = Coord("A1")

    if titulo:
        c_end = c_start.addColumnCopy(len(headers) - 1)
        range_titulo = Range.from_coords(c_start, c_end)
        doc.addCellMergedWithStyle(range_titulo, titulo, ColorsNamed.Orange, "BoldCenter", word_wrap=word_wrap)
        c_start.addRow(1)

    doc.addRowWithStyle(c_start, headers, ColorsNamed.Orange, "BoldCenter", word_wrap=word_wrap)
    c_start_data = c_start.addRowCopy(1)

    range_data = doc.addListOfRowsWithStyle(c_start_data, lor, word_wrap=word_wrap)

    if column_of_totals or row_of_totals:
        cross_totals_from_range(doc, range_data, "#SUM", column_of_totals, row_of_totals, "BoldCenter", "BoldCenter", False)

    data_to_measure = [headers] + lor
    doc.setColumnsWidth(data_to_measure, columns_width_mode, char_to_cm, padding_cm, min_width_cm, max_width_cm)

    if freezeandselect:
        doc.freezeAndSelect(freezeandselect,freezeandselect, freezeandselect)
    
    return range_data

def sheet_split_with_big_lol(doc, sheet_name, lor, headers, headers_colors=ColorsNamed.Orange, coord_to_freeze="A2",  max_rows=1048575, word_wrap=False):
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
        word_wrap (bool, optional): Enable word wrap and optimal height. Defaults to False.
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
    

def sheet_from_lod(doc, sheetname, lod_,  column_of_totals=False, row_of_totals=False, freezeandselect=None, titulo=None, word_wrap=False, **kwargs_columnswidth):
    """
        kwargs son los parametros de la funcion setColumnsWidth
    """
    columns_width_mode=kwargs_columnswidth.get("columns_width_mode", types.ColumnsWidthMode.FROM_LOD)
    char_to_cm=kwargs_columnswidth.get("char_to_cm", 0.22)
    padding_cm=kwargs_columnswidth.get("padding_cm", 0.5)
    min_width_cm=kwargs_columnswidth.get("min_width_cm", 2.0)
    max_width_cm=kwargs_columnswidth.get("max_width_cm", 15.0)

    doc.createSheet(sheetname)
    if len(lod_)==0:
        if titulo:
            doc.addCellMergedWithStyle("A1:D1", titulo, ColorsNamed.Red, "BoldCenter")
        else:
            doc.addCellMergedWithStyle("A1:D1", "No hay datos", ColorsNamed.Red, "BoldCenter")
        return

         
    keys=lod_[0].keys()
    if titulo is None:
        c_start=Coord("A1")
    else:
        c_end=Coord("A1").addColumnCopy(len(keys)-1)
        range_=Range.from_coords("A1", c_end)
        doc.addCellMergedWithStyle(range_, titulo, ColorsNamed.Red, "BoldCenter", word_wrap=word_wrap)
        c_start=Coord("A2")#Empieza abajo
            
    
    range_final=block_from_lod(doc, c_start, lod_, column_of_totals=column_of_totals, row_of_totals=row_of_totals, word_wrap=word_wrap )
    if column_of_totals or row_of_totals:
        range_cross=cross_totals_from_range (doc, range_final, "#SUM", column_of_totals, row_of_totals, "BoldCenter", "BoldCenter", False)
    doc.setColumnsWidth(lod_, columns_width_mode, char_to_cm, padding_cm, min_width_cm, max_width_cm)
    if freezeandselect:
        doc.freezeAndSelect(freezeandselect,freezeandselect, freezeandselect)
    return range_cross if column_of_totals or row_of_totals else range_final




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
        word_wrap (bool, optional): Enable word wrap and optimal height. Defaults to False.

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
        doc.addCellMergedWithStyle(Range.from_coords(coord,coord.addColumnCopy(len(keys)-1)), titulo, ColorsNamed.Red, "BoldCenter", word_wrap=word_wrap)
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
    
