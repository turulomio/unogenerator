from unogenerator.commons import ColorsNamed, Coord, Range, guess_object_style, generate_formula_total_string
from unogenerator import ODS, types
from pydicts import lod
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
        totalcolumns=True, 
        totalrows=True, 
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
        totalcolumns (bool, optional): Whether to generate column totals. Defaults to True.
        totalrows (bool, optional): Whether to generate row totals. Defaults to True.
        vertical_total_title_style (str, optional): Style for the vertical total title. Defaults to "BoldCenter".
        horizontal_total_title_style (str, optional): Style for the horizontal total title. Defaults to "BoldCenter".
        showing (bool, optional): If True, shows a 'Sum of totals' cell when either totalcolumns or totalrows is True. Defaults to False.

    Returns:
        Range: The original data range.
    """
    range=Range.assertRange(range_of_data)
    data_rows=range.numRows()
    data_columns=range.numColumns()
    coord_horizontal_title=range.c_start.addColumnCopy(-1).addRowCopy(data_rows) 
    coord_vertical_title=range.c_start.addRowCopy(-1).addColumnCopy(data_columns)
    style_data=guess_object_style(doc.getValue(range.c_end))
    
    if totalcolumns==True and totalrows==True:
        doc.addCellWithStyle(coord_horizontal_title, _("Total"), ColorsNamed.GrayLight, horizontal_total_title_style)
        row_totals(doc, coord_horizontal_title.addColumnCopy(1), [key]*data_columns,styles=style_data, row_from=range.c_start.number)
        doc.addCellWithStyle(coord_vertical_title, _("Total"), ColorsNamed.GrayLight, vertical_total_title_style)
        column_totals(doc, coord_vertical_title.addRowCopy(1),[key]*(data_rows+1), styles=style_data, column_from=range.c_start.letter)
    elif totalcolumns==True:
        doc.addCellWithStyle(coord_vertical_title, _("Total"), ColorsNamed.GrayLight, vertical_total_title_style)
        column_totals(doc, coord_vertical_title.addRowCopy(1),[key]*(data_rows+0), styles=style_data, column_from=range.c_start.letter)
        if showing is True:
            coord_sum_totals=coord_vertical_title.addRowCopy(data_rows+1)
            doc.addCellWithStyle(coord_sum_totals, generate_formula_total_string(key, range.c_start.addColumnCopy(data_columns+1), range.c_end.addColumnCopy(1)), ColorsNamed.GrayLight, style_data)
            doc.addCellWithStyle(coord_sum_totals.addColumnCopy(-1), _("Sum of totals"), ColorsNamed.GrayDark, style_data)
    elif totalrows==True:
        doc.addCellWithStyle(coord_horizontal_title, _("Total"), ColorsNamed.GrayLight, horizontal_total_title_style)
        row_totals(doc, coord_horizontal_title.addColumnCopy(1),[key]*(data_columns+0), styles=style_data, row_from=range.c_start.number) #1 menos por la esquina
        if showing is True:
            coord_sum_totals=coord_horizontal_title.addColumnCopy(data_columns+1)
            doc.addCellWithStyle(coord_sum_totals, generate_formula_total_string(key, range.c_start.addRowCopy(data_rows+1), range.c_end.addRowCopy(1)), ColorsNamed.GrayLight, style_data)
            doc.addCellWithStyle(coord_sum_totals.addRowCopy(-1), _("Sum of totals"), ColorsNamed.GrayDark, style_data)

    return range_of_data



## Write cells from a list of ordered dictionaries
## @param lod List of ordered dictionaries
## @param keys. If None write all keys, Else must be a list of keys
## @param columns_header. Integer with the number of columns to apply color_header
## @return Range. Returns the range of the data without headers. Useful to set totals.
def block_from_lod(doc, coord_start,  lod_, keys=None, columns_header=0,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None):
    coord_start=Coord.assertCoord(coord_start)
    
    if len(lod_)==0 and keys is None:
        doc.addCellWithStyle(coord_start, _("No data to show"), ColorsNamed.Red, "BoldCenter")
        return None

        
    #Header
    if keys is None:
        keys=lod.lod_keys(lod_)
    
    for column,  key in enumerate(keys):       
        doc.addCellWithStyle(coord_start.addColumnCopy(column), key, color_row_header, "BoldCenter")
    coord_data=coord_start.addRowCopy(1)
    
    
    lor=lod.lod2lol(lod_, keys)
    
    #Generate list of colors
    colors=[]
    for i in range(len(keys)):
        if i <= columns_header-1:
            colors.append(color_column_header)
        else:
            colors.append(color)
   
    #Generate list of rows
    return doc.addListOfRowsWithStyle(coord_data, lor, colors, styles)

def block_from_lod_with_totals(doc, coord_start,  lod, keys=None, columns_header=1,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None, totalcolumns=True, totalrows=True, key="#SUM"):
    """
    Writes data from a list of ordered dictionaries and appends totals.

    Args:
        doc (ODS): The ODS document object.
        coord_start (Coord or str): Starting coordinate.
        lod (list): List of ordered dictionaries containing the data.
        keys (list, optional): List of keys to write. Defaults to None.
        columns_header (int, optional): Number of leading columns treated as headers. Defaults to 1.
        color_row_header (int, optional): Color for the top header row. Defaults to ColorsNamed.Orange.
        color_column_header (int, optional): Color for the side header columns. Defaults to ColorsNamed.Green.
        color (int, optional): Default color for data cells. Defaults to ColorsNamed.White.
        styles (list or str, optional): Styles for data columns. Defaults to None.
        totalcolumns (bool, optional): Whether to generate column totals. Defaults to True.
        totalrows (bool, optional): Whether to generate row totals. Defaults to True.
        key (str, optional): Formula key to apply (e.g., "#SUM"). Defaults to "#SUM".

    Returns:
        Range: The range of the data including the generated totals.
    """
    print("QWUITAR CON parametros")
    coord_start=Coord.assertCoord(coord_start)
    block_from_lod(doc, coord_start,  lod, keys, columns_header,  color_row_header, color_column_header,  color, styles)
    range_lod=Range.from_iterable_object(coord_start.addRow(1), lod)## Adds q to skip top headers
    range_lod.c_start.addColumn(columns_header) ## Adds to skip columns headers
    return cross_totals_from_range (doc, range_lod, key, totalcolumns, totalrows)

def sheet_stylenames(doc):
    """
    Creates a new sheet called "Internal style names" listing all ODS styles grouped by families.

    Args:
        doc (ODS): The ODS document object.
    """
    lod_=[]
    for family,  style_names in doc.dict_stylenames.items():
        lod_.append({
            "Family":family,
            "Styles":style_names
        })
    sheet_from_lod(doc, "Internal style names", lod_, freezeandselect="A2", columns_width_mode=types.ColumnsWidthMode.FROM_LOD)

def sheet_from_lol(doc, sheetname, lor, headers, totalcolumns=False, totalrows=False, freezeandselect=None, titulo=None, **kwargs_columnswidth):
    """
    Creates a sheet from a list of lists (lol) with headers and optional totals.

    Args:
        doc (ODS): The ODS document object.
        sheetname (str): The name for the new sheet.
        lor (list): The list of lists (data rows).
        headers (list): The list of header strings.
        totalcolumns (bool, optional): Whether to generate column totals. Defaults to False.
        totalrows (bool, optional): Whether to generate row totals. Defaults to False.
        freezeandselect (str, optional): Coordinate to freeze panes at. Defaults to None.
        titulo (str, optional): An optional title to merge across the top of the sheet. Defaults to None.
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
        doc.addCellMergedWithStyle(range_titulo, titulo, ColorsNamed.Orange, "BoldCenter")
        c_start.addRow(1)

    doc.addRowWithStyle(c_start, headers, ColorsNamed.Orange, "BoldCenter")
    c_start_data = c_start.addRowCopy(1)

    range_data = doc.addListOfRowsWithStyle(c_start_data, lor)

    if totalcolumns or totalrows:
        cross_totals_from_range(doc, range_data, "#SUM", totalcolumns, totalrows, "BoldCenter", "BoldCenter", False)

    data_to_measure = [headers] + lor
    doc.setColumnsWidth(data_to_measure, columns_width_mode, char_to_cm, padding_cm, min_width_cm, max_width_cm)

    if freezeandselect:
        doc.freezeAndSelect(freezeandselect,freezeandselect, freezeandselect)
    
    return range_data

def sheet_split_with_big_lol(doc, sheet_name, lor, headers, headers_colors=ColorsNamed.Orange, coord_to_freeze="A2",  max_rows=1048575):
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
        doc.addRowWithStyle("A1", headers, headers_colors, "BoldCenter")
        
        #Splits data
        from_=max_rows*num_sheet
        to_=max_rows*(num_sheet+1) if len(lor)>=max_rows*(num_sheet+1) else len(lor)
        doc.addListOfRowsWithStyle("A2", lor[from_:to_])

        #Sets width of columns
    
        doc.setColumnsWidth(lor[from_:to_], types.ColumnsWidthMode.FROM_LOL)

        doc.freezeAndSelect(Coord.assertCoord(coord_to_freeze))
    

def sheet_from_lod(doc, sheetname, lod_,  totalcolumns=False, totalrows=False, freezeandselect=None, titulo=None, **kwargs_columnswidth):
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
        doc.addCellMergedWithStyle(range_, titulo, ColorsNamed.Red, "BoldCenter")
        c_start=Coord("A2")#Empieza abajo
            
    
    range_final=block_from_lod_with_totals(doc, c_start, lod_, totalcolumns=totalcolumns, totalrows=totalrows )
    if totalcolumns or totalrows:
        range_cross=cross_totals_from_range (doc, range_final, "#SUM", totalcolumns, totalrows, "BoldCenter", "BoldCenter", False)
    doc.setColumnsWidth(lod_, columns_width_mode, char_to_cm, padding_cm, min_width_cm, max_width_cm)
    if freezeandselect:
        doc.freezeAndSelect(freezeandselect,freezeandselect, freezeandselect)
    return range_cross if totalcolumns or totalrows else range_final




def block_from_lod_with_headers(doc, lod_, coord, subtitles=[], titulo=None, totalcolumns=False, totalrows=False, freezeandselect=None, key="#SUM"):
    """
    Writes data from a list of ordered dictionaries with custom header groups, and optional totals.

    Args:
        doc (ODS): The ODS document object.
        lod_ (list): List of ordered dictionaries containing the data.
        coord (Coord or str): Starting coordinate.
        subtitles (list): List of lists [title, first_key] defining header groups.
        titulo (str, optional): Main title for the entire block. Defaults to None.
        totalcolumns (bool, optional): Whether to generate column totals. Defaults to False.
        totalrows (bool, optional): Whether to generate row totals. Defaults to False.
        freezeandselect (str or Coord, optional): Coordinate to freeze and select. Defaults to None.
        key (str, optional): Formula key for totals (e.g., "#SUM"). Defaults to "#SUM".

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
        doc.addCellMergedWithStyle(Range.from_coords(coord,coord.addColumnCopy(len(keys)-1)), titulo, ColorsNamed.Red, "BoldCenter")
        coord.addRow(1)
     
         
    # Crea titulos
    for title, key_start, index_start, index_end in subtitles:
        c_start=coord.addColumnCopy(index_start)
        c_fin=coord.addColumnCopy(index_end)
        doc.addCellMergedWithStyle(Range.from_coords(c_start,c_fin), title, ColorsNamed.Orange, "BoldCenter")


    #Imprime listas de diccionarios
    range_=block_from_lod(doc, coord.addRowCopy(1),lod_,color_row_header=ColorsNamed.Yellow)

    if totalcolumns or totalrows:
        cross_totals_from_range(doc, range_, key, totalcolumns, totalrows)
        
    if freezeandselect:
        doc.freezeAndSelect(freezeandselect, freezeandselect, freezeandselect)

    return range_
    
