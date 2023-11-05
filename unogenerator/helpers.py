## @param cood Coord from we are going to add totals
## @param list_of_totals List with strings or keys. Example: ["Total", "#SUM", "#AVG"]...
## @param styles List with string styles or None. If none tries to guest from top column object. List example: ["GrayLightPercentage", "GrayLightInteger"]
## @param string with the row where th3e total begins
## @param string with the rew where the formula ends. If None it's a coord.row -1
from unogenerator.commons import ColorsNamed, Coord as C, Range as R, guess_object_style, generate_formula_total_string
from pydicts import lod
from gettext import translation
from logging import debug
from math import ceil
from importlib.resources import files

try:
    t=translation('unogenerator', files("unogenerator") / 'locale')
    _=t.gettext
except:
    _=str

def helper_totals_row(doc, coord, list_of_totals, color=ColorsNamed.GrayLight, styles=None, row_from="2", row_to=None):
    coord=C.assertCoord(coord)
    for letter, total in enumerate(list_of_totals):
        coord_total=coord.addColumnCopy(letter)
        coord_total_from=C(coord_total.letter+row_from)
        if row_to is None:
            coord_total_to=coord_total.addRowCopy(-1)# row above
        else:
            coord_total_to=C(coord_total.letter+row_to)

        if styles is None:
            style=guess_object_style(doc.getValue(coord_total_from))
        elif styles.__class__.__name__ != "list":
            style=styles
        else:
            style=styles[letter]

        doc.addCellWithStyle(coord_total, generate_formula_total_string(total, coord_total_from, coord_total_to), color, style)


def helper_totals_column(doc, coord, list_of_totals, color=ColorsNamed.GrayLight, styles=None, column_from="B", column_to=None):
    coord=C.assertCoord(coord)
    for number, total in enumerate(list_of_totals):
        coord_total=coord.addRowCopy(number)
        coord_total_from=C(column_from + coord_total.number)
        if column_to is None:
            coord_total_to=coord_total.addColumnCopy(-1)# row above
        else:
            coord_total_to=C(column_to + coord_total.number)

        if styles is None:
            style=guess_object_style(doc.getValue(coord_total_from))
        elif styles.__class__.__name__ != "list":
            style=styles
        else:
            style=styles[number]

        doc.addCellWithStyle(coord_total, generate_formula_total_string(total, coord_total_from, coord_total_to), color, style)
        
def helper_title_values_total_row( doc, coord, title, values, 
        style_title=None, color_title=ColorsNamed.Orange, 
        style_values=None, color_values=ColorsNamed.White, 
        style_total=None, color_total=ColorsNamed.GrayLight
    ):
    coord=C.assertCoord(coord)

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

def helper_title_values_total_column(doc, coord, title, values,
        style_title=None, color_title=ColorsNamed.Orange, 
        style_values=None, color_values=ColorsNamed.White, 
        style_total=None, color_total=ColorsNamed.GrayLight
    ):
    coord=C.assertCoord(coord)

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
        

## Genera totales verticales y horizontales directamente partiendo de un rango. Todos con sumas, añade un "Total" en la fila columna anterior
## @param s xlsx doc
## @param range_of_data. Range with data values
## @param keys Key of formula if List, it has all values
## @param showing When totalcolumns=True or totalrows=True only, shows a Total of Totals
def helper_totals_from_range (
                                                    doc, 
                                                    range_of_data, 
                                                    key="#SUM", 
                                                    totalcolumns=True, 
                                                    totalrows=True, 
                                                    vertical_total_title_style="BoldCenter", 
                                                    horizontal_total_title_style="BoldCenter", 
                                                    showing=False
                                                ):
    range=R.assertRange(range_of_data)
    data_rows=range.numRows()
    data_columns=range.numColumns()
    coord_horizontal_title=range.c_start.addColumnCopy(-1).addRowCopy(data_rows) 
    coord_vertical_title=range.c_start.addRowCopy(-1).addColumnCopy(data_columns)
    style_data=guess_object_style(doc.getValue(range.c_end))
    
    if totalcolumns==True and totalrows==True:
        doc.addCellWithStyle(coord_horizontal_title, _("Total"), ColorsNamed.GrayLight, horizontal_total_title_style)
        helper_totals_row(doc, coord_horizontal_title.addColumnCopy(1), [key]*data_columns,styles=style_data, row_from=range.c_start.number)
        doc.addCellWithStyle(coord_vertical_title, _("Total"), ColorsNamed.GrayLight, vertical_total_title_style)
        helper_totals_column(doc, coord_vertical_title.addRowCopy(1),[key]*(data_rows+1), styles=style_data, column_from=range.c_start.letter)
    elif totalcolumns==True:
        doc.addCellWithStyle(coord_vertical_title, _("Total"), ColorsNamed.GrayLight, vertical_total_title_style)
        helper_totals_column(doc, coord_vertical_title.addRowCopy(1),[key]*(data_rows+0), styles=style_data, column_from=range.c_start.letter)
        if showing is True:
            coord_sum_totals=coord_vertical_title.addRowCopy(data_rows+1)
            doc.addCellWithStyle(coord_sum_totals, generate_formula_total_string(key, range.c_start.addColumnCopy(data_columns+1), range.c_end.addColumnCopy(1)), ColorsNamed.GrayLight, style_data)
            doc.addCellWithStyle(coord_sum_totals.addColumnCopy(-1), _("Sum of totals"), ColorsNamed.GrayDark, style_data)
    elif totalrows==True:
        doc.addCellWithStyle(coord_horizontal_title, _("Total"), ColorsNamed.GrayLight, horizontal_total_title_style)
        helper_totals_row(doc, coord_horizontal_title.addColumnCopy(1),[key]*(data_columns+0), styles=style_data, row_from=range.c_start.number) #1 menos por la esquina
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
def helper_list_of_ordereddicts(doc, coord_start,  lod_, keys=None, columns_header=0,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None):
    coord_start=C.assertCoord(coord_start)
    
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

## Write cells from a list of ordered dictionaries
## @param lod List of ordered dictionaries
## @param keys. If None write all keys, Else must be a list of keys
## @param columns_header. Integer with the number of columns to apply color_header
## @return Range of the data
def helper_list_of_ordereddicts_with_totals(doc, coord_start,  lod, keys=None, columns_header=1,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None, totalcolumns=True, totalrows=True, key="#SUM"):
    coord_start=C.assertCoord(coord_start)
    helper_list_of_ordereddicts(doc, coord_start,  lod, keys, columns_header,  color_row_header, color_column_header,  color, styles)
    range_lod=R.from_iterable_object(coord_start.addRow(1), lod)## Adds q to skip top headers
    range_lod.c_start.addColumn(columns_header) ## Adds to skip columns headers
    return helper_totals_from_range (doc, range_lod, key, totalcolumns, totalrows)
    
## It's the same of helper_list_of_ordereddicts but withth mandatory keys
## @return Range of the data
def helper_list_of_dicts(doc, coord_start,  lod, keys, columns_header=0,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None):
    return helper_list_of_ordereddicts(doc, coord_start,  lod, keys, columns_header=0,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None)

## Creates a new sheet called "Style names" with alll ods styles grouped by families
def helper_ods_sheet_stylenames(doc):
    doc.createSheet("Internal style names")
    doc.setColumnsWidth([5, 5])
    for column, (family,  style_names) in enumerate(doc.dict_stylenames.items()):
        doc.addCellWithStyle(C("A1").addColumn(column), family, ColorsNamed.Orange, "BoldCenter")
        doc.addColumnWithStyle(C("A2").addColumn(column), style_names)
    doc.freezeAndSelect("A2")

## This helper is used when lor length is bigger than localc limits (1048576)
## With this function you can split a lor automatically in all sheets needed
## If the number of rows is lower than max_rows makes a normal sheet
## You have to set a header of one line. 
## @param doc ODS Document
## @param sheet_name Root name of the sheet
## @param lor List of rows with data
## @param headers List of strings
## @param headers_colors Color of the sheet header
## @param columns_width None=3cm. Integer=all that value, List. Defines all columns width
## @param coord_to_freeze Coord with coord to freeze
def helper_split_big_listofrows(doc, sheet_name, lor, headers, headers_colors=ColorsNamed.Orange, columns_width=None,  coord_to_freeze="A2",  max_rows=1048575):
    ceil_=ceil(len(lor)/max_rows)
    for num_sheet in range(ceil_):
        #Sets name and headers
        if ceil_>1:
            name=_("{0} ({1} of {2})").format(sheet_name, num_sheet+1,  ceil_)
            debug(_("More than {0} rows. Spliting {1} of {2} sheets").format(max_rows, num_sheet+1,  ceil_))
        else:
            name=sheet_name
        doc.createSheet(name)
        doc.addRowWithStyle("A1", headers, headers_colors, "BoldCenter")
        
        #Sets width of columns
        if len(lor)>0:
            if columns_width is None:
                columns_width=[3]*len(lor)
            elif columns_width.__class__.__name__=="int":
                columns_width=[columns_width]*len(lor)
            doc.setColumnsWidth(columns_width)
        
        #Splits data
        from_=max_rows*num_sheet
        to_=max_rows*(num_sheet+1) if len(lor)>=max_rows*(num_sheet+1) else len(lor)
        doc.addListOfRowsWithStyle("A2", lor[from_:to_])
        doc.freezeAndSelect(C.assertCoord(coord_to_freeze))
    
    
