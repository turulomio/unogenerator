## @param cood Coord from we are going to add totals
## @param list_of_totals List with strings or keys. Example: ["Total", "#SUM", "#AVG"]...
## @param styles List with string styles or None. If none tries to guest from top column object. List example: ["GrayLightPercentage", "GrayLightInteger"]
## @param string with the row where th3e total begins
## @param string with the rew where the formula ends. If None it's a coord.row -1
from unogenerator.commons import ColorsNamed, Coord as C, Range as R, guess_object_style, generate_formula_total_string

from gettext import translation
from pkg_resources import resource_filename

try:
    t=translation('unogenerator', resource_filename("unogenerator","locale"))
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
def helper_totals_from_range (
                                                    doc, 
                                                    range_of_data, 
                                                    key="#SUM", 
                                                    totalcolumns=True, 
                                                    totalrows=True, 
                                                    vertical_total_title_style="BoldCenter", 
                                                    horizontal_total_title_style="BoldCenter"
                                                ):
    range=R.assertRange(range_of_data)
    data_rows=range.numRows()
    data_columns=range.numColumns()
    coord_horizontal_title=range.start.addColumnCopy(-1).addRowCopy(data_rows) 
    coord_vertical_title=range.start.addRowCopy(-1).addColumnCopy(data_columns)
    style_data=guess_object_style(doc.getValue(range.start))
    
    if totalcolumns==True and totalrows==True:
        doc.addCellWithStyle(coord_horizontal_title, _("Total"), ColorsNamed.GrayLight, horizontal_total_title_style)
        helper_totals_row(doc, coord_horizontal_title.addColumnCopy(1), [key]*data_columns,styles=style_data, row_from=range.start.number)
        doc.addCellWithStyle(coord_vertical_title, _("Total"), ColorsNamed.GrayLight, vertical_total_title_style)
        helper_totals_column(doc, coord_vertical_title.addRowCopy(1),[key]*(data_rows+1), styles=style_data, column_from=range.start.letter)
    elif totalcolumns==True:
        doc.addCellWithStyle(coord_vertical_title, _("Total"), ColorsNamed.GrayLight, vertical_total_title_style)
        helper_totals_column(doc, coord_vertical_title.addRowCopy(1),[key]*(data_rows+1), styles=style_data, column_from=range.start.letter)
        doc.addCellWithStyle(coord_vertical_title.addRowCopy(data_rows+1), generate_formula_total_string(key, range.start.addColumnCopy(data_columns+1), range.end.addColumnCopy(1)), ColorsNamed.GrayLight, style_data)
    elif totalrows==True:
        doc.addCellWithStyle(coord_horizontal_title, _("Total"), ColorsNamed.GrayLight, horizontal_total_title_style)
        helper_totals_row(doc, coord_horizontal_title.addColumnCopy(1),[key]*(data_columns+0), styles=style_data, row_from=range.start.number) #1 menos por la esquina
        doc.addCellWithStyle(coord_horizontal_title.addColumnCopy(data_columns+1), generate_formula_total_string(key, range.start.addRowCopy(data_rows+1), range.end.addRowCopy(1)), ColorsNamed.GrayLight, style_data)


## Write cells from a list of ordered dictionaries
## @param lod List of ordered dictionaries
## @param keys. If None write all keys, Else must be a list of keys
## @param columns_header. Integer with the number of columns to apply color_header
def helper_list_of_ordereddicts(doc, coord_start,  lod, keys=None, columns_header=0,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None):
    coord_start=C.assertCoord(coord_start)
    
    if len(lod)==0 and keys is None:
        doc.addCellWithStyle(coord_start, _("No data to show"), ColorsNamed.Red, "BoldCenter")
        return

        
    #Header
    if keys is None:
        keys=lod[0].keys()
    
    for column,  key in enumerate(keys):       
        doc.addCellWithStyle(coord_start.addColumnCopy(column), key, color_row_header, "BoldCenter")
    coord_data=coord_start.addRowCopy(1)
    
    #Data
    for row, od in enumerate(lod):
        for column, key in enumerate(keys):
            if styles is None:
                style=guess_object_style(od[key])
            elif styles.__class__.__name__ != "list":
                style=styles
            else:
                style=styles[column]

            if column+1<=columns_header:
                color_=color_column_header
            else:
                color_=color
            
            doc.addCellWithStyle(coord_data.addRowCopy(row).addColumnCopy(column), od[key], color_, style)


## It's the same of helper_list_of_ordereddicts but withth mandatory keys
def helper_list_of_dicts(doc, coord_start,  lod, keys, columns_header=0,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None):
    helper_list_of_ordereddicts(doc, coord_start,  lod, keys, columns_header=0,  color_row_header=ColorsNamed.Orange, color_column_header=ColorsNamed.Green,  color=ColorsNamed.White, styles=None)
