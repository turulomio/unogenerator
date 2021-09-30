## @param cood Coord from we are going to add totals
## @param list_of_totals List with strings or keys. Example: ["Total", "#SUM", "#AVG"]...
## @param styles List with string styles or None. If none tries to guest from top column object. List example: ["GrayLightPercentage", "GrayLightInteger"]
## @param string with the row where th3e total begins
## @param string with the rew where the formula ends. If None it's a coord.row -1
from unogenerator.commons import ColorsNamed, Coord as C, Range as R, guess_object_style, generate_formula_total_string


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
def helper_totals_from_range(doc, range_of_data, key="#SUM", totalcolumns=True, totalrows=True):
    range=R.assertRange(range_of_data)
    coord_start=range.start
    data_rows=range.numRows()
    data_columns=range.numColumns()
    horizontal_start=coord_start.addColumnCopy(-1) #Not logical, it's emprical, por eso data_start
    vertical_start=coord_start.addRowCopy(-1).addColumnCopy(-1)
    vertical_start=coord_start.addRowCopy(-1).addColumnCopy(-1)
    if totalcolumns==True and totalrows==True:
        helper_totals_row(doc, horizontal_start.addRowCopy(data_rows),["Total"]+[key]*(data_columns+1),styles=None, row_from=coord_start.number)
        helper_totals_column(doc, vertical_start.addColumnCopy(data_columns+1),["Total"]+[key]*(data_rows+1), styles=None, column_from=coord_start.letter)
    elif totalcolumns==True:
        helper_totals_column(doc, vertical_start.addColumnCopy(data_columns+1),["Total"]+[key]*(data_rows+1), styles=None, column_from=coord_start.letter)
    elif totalrows==True:
        helper_totals_row(doc, horizontal_start.addRowCopy(data_rows),["Total"]+[key]*(data_columns+0), styles=None, row_from=coord_start.number) #1 menos por la esquina
 