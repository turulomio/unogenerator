## @param cood Coord from we are going to add totals
## @param list_of_totals List with strings or keys. Example: ["Total", "#SUM", "#AVG"]...
## @param list_of_styles List with string styles or None. If none tries to guest from top column object. List example: ["GrayLightPercentage", "GrayLightInteger"]
## @param string with the row where th3e total begins
## @param string with the rew where the formula ends. If None it's a coord.row -1
from unogenerator.commons import ColorsNamed, Coord as C, guess_object_style, generate_formula_total_string


def helper_totals_row(doc, coord, list_of_totals, color=ColorsNamed.GrayLight, list_of_styles=None, row_from="2", row_to=None):
    coord=C.assertCoord(coord)
    for letter, total in enumerate(list_of_totals):
        coord_total=coord.addColumnCopy(letter)
        coord_total_from=C(coord_total.letter+row_from)
        if row_to is None:
            coord_total_to=coord_total.addRowCopy(-1)# row above
        else:
            coord_total_to=C(coord_total.letter+row_to)

        if list_of_styles is None:
            style=guess_object_style(doc.getValue(coord_total_from))
        else:
            style=list_of_styles[letter]

        doc.addCellWithStyle(coord_total, generate_formula_total_string(total, coord_total_from, coord_total_to), color, style)


def helper_totals_column(doc, coord, list_of_totals, color=ColorsNamed.GrayLight, list_of_styles=None, column_from="B", column_to=None):
    coord=C.assertCoord(coord)
    for number, total in enumerate(list_of_totals):
        coord_total=coord.addRowCopy(number)
        coord_total_from=C(column_from + coord_total.number)
        if column_to is None:
            coord_total_to=coord_total.addColumnCopy(-1)# row above
        else:
            coord_total_to=C(column_to + coord_total.number)

        if list_of_styles is None:
            style=guess_object_style(doc.getValue(coord_total_from))
        else:
            style=list_of_styles[number]

        doc.addCellWithStyle(coord_total, generate_formula_total_string(total, coord_total_from, coord_total_to), color, style)
        
def helper_values_with_total(
        doc, coord, title, values, horizontal=True, 
        style_title=None, color_title=ColorsNamed.Orange, 
        style_values=None, color_values=ColorsNamed.White, 
        style_total=None, color_total=ColorsNamed.GrayLight
    ):
    coord=C.assertCoord(coord)

    if style_title is None:
        if horizontal is True:
            style_title="Bold"
        else: #Vertical
            style_title="BoldCenter"

    if style_total is None and len(values)>0:
        style_total=guess_object_style(values[0])


    i=0
    if title is not None:
        doc.addCellWithStyle(coord,title,color_title,style_title)
        i=i+1

    if horizontal is True:
        doc.addRowWithStyle(coord.addColumnCopy(i),values,colors=color_values,styles=style_values)
        doc.addCellWithStyle(coord.addColumnCopy(i+len(values)),f"=sum({coord.addColumnCopy(i).string()}:{coord.addColumnCopy(i+len(values)-1).string()}",color_total,style_total)
    else:
        doc.addColumnWithStyle(coord.addRowCopy(i),values,colors=color_values,styles=style_values)
        doc.addCellWithStyle(coord.addRowCopy(i+len(values)),f"=sum({coord.addRowCopy(i).string()}:{coord.addRowCopy(i+len(values)-1).string()}",color_total,style_total)
