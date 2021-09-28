## @param cood Coord from we are going to add totals
## @param list_of_totals List with strings or keys. Example: ["Total", "#SUM", "#AVG"]...
## @param list_of_styles List with string styles or None. If none tries to guest from top column object. List example: ["GrayLightPercentage", "GrayLightInteger"]
## @param string with the row where th3e total begins
## @param string with the rew where the formula ends. If None it's a coord.row -1
from unogenerator.commons import ColorsNamed, Coord as C


def addTotalsHorizontal(doc, coord, list_of_totals, list_of_styles=None, row_from="2", row_to=None):
    coord=C.assertCoord(coord)
    for letter, total in enumerate(list_of_totals):
        coord_total=coord.addColumnCopy(letter)
        coord_total_from=C(coord_total.letter+row_from)
        if row_to is None:
            coord_total_to=coord_total.addRowCopy(-1)# row above
        else:
            coord_total=C(coord_total.letter+row_to)

        if list_of_styles is None:
            style=guess_ods_style("GrayLight", self.getCellValue(coord_total_from))
        else:
            style=list_of_styles[letter]

        doc.add(coord_total, generate_formula_total_string(total, coord_total_from, coord_total_to),  style)
        
def helper_values_with_total(doc, coord, title, values,  horizontal=True, style_title="Bold", color_title=ColorsNamed.Orange, style_values=None, color_values=ColorsNamed.White, style_total=None, color_total=ColorsNamed.GrayLight):
    coord=C.assertCoord(coord)
    i=0
    if title is not None:
        doc.addCellWithStyle(coord,title,color_title,style_title)
        i=i+1
    doc.addRowWithStyle(coord.addRowCopy(i),values,colors=color_values,styles=style_values)
    doc.addCellWithStyle(coord.addRowCopy(i+len(values)+1),f"=sum({coord.addRowCopy(i)}:{coord.addRowCopy(i+len(values))}",color_title,style_title)
