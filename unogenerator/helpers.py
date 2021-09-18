## @param cood Coord from we are going to add totals
## @param list_of_totals List with strings or keys. Example: ["Total", "#SUM", "#AVG"]...
## @param list_of_styles List with string styles or None. If none tries to guest from top column object. List example: ["GrayLightPercentage", "GrayLightInteger"]
## @param string with the row where th3e total begins
## @param string with the rew where the formula ends. If None it's a coord.row -1
def addTotalsHorizontal(self, coord, list_of_totals, list_of_styles=None, row_from="2", row_to=None):
    coord=Coord.assertCoord(coord)
    for letter, total in enumerate(list_of_totals):
        coord_total=coord.addColumnCopy(letter)
        coord_total_from=Coord(coord_total.letter+row_from)
        if row_to is None:
            coord_total_to=coord_total.addRowCopy(-1)# row above
        else:
            coord_total=Coord(coord_total.letter+row_to)

        if list_of_styles is None:
            style=guess_ods_style("GrayLight", self.getCellValue(coord_total_from))
        else:
            style=list_of_styles[letter]

        self.add(coord_total, generate_formula_total_string(total, coord_total_from, coord_total_to),  style)
        
def helper_values_with_total(self, title, values, total=True,  horizontal=True, style_title=None, color_title=None, style_values=None, color_values=None, total_style=None, color_style=None):
    pass
