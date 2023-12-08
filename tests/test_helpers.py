from os import remove

from unogenerator import can_import_uno
if can_import_uno():
    from unogenerator import helpers, ODS_Standard, ColorsNamed
    headers=["A", "B", "C", "D"]
    lor=[[1, 2, 3, 4], [5, 6, 7, 8]]

    def test_helper_totals_column():
        with ODS_Standard() as doc:
            doc.addListOfRowsWithStyle("A1", lor)
            helpers.helper_totals_column(doc, "E1", ["#SUM"]*len(lor),column_from="A")
            helpers.helper_totals_column(doc, "F1", ["#SUM"]*len(lor),column_from="B", column_to="C",styles=["BoldCenter"]*len(lor))
            doc.export_pdf("helper_totals_column.pdf")
        
        remove("helper_totals_column.pdf")
        
    def test_helper_totals_row():
        with ODS_Standard() as doc:
            doc.addListOfRowsWithStyle("A1", lor)
            helpers.helper_totals_row(doc, "A3", ["#SUM"]*len(lor[0]),row_from="1")
            doc.export_pdf("test_helper_totals_row.pdf")    
        remove("test_helper_totals_row.pdf")
        
    def test_helper_totals_from_range():
        with ODS_Standard() as doc:
            doc.createSheet("Both")
            doc.addRowWithStyle("A1", headers, ColorsNamed.Orange, "BoldCenter")
            range_=doc.addListOfRowsWithStyle("A2", lor)
            helpers.helper_totals_from_range(doc, range_,)
            doc.createSheet("Columns")
            doc.addRowWithStyle("A1", headers, ColorsNamed.Orange, "BoldCenter")
            range_=doc.addListOfRowsWithStyle("A2", lor)
            helpers.helper_totals_from_range(doc, range_, totalcolumns=True, totalrows=False)
            doc.createSheet("Rows")
            doc.addRowWithStyle("B1", headers, ColorsNamed.Orange, "BoldCenter")
            range_=doc.addListOfRowsWithStyle("B2", lor)
            helpers.helper_totals_from_range(doc, range_, totalcolumns=False, totalrows=True)
            doc.export_pdf("test_helper_totals_from_range.pdf")
        remove("test_helper_totals_from_range.pdf")
            
    def test_helper_list_of_ordereddicts():
        with ODS_Standard() as doc:
            helpers.helper_list_of_ordereddicts(doc, "A1", [])
        

    def test_helper_split_big_listofrows():
        r=[]
        for i in range (100):
            r.append([i, i+1])
            
        with ODS_Standard() as doc:
            helpers.helper_split_big_listofrows(doc, "Big LOR", r, ["N", "N+1"])#, headers_colors=ColorsNamed.Orange, columns_width=None,  coord_to_freeze="A2",  max_rows=1048575): 
            helpers.helper_split_big_listofrows(doc, "Big LOR de 10", r, ["N", "N+1"], columns_width=3, max_rows=10)#, headers_colors=ColorsNamed.Orange, columns_width=None,  coord_to_freeze="A2",  max_rows=1048575): 
            
            doc.export_xlsx("helper_split_big_listofrows.xlsx")
        
        remove("helper_split_big_listofrows.xlsx")
