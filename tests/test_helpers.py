from os import remove

from unogenerator import can_import_uno
if can_import_uno():
    from unogenerator import helpers, ODS_Standard, ColorsNamed
    headers=["A", "B", "C", "D"]
    lor=[[1, 2, 3, 4], [5, 6, 7, 8]]

    def test_column_totals(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            doc.addListOfRowsWithStyle("A1", lor)
            helpers.column_totals(doc, "E1", ["#SUM"]*len(lor),column_from="A")
            helpers.column_totals(doc, "F1", ["#SUM"]*len(lor),column_from="B", column_to="C",styles=["BoldCenter"]*len(lor))
            doc.export_pdf("column_totals.pdf")
        
        remove("column_totals.pdf")
        
    def test_row_totals(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            doc.addListOfRowsWithStyle("A1", lor)
            helpers.row_totals(doc, "A3", ["#SUM"]*len(lor[0]),row_from="1")
            doc.export_pdf("test_row_totals.pdf")    
        remove("test_row_totals.pdf")
        
    def test_cross_totals_from_range(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            doc.createSheet("Both")
            doc.addRowWithStyle("A1", headers, ColorsNamed.Orange, "BoldCenter")
            range_=doc.addListOfRowsWithStyle("A2", lor)
            helpers.cross_totals_from_range(doc, range_,)
            doc.createSheet("Columns")
            doc.addRowWithStyle("A1", headers, ColorsNamed.Orange, "BoldCenter")
            range_=doc.addListOfRowsWithStyle("A2", lor)
            helpers.cross_totals_from_range(doc, range_, totalcolumns=True, totalrows=False)
            doc.createSheet("Rows")
            doc.addRowWithStyle("B1", headers, ColorsNamed.Orange, "BoldCenter")
            range_=doc.addListOfRowsWithStyle("B2", lor)
            helpers.cross_totals_from_range(doc, range_, totalcolumns=False, totalrows=True)
            doc.export_pdf("test_cross_totals_from_range.pdf")
        remove("test_cross_totals_from_range.pdf")
            
    def test_block_from_lod(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            helpers.block_from_lod(doc, "A1", [])
        

    def test_sheet_split_with_big_lol(libreoffice_server):
        r=[]
        for i in range (100):
            r.append([i, i+1])
            
        with ODS_Standard(server=libreoffice_server) as doc:
            helpers.sheet_split_with_big_lol(doc, "Big LOR", r, ["N", "N+1"])#, headers_colors=ColorsNamed.Orange, columns_width=None,  coord_to_freeze="A2",  max_rows=1048575): 
            helpers.sheet_split_with_big_lol(doc, "Big LOR de 10", r, ["N", "N+1"], columns_width=3, max_rows=10)#, headers_colors=ColorsNamed.Orange, columns_width=None,  coord_to_freeze="A2",  max_rows=1048575): 
            
            doc.export_xlsx("sheet_split_with_big_lol.xlsx")
        
        remove("sheet_split_with_big_lol.xlsx")
