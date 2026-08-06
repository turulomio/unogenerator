from os import remove
from datetime import date, datetime, time, timedelta
from pydicts.currency import Currency
from pydicts.percentage import Percentage

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
            helpers.cross_totals_from_range(doc, range_, column_of_totals=True, row_of_totals=False)
            doc.createSheet("Rows")
            doc.addRowWithStyle("B1", headers, ColorsNamed.Orange, "BoldCenter")
            range_=doc.addListOfRowsWithStyle("B2", lor)
            helpers.cross_totals_from_range(doc, range_, column_of_totals=False, row_of_totals=True)
            doc.export_pdf("test_cross_totals_from_range.pdf")
        remove("test_cross_totals_from_range.pdf")
            
    def test_block_from_lod(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            helpers.block_from_lod(doc, "A1", [])
        

    def test_block_from_lod_with_null_first_row(libreoffice_server):
        """
        Test that styles are correctly guessed in block_from_lod for all types
        even if the first row has None values.
        """
        with ODS_Standard(server=libreoffice_server) as doc:
            lod_data = [
                {
                    "int": None, 
                    "timedelta": None, 
                    "currency": None, 
                    "percentage": None, 
                    "datetime": None, 
                    "date": None, 
                    "time": None, 
                    "bool": None,
                    "float": None,
                    "string": None,
                    "all_null": None
                },
                {
                    "int": 10, 
                    "timedelta": timedelta(seconds=10), 
                    "currency": Currency(10, "EUR"), 
                    "percentage": Percentage(1, 10), 
                    "datetime": datetime.now(), 
                    "date": date.today(), 
                    "time": time(12, 0), 
                    "bool": True,
                    "float": 10.5,
                    "string": "hello",
                    "all_null": None
                }
            ]

            helpers.block_from_lod(doc, "A1", lod_data)

            # Headers in row 1, Data in row 2 and 3
            # Column mapping:
            # A: int, B: timedelta, C: currency, D: percentage, E: datetime, 
            # F: date, G: time, H: bool, I: float, J: string, K: all_null

            expected = [
                (0, "Integer"),
                (1, "TimedeltaSeconds"),
                (2, "EUR"),
                (3, "Percentage"),
                (4, "Datetime"),
                (5, "Date"),
                (6, "Time"),
                (7, "Bool"),
                (8, "Float2"),
                (9, "Normal"),
                (10, "Normal") # all_null -> Default
            ]

            for col_idx, expected_style in expected:
                # Check row 2 (index 1) and row 3 (index 2)
                assert doc.sheet.getCellByPosition(col_idx, 1).CellStyle == expected_style
                assert doc.sheet.getCellByPosition(col_idx, 2).CellStyle == expected_style





    def test_sheet_split_with_big_lol(libreoffice_server):
        r=[]
        for i in range (100):
            r.append([i, i+1])
            
        with ODS_Standard(server=libreoffice_server) as doc:
            helpers.sheet_split_with_big_lol(doc, "Big LOR", r, ["N", "N+1"])#, headers_colors=ColorsNamed.Orange, columns_width=None,  coord_to_freeze="A2",  max_rows=1048575): 
            helpers.sheet_split_with_big_lol(doc, "Big LOR de 10", r, ["N", "N+1"], max_rows=10)#, headers_colors=ColorsNamed.Orange, columns_width=None,  coord_to_freeze="A2",  max_rows=1048575): 
            
            doc.export_xlsx("sheet_split_with_big_lol.xlsx")
        
        remove("sheet_split_with_big_lol.xlsx")

    def test_photos_from_lod_ods(libreoffice_server):
        sample_png = (
            b"\x89PNG\r\n\x1a\n"
            b"\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
            b"\x00\x00\x00\x0dIDATx\x9cc`\x00\x00\x00\x02\x00\x01\x0e\xfe\x02\x06"
            b"\x00\x00\x00\x00IEND\xaeB`\x82"
        )
        lod_photos = [
            {"name": "Image 1", "photo_blob": sample_png, "width": 2.0, "height": 2.0},
            {"name": "Image 2", "photo_blob": sample_png, "width": 3.0, "height": 1.5},
            {"name": "Image 2", "photo_blob": b"", "width": 3.0, "height": 1.5},
            {"name": "Image 2", "photo_blob": None, "width": 3.0, "height": 1.5}
        ]
        with ODS_Standard(server=libreoffice_server) as doc:
            helpers.photos_from_lod_ods(doc, "A1", lod_photos, headers=["Name", "Photo"], title="Test Photos Catalog")
            draw_page = doc.sheet.getDrawPage()
            assert draw_page.getCount() == 2
            doc.save("test_photos_from_lod_ods.ods")

        remove("test_photos_from_lod_ods.ods")

