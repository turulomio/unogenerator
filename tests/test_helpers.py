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

    def test_sheet_photos_from_lod(libreoffice_server):
        assert helpers.sheet_photos_from_lod is helpers.sheet_photos_from_lod
        sample_png = (
            b"\x89PNG\r\n\x1a\n"
            b"\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
            b"\x00\x00\x00\x0dIDATx\x9cc`\x00\x00\x00\x02\x00\x01\x0e\xfe\x02\x06"
            b"\x00\x00\x00\x00IEND\xaeB`\x82"
        )
        lod_photos = [
            {"name": "Image 1", "photo_blob": sample_png, "width": 2.0, "height": 2.0},
            {"name": "Image 2", "photo_blob": sample_png, "width": 3.0, "height": 1.5},
            {"name": "Image 3", "photo_blob": b"", "width": 3.0, "height": 1.5},
            {"name": "Image 4", "photo_blob": None, "width": 3.0, "height": 1.5}
        ]
        with ODS_Standard(server=libreoffice_server) as doc:
            helpers.sheet_photos_from_lod(doc, "A1", lod_photos, headers=["Name", "Photo"], title="Test Photos Catalog")
            draw_page = doc.sheet.getDrawPage()
            assert draw_page.getCount() == 2
            # Check cell for b"" (row 5) and None (row 6)
            from unogenerator.commons import _
            assert doc.sheet.getCellRangeByName("B5").getString() == _("Image couldn't be loaded")
            assert doc.sheet.getCellRangeByName("B6").getString() == _("Image couldn't be loaded")
            doc.save("test_sheet_photos_from_lod.ods")

        remove("test_sheet_photos_from_lod.ods")

        # Test invalid bytes uses default 'Image couldn't be loaded'
        lod_photos_invalid = [
            {"name": "Image Invalid", "photo_blob": b"invalid_bytes", "width": 2.0, "height": 2.0}
        ]
        with ODS_Standard(server=libreoffice_server) as doc:
            helpers.sheet_photos_from_lod(doc, "A1", lod_photos_invalid, headers=["Name", "Photo"])
            assert doc.sheet.getCellRangeByName("B2").getString() == _("Image couldn't be loaded")

        # Test invalid bytes sets custom on_error_str when specified
        with ODS_Standard(server=libreoffice_server) as doc:
            helpers.sheet_photos_from_lod(
                doc, "A1", lod_photos_invalid, headers=["Name", "Photo"], on_error_str="[Image Error]"
            )
            assert doc.sheet.getCellRangeByName("B2").getString() == "[Image Error]"

    def test_styles_from_lod_and_lol():
        from unogenerator.commons import styles_from_lod, styles_from_lol
        lod_data = [
            {"id": 1, "price": None, "currency": None},
            {"id": None, "price": 19.99, "currency": Currency(10, "EUR")},
        ]
        styles = styles_from_lod(lod_data, keys=["id", "price", "currency"])
        assert styles == ["Integer", "Float2", "EUR"]

        lol_data = [
            [None, 10, Currency(5, "USD")],
            [1, None, None],
        ]
        styles_lol = styles_from_lol(lol_data)
        assert styles_lol == ["Integer", "Integer", "USD"]

    def test_block_from_lod_totals_styles(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            lod_data = [
                {"name": "Item A", "qty": 10, "amount": Currency(100.5, "EUR")},
                {"name": "Item B", "qty": 20, "amount": Currency(200.25, "EUR")},
            ]
            helpers.block_from_lod(doc, "A1", lod_data, row_of_totals=True)

            # Row 1: Headers ("name", "qty", "amount")
            # Row 2 & 3: Data rows
            # Row 4 (index 3): Totals row. B4 (col 1) = Integer, C4 (col 2) = EUR
            assert doc.sheet.getCellByPosition(1, 3).CellStyle == "Integer"
            assert doc.sheet.getCellByPosition(2, 3).CellStyle == "EUR"

    def test_cross_totals_from_range_styles_parameters(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            lod_data = [
                {"name": "Singer 1", "sales": Currency(50, "EUR")},
                {"name": "Singer 2", "sales": Currency(75, "EUR")},
            ]
            helpers.block_from_lod(doc, "A1", lod_data, column_of_totals=True, row_of_totals=True)

            # Row 1: Headers (A1: name, B1: sales, C1: Total)
            # Row 2 & 3: Data
            # Column C (index 2): Column of totals
            # Row 4 (index 3): Row of totals
            # Check column of totals gets EUR style automatically guessed per row
            assert doc.sheet.getCellByPosition(2, 1).CellStyle == "EUR"
            assert doc.sheet.getCellByPosition(2, 2).CellStyle == "EUR"

    def test_custom_styles_column_totals(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            lod_data = [
                {"name": "Singer 1", "sales": 50},
                {"name": "Singer 2", "sales": 75},
            ]
            helpers.sheet_from_lod(
                doc, "SheetTest", lod_data, 
                column_of_totals=True, row_of_totals=True, 
                styles_column_totals="Float2", styles_row_totals="Integer"
            )

            # Check column of totals got custom "Float2"
            assert doc.sheet.getCellByPosition(2, 1).CellStyle == "Float2"
            assert doc.sheet.getCellByPosition(2, 2).CellStyle == "Float2"
            # Check row of totals got custom "Integer"
            assert doc.sheet.getCellByPosition(1, 3).CellStyle == "Integer"

    def test_block_from_lod_with_headers_merged_total_header(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            lod_singers = [
                {"Singer": "Elvis", "Songs": 10000, "Albums": 100, "Best song": "Always on my mind"},
                {"Singer": "Roy Orbison", "Songs": 100, "Albums": 20, "Best song": "Crying"},
            ]
            helpers.block_from_lod_with_headers(
                doc, lod_singers, "A1", 
                [["Singer header", "Singer"], ["Song header", "Best song"]],
                titulo="Title", column_of_totals=True
            )

            # Row 1: Merged title (A1:E1)
            # Row 2 (index 1): Subtitles. Col E (index 4) should be merged with Row 3 (index 2)
            cell_e2 = doc.sheet.getCellByPosition(4, 1)
            assert cell_e2.getString() == "Total"
            assert cell_e2.IsMerged





