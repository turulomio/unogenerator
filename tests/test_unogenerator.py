from datetime import date, datetime, time, timedelta
from decimal import Decimal
from os import remove, path
from pytest import raises
from pydicts import casts, currency, percentage, lod
from importlib.resources import files
import logging

from unogenerator import can_import_uno, types

logger = logging.getLogger(__name__) # Get logger for this module

if can_import_uno():
    from unogenerator import ODT_Standard, ODT, ODS_Standard, ODS, ColorsNamed, Range, Coord, exceptions
    from unogenerator import LibreofficeServer # Explicitly import LibreofficeServer
    
    
    row=[1, 2, 3, 4, 5]

    lor=[]
    for i in range(4):
        lor.append(row)
        
    now=casts.dtnaive_now()
    lod_types=[
    {
        "datetime": casts.dtnaive_now(), 
        "date":date.today(), 
        "integer":10000, 
        "currency": currency.Currency(12, "EUR"), 
        "percentage": percentage.Percentage(1, 2), 
        "float": 12.24, 
        "timedelta_seconds": casts.dtnaive_now()-(now-timedelta(seconds=6000)), 
        "timedelta_iso": casts.timedelta2str(casts.dtnaive_now()-(now-timedelta(seconds=6000))), 
        "time":time(12, 12, 12), 
        "bool":True, 
        "formula":"=2+2"
    }, 
    {
        "datetime": casts.dtnaive_now(), 
        "date":date.today(), 
        "integer":-10000, 
        "currency": currency.Currency(-12, "EUR"), 
        "percentage": percentage.Percentage(-1, 2), 
        "float": -12.24, 
        "timedelta_seconds": casts.dtnaive_now()-now, 
        "timedelta_iso": casts.timedelta2str(casts.dtnaive_now()-now), 
        "time":time(12, 12, 12), 
        "bool":False, 
        "formula":"=2+2"
    }
    ]

    lor_types=lod.lod2lol(lod_types)
    lor_types_styles=["Datetime", "Date", "Integer", "EUR", "Percentage", "Float2",  "TimedeltaSeconds","TimedeltaISO", "Time", "Bool", "Integer"]



    def test_ods_row_height_consistency(libreoffice_server):
        """
        Regression test for row height bug when using setColumnsWidth with many rows/columns.
        """
        with ODS_Standard(server=libreoffice_server) as doc:
            num_rows = 250
            num_cols = 20
            # Data that would trigger height increase to 841 if glitchy (e.g. 11+ chars)
            lod_data = []
            for i in range(num_rows):
                d = {f"Col{j}": f"Row{i:03} Col{j:02}" for j in range(num_cols)}
                lod_data.append(d)
            
            # 1. Add data
            lol_data = lod.lod2lol(lod_data)
            doc.addListOfRowsWithStyle("A1", lol_data)
            
            # 2. Apply column widths (this used to trigger the height glitch)
            doc.setColumnsWidth(lod_data, types.ColumnsWidthMode.FROM_LOD, char_to_cm=0.25)
            
            # 3. Verify row heights are consistent (Standard is 452)
            # Check row 2 (index 1) as baseline
            h_ref = doc.sheet.getRows().getByIndex(1).Height
            assert h_ref == 452, f"Expected default row height 452, got {h_ref}"
            
            # Check a sampling of rows including those > 100
            for i in [50, 100, 150, 200, 249]:
                h = doc.sheet.getRows().getByIndex(i).Height
                assert h == h_ref, f"Row {i+1} height {h} differs from reference {h_ref}"

    def test_ods_normal_style_applied(libreoffice_server):
        """
        Verify that ODS_Standard correctly applies 'Normal' style instead of 'Default'.
        """
        with ODS_Standard(server=libreoffice_server) as doc:
            doc.addCellWithStyle("A1", "Test Style", style="Normal")
            cell = doc.sheet.getCellByPosition(0, 0)
            assert cell.CellStyle == "Normal", f"Expected style 'Normal', got '{cell.CellStyle}'"
            
            # Test auto-guessing style also uses Normal as default
            doc.addListOfRowsWithStyle("A2", [["Guess Me"]])
            cell2 = doc.sheet.getCellByPosition(0, 1)
            assert cell2.CellStyle == "Normal", f"Expected guessed style 'Normal', got '{cell2.CellStyle}'"

    def test_odt_metadata(libreoffice_server):
        with ODT_Standard(server=libreoffice_server) as doc:
            doc.setMetadata(
                title="Metadata example",
                subject="UnoGenerator testing",
                author="Turulomio",
                description="This testing works with metadata methods",
                keywords=["Metadata", "Testing"],
            )
            doc.save("delete_metadata.odt")
            #Bad export
            doc.export_docx("delete_metadata.odt")
            doc.export_pdf("delete_metadata.odt")

        with ODT("delete_metadata.odt") as doc:
            assert doc.getMetadata()["Title"]== "Metadata example"
            
        remove("delete_metadata.odt")



    def test_odt_export(libreoffice_server):
        with ODT_Standard(server=libreoffice_server) as doc:
            doc.addParagraph("Hello")
            doc.save("test_odt_export.doc")
            doc.export_docx("test_odt_export.doc")
            doc.export_pdf("test_odt_export.odt")

    def test_ods_calculate_all(libreoffice_server):
        filename="test_ods_calculate_all.ods"
        with ODS_Standard(server=libreoffice_server) as doc:
            #Addrow with one color and style
            doc.addRowWithStyle("A1", [1, 1])
            #Add row with list of colors and styles
            doc.addRowWithStyle("A2", [1, 1], [ColorsNamed.White]*2, ["BoldCenter"]*2)
            #Add row empty
            doc.addRowWithStyle("A3", [])
            doc.addCellWithStyle("C1",  "=A1+B1")
            doc.calculateAll()
            doc.save(filename)
            
        with ODS(filename) as doc:
            r=doc.getValue("C1", detailed=True)
            assert r["value"]=="2"
            
        remove(filename)

    def test_ods_addCell(libreoffice_server):
        filename="test_ods_addCell.pdf"
        with ODS("unogenerator/templates/colored.ods", server=libreoffice_server) as doc:
            doc.addCell("B1", 12.44)
            doc.addCell("G1", date.today())
            doc.addCellWithStyle("B7", 12.44, ColorsNamed.Yellow, "Float6")
            doc.addCellWithStyle("G7", date.today(), ColorsNamed.Yellow, "Date")
            doc.export_pdf(filename)
        remove(filename)
        
    def test_ods_addCellMerged(libreoffice_server):
        filename="test_ods_addCellMerged.pdf"
        with ODS("unogenerator/templates/colored.ods", server=libreoffice_server) as doc:
            doc.addCellMerged("B1:C1", 12.44)
            doc.addCellMerged("G1:H1", date.today())
            doc.addCellMergedWithStyle("B7:B8", 12.44, ColorsNamed.Yellow, "Float6")
            doc.addCellMergedWithStyle("G7:G8", date.today(), ColorsNamed.Yellow, "Date")
            doc.export_pdf(filename)
        remove(filename)

    def test_ods_addListOfRows(libreoffice_server):
        filename="test_ods_addListOfRows.pdf"
        with ODS("unogenerator/templates/colored.ods", server=libreoffice_server) as doc:
            #Rows
            doc.addListOfRows("B1", lor)
            doc.addListOfRows("A1", [])
            #Rows with style
            doc.addListOfRowsWithStyle("B7", lor)
            doc.addListOfRowsWithStyle("A1", [])
            #Columns
            doc.addListOfColumns("H1", lor)
            doc.addListOfColumns("A1", [])
            #Columns with style
            doc.addListOfColumnsWithStyle("H7", lor)
            doc.addListOfColumnsWithStyle("A1", [])
                        
            doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)
            doc.export_pdf(filename)
        remove(filename)
        

    def test_ods_addListOfRows_null_first_row(libreoffice_server):
        """
        Test that styles are correctly guessed even if the first row has None values.
        """
        with ODS_Standard(server=libreoffice_server) as doc:
            # Column A: First row None, second row Decimal -> Should be Float2
            # Column B: First row None, second row int -> Should be Integer
            # Column C: Both rows None -> Should be Default (Normal)
            data = [
                [None, None, None],
                [Decimal("10.5"), 100, None]
            ]
            
            coord_start = Coord("A1")
            doc.addListOfRowsWithStyle(coord_start, data)
            
            # Verify styles
            # A1/A2 should have Float2
            assert doc.sheet.getCellByPosition(0, 0).CellStyle == "Float2"
            assert doc.sheet.getCellByPosition(0, 1).CellStyle == "Float2"
            
            # B1/B2 should have Integer
            assert doc.sheet.getCellByPosition(1, 0).CellStyle == "Integer"
            assert doc.sheet.getCellByPosition(1, 1).CellStyle == "Integer"
            
            # C1/C2 should have Normal (default for ODS_Standard)
            assert doc.sheet.getCellByPosition(2, 0).CellStyle == "Normal"
            assert doc.sheet.getCellByPosition(2, 1).CellStyle == "Normal"


    def test_ods_addFormulaArray(libreoffice_server):
        filename="test_ods_addFormulaArray.pdf"
        with ODS_Standard(server=libreoffice_server) as doc:
           
            # Checks with List of Rows
            doc.addListOfRows("A1", [["=2+2", "=3+3"], ], formulas=False)
            doc.addListOfRows("D1", [["=2+2", "=3+3"], ], formulas=True)
            doc.addListOfRowsWithStyle("A3",  lor_types, formulas=False, styles=lor_types_styles)
            doc.addListOfRowsWithStyle("A6",  lor_types, formulas=True, styles=lor_types_styles)
            
            d3=doc.getValue("D3", detailed=True)
            d6=doc.getValue("D6", detailed=True)
            assert d3["value"]==d6["value"]
            k3=doc.getValue("K3", detailed=True)
            k6=doc.getValue("K6", detailed=True)
            assert k3["value"]!=k6["value"]#0,4
            assert k3["string"]=="=2+2"
            assert k6["string"]=="4"
            
            #Check with addRow
            doc.addRow("A9", ["=2+2", "=3+3"], formulas=False)
            doc.addRow("D9", ["=2+2", "=3+3"], formulas=True)
            doc.addRowWithStyle("A11",  lor_types[0], formulas=False, styles=lor_types_styles)
            doc.addRowWithStyle("A12",  lor_types[0], formulas=True, styles=lor_types_styles)
            
            #Check with addColumn
            doc.addColumn("A14", ["=2+2", "=3+3"], formulas=False)
            doc.addColumn("B14", ["=2+2", "=3+3"], formulas=True)
            doc.addColumnWithStyle("D14",  lor_types[0], formulas=False, styles=lor_types_styles)
            doc.addColumnWithStyle("E14",  lor_types[0], formulas=True, styles=lor_types_styles)
            
            
            # Checks with List of Columns
            doc.addListOfColumns("G14", [["=2+2", "=3+3"], ], formulas=False)
            doc.addListOfColumns("H14", [["=2+2", "=3+3"], ], formulas=True)
            doc.addListOfColumnsWithStyle("J14",  lor_types, formulas=False, styles=lor_types_styles)
            doc.addListOfColumnsWithStyle("M14",  lor_types, formulas=True, styles=lor_types_styles)
            doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)
            
            doc.export_pdf(filename)
        remove(filename)

    def test_ods_addRow(libreoffice_server):
        with ODS("unogenerator/templates/colored.ods", server=libreoffice_server) as doc:
            #Checking range - range_uno conversions
            range_=Range("B2:C3")
            range_uno=range_.uno_range(doc.sheet)
            range_2=Range.from_uno_range(range_uno)
            assert range_==range_2
            #row
            doc.addRow("A1", [])
            doc.addRow("B1", row)
            #row with style
            doc.addRowWithStyle("A1", [])
            doc.addRowWithStyle("B7", row)
            doc.addRowWithStyle("B8", row, ColorsNamed.Yellow, "Integer")
            #Column
            doc.addColumn("A1", [])
            doc.addColumn("H1", row)
            #Column with style
            doc.addColumnWithStyle("A1", [])
            doc.addColumnWithStyle("H7", row)
            doc.addColumnWithStyle("I7", row, ColorsNamed.Yellow, "Integer")
            doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)
            doc.export_pdf("test_ods_addRow.pdf")
            
            # Replace colored cell
            doc.addRowWithStyle("F4", ["Elvis", "Presley"], ColorsNamed.Orange, "BoldCenter")
            doc.addListOfRowsWithStyle("J1", lor_types, ColorsNamed.Orange, None)
            doc.addListOfRows("J4", lor_types)
            doc.export_pdf("test_ods_addRow.pdf")
        remove("test_ods_addRow.pdf")


        
    def test_ods_freezeandselect(libreoffice_server):
        filename="test_ods_freezeandselect.ods"
        with ODS_Standard(server=libreoffice_server) as doc:
            doc.createSheet("Outside range after")
            doc.addCell("A1", "Hola")
            doc.freezeAndSelect("C3")

            doc.createSheet("Outside range before")
            doc.freezeAndSelect("C3")
            doc.addCell("A1", "Hola")

            doc.createSheet("Header")
            doc.freezeAndSelect("A2")
            doc.addCell("A1", "Hola")
            doc.createSheet("Lateral header")
            doc.freezeAndSelect("B1")
            doc.addCell("A1", "Hola")
            
            doc.createSheet("With selected")
            doc.freezeAndSelect("B1", "G6")
            doc.addCell("A1", "Hola")
            
            doc.createSheet("With selected and topleft")
            doc.freezeAndSelect("B2", "G6",  "F5")
            doc.addCell("A1", "Hola")
            
            doc.createSheet("With selected None")
            doc.freezeAndSelect("B2", "G6",  "F5")
            doc.addCell("A1", "Hola")
            
            doc.save(filename)
        remove(filename)
        
    def test_ods_createSheet_nocommonserver():
        # Bad extensions
        with ODS_Standard() as doc:
            with raises(exceptions.UnogeneratorException):
                doc.createSheet("Same name")
                doc.createSheet("Same name")
        assert not path.exists(f"/tmp/unogenerator{doc.server.port}")
                
        with ODS_Standard() as doc:
            doc.createSheet("Same name")
        assert not path.exists(f"/tmp/unogenerator{doc.server.port}")

    def test_ods_export(libreoffice_server):
        # Bad extensions
        with ODS_Standard(server=libreoffice_server) as doc:
            doc.addCellWithStyle("A1", "test_ods_export")
            doc.export_xlsx("test_ods_export.xls")
            doc.export_pdf("test_ods_export.ods")
            doc.save("test_ods_export.xlsx")

    def test_ods_getvalues(libreoffice_server):
        filename="test_get_values.ods"

        number=10
        range_=Range(f"A1:{Coord.from_index(number-1, number-1)}")
        print(range_)

        with ODS_Standard(server=libreoffice_server) as doc:
            doc.createSheet("Get Values")
            lor=[]
            for row in range(number):
                lor_row=[]
                for column in range(number):
                    lor_row.append(str(Coord.from_index(column, row)))
                lor.append(lor_row)
            doc.addListOfRowsWithStyle("A1",  lor)
            doc.save(filename)

        with ODS(filename) as doc:
            assert doc.getValue("A1", detailed=False)=="A1"
            assert doc.getValue("A1", detailed=True)["string"]=="A1"
            
            assert doc.getValueByPosition(1, 1, detailed=False)=="B2"
            assert doc.getValueByPosition(1, 1, detailed=True)["string"]=="B2"
            
            r=doc.getValues(skip_up=0, skip_down=0, skip_left=0, skip_right=0)
            assert len(r)==number
            
            r=doc.getValues(skip_up=0, skip_down=0, skip_left=0, skip_right=0, detailed=True)
            assert len(r)==number
            
            
            r=doc.getValues(skip_up=0, skip_down=0, skip_left=0, skip_right=0, cast=[str]*number)
            assert r[0][0]=="A1"
            
            r=doc.getValuesByRange(range_)
            assert r[0].__class__==tuple
            
            r=doc.getValuesByColumn("A", skip_up=0, skip_down=0)
            assert len(r)==number
        
            r=doc.getValuesByRow("1",  skip_left=0, skip_right=0)
            assert len(r)==number
            
        remove(filename)



    def test_ods_setCellName(libreoffice_server):
        filename="test_ods_setCellName.ods"
        with ODS_Standard(server=libreoffice_server) as doc:
            doc.createSheet("Sheet")
            #Rows
            doc.setCellName("$Sheet.$A$1", "hola")      
            doc.addCellWithStyle("A1", 2, ColorsNamed.White, "Integer")
            doc.addCellWithStyle("A2", "=2*hola", ColorsNamed.White, "Integer")
            values=doc.getValues(detailed=True)
            assert values[1][0]["value"]==4
            doc.save(filename)
        remove(filename)

    def test_ods_toDictionaryOfDetailedValues(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            doc.createSheet("One")
            #Adding values to document
            doc.setCellName("$One.$A$1", "hola")      
            doc.addCellWithStyle("A1", 2, ColorsNamed.White, "Integer")
            doc.addCellWithStyle("A3", "=2*hola", ColorsNamed.White, "Integer")
            
            doc.createSheet("Two")
            doc.addCellWithStyle("A1", 5, ColorsNamed.White, "Integer")
            
            doc.removeSheet(0) #Removing default sheet
            
            #Export to dictionaryOfDetailedValues
            detailed_values=doc.toDictionaryOfDetailedValues()
            
            assert detailed_values["sheets"][1]["name"]=="Two"
            assert detailed_values["sheets"][0]["rows"]==3
            assert detailed_values["dictionary"][("One", "A3")]["is_formula"]==True
            assert detailed_values["dictionary"][("One", "A3")]["value"]==4


    def test_ods_template_with_importlib_resources_files():
        with ODS(files("unogenerator")/"templates/colored.ods") as doc:
            doc.addCell("A1", "Hola")
            doc.export_pdf("test_ods_template_with_importlib_resources_files.pdf")
        remove("test_ods_template_with_importlib_resources_files.pdf")

    def test_ods_setColumnsWidth(libreoffice_server):
        num_columns = 20
        num_rows = 100
        
        # Preparamos una fila de ejemplo con contenido de texto representativo
        row_data = [f"Dato de prueba para columna {i} con cierta longitud" for i in range(num_columns)]
        # Creamos una lista de 100,000 filas. Al usar la misma referencia de lista para cada fila,
        # ahorramos memoria significativamente en el proceso de Python.
        lor = [row_data] * num_rows
        
        # A la primera celda de la segunda columna le añado 1000 caracteres

        lor[0][2]="abcdefghijklmnopqr "*20

        filename = "test_performance_calc.ods"
        with ODS_Standard(server=libreoffice_server) as doc:
            # Medimos la inserción inicial (operación previa necesaria)
            doc.addListOfRows("A1", lor)            
            # Test de rendimiento: OptimalWidth (Ancho automático)
            start_auto = datetime.now()
            # Nota: Al pasar num_columns evitamos que setColumnsWidth llame a getSheetSize()
            doc.setColumnsWidth(lor, types.ColumnsWidthMode.FROM_LOL)
            print(f"Rendimiento setColumnsWidth(): {datetime.now() - start_auto}")
            
            
            doc.save(filename)
        if path.exists(filename):
            remove(filename)



    def test_add_image_invalid_source(libreoffice_server):
        with ODS_Standard(server=libreoffice_server) as doc:
            with raises(ValueError, match="Could not load graphic"):
                doc.addImageToCell("A1", b"invalid_bytes")


