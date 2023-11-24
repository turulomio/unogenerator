from datetime import date
from os import remove
from unogenerator import ODT_Standard, ODT, ODS_Standard, ODS, ColorsNamed, Range

row=[1, 2, 3, 4, 5]

lor=[]
for i in range(4):
    lor.append(row)



def test_odt_metadata():
    with ODT_Standard() as doc:
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



def test_ods_calculate_all():
    filename="test_ods_calculate_all.ods"
    with ODS_Standard() as doc:
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

def test_ods_addCell():
    filename="test_ods_addCell.pdf"
    with ODS("unogenerator/templates/colored.ods") as doc:
        doc.addCell("B1", 12.44)
        doc.addCell("G1", date.today())
        doc.addCellWithStyle("B7", 12.44, ColorsNamed.Yellow, "Float6")
        doc.addCellWithStyle("G7", date.today(), ColorsNamed.Yellow, "Date")
        doc.export_pdf(filename)
    remove(filename)
    
def test_ods_addCellMerged():
    filename="test_ods_addCellMerged.pdf"
    with ODS("unogenerator/templates/colored.ods") as doc:
        doc.addCellMerged("B1:C1", 12.44)
        doc.addCellMerged("G1:H1", date.today())
        doc.addCellMergedWithStyle("B7:B8", 12.44, ColorsNamed.Yellow, "Float6")
        doc.addCellMergedWithStyle("G7:G8", date.today(), ColorsNamed.Yellow, "Date")
        doc.export_pdf(filename)
    remove(filename)

def test_ods_addListOfRows():
    filename="test_ods_addListOfRows.pdf"
    with ODS("unogenerator/templates/colored.ods") as doc:
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
        
        doc.export_pdf(filename)
    remove(filename)

def test_ods_addRow():
    with ODS("unogenerator/templates/colored.ods") as doc:
        #Checking range - range_uno conversions
        range_=Range("B2:C3")
        range_uno=range_.uno_range(doc.sheet)
        range_2=Range.from_uno_range(range_uno)
        assert range_==range_2
  
        doc.addRow("B1", row)
        doc.addRowWithStyle("B7", row)
        doc.addRowWithStyle("B8", row, ColorsNamed.Yellow, "Integer")
        doc.addColumn("H1", row)
        doc.addColumnWithStyle("H7", row)
        doc.addColumnWithStyle("I7", row, ColorsNamed.Yellow, "Integer")
        doc.export_pdf("test_ods_addRow.pdf")
        
    remove("test_ods_addRow.pdf")
