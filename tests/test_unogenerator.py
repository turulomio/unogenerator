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
    with ODS_Standard() as doc:
        #Addrow with one color and style
        doc.addRowWithStyle("A1", [1, 1])
        #Add row with list of colors and styles
        doc.addRowWithStyle("A2", [1, 1], [ColorsNamed.White]*2, ["BoldCenter"]*2)
        #Add row empty
        doc.addRowWithStyle("A3", [])
        doc.addCellWithStyle("C1",  "=A1+B1")
        doc.calculateAll()
        doc.save("test_ods_calculate_all.ods")
        
    with ODS("test_ods_calculate_all.ods") as doc:
        r=doc.getValue("C1", detailed=True)
        assert r["value"]=="2"
        
    remove("calculate_all.ods")

def test_ods_addListOfRows():
    with ODS("unogenerator/templates/colored.ods") as doc:
        doc.addListOfRows("B1", lor)
        doc.addListOfRowsWithStyle("B7", lor)
        doc.addListOfColumns("H1", lor)
        doc.addListOfColumnsWithStyle("H7", lor)
        doc.export_pdf("test_ods_addListOfRows.pdf")
        
    remove("test_ods_addListOfRows.pdf")

def test_ods_addRow():
    with ODS("unogenerator/templates/colored.ods") as doc:
        #Checking range - range_uno conversions
        range_=Range("B2:C3")
        range_uno=range_.uno_range(doc)
        range_2=Range.from_uno_range(range_uno)
        assert range_==range_2
  
        doc.addRow("B1", row)
        doc.addRowWithStyle("B7", row)
#        doc.addColumn("H1", lor)
#        doc.addColumnWithStyle("H7", lor)
        doc.export_pdf("test_ods_addRow.pdf")
        
#    remove("test_ods_addRow.pdf")
