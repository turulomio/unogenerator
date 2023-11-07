from os import remove
from unogenerator import ODT_Standard, ODT, ODS_Standard, ODS

class TestODT():
    def test_metadata(self):
        with ODT_Standard() as doc:
            doc.setMetadata(
                title="Metadata example",
                subject="UnoGenerator testing",
                author="Turulomio",
                description="This testing works with metadata methods",
                keywords=["Metadata", "Testing"],
            )
            doc.save("delete_metadata.odt")

        with ODT("delete_metadata.odt") as doc:
            assert doc.getMetadata()["Title"]== "Metadata example"
            
        remove("delete_metadata.odt")



class TestODS():
    def test_calculate_all(self):
        with ODS_Standard() as doc:
            doc.addRowWithStyle("A1", [1, 1])
            doc.addCellWithStyle("C1",  "=A1+B1")
            doc.calculateAll()
            doc.save("calculate_all.ods")
            
        with ODS("calculate_all.ods") as doc:
            r=doc.getValue("C1", detailed=True)
            assert r["value"]=="2"
            
        remove("calculate_all.ods")
