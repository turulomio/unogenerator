from unogenerator import ODS_Standard

def test_debug():
    with ODS_Standard() as doc:
        doc.addCellWithStyle("A1", "Hola")
        doc.statistics.debug()
