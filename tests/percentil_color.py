from unogenerator import ODS_Standard
from unogenerator.commons import Coord as C, addDebugSystem

addDebugSystem("DEBUG")

with ODS_Standard() as doc:
    doc.createSheet("Get Values")
    doc.addColumnWithStyle("A1",  [1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
    # Aplicamos la escala de color dinámica en el rango A1:A10
    doc.setColorScale("A1:A10")

    # doc.createSheet("Get Values2")
    # doc.addColumnWithStyle("A1",  range(100))
    # # doc.setColorScale("A1:A100")

    doc.export_pdf("percentil_color.pdf")
    doc.save("percentil_color.ods")
