from unogenerator import ODS, ODS_Standard, ODT_Standard

with ODS() as doc:
    doc.createSheet("With")
    doc.addCellWithStyle("A1", "HOLA WITH")
    doc.save("borrar.ods")
    
with ODS_Standard(2003) as doc:
    doc.createSheet("With")
    doc.addCellMergedWithStyle("A1:C3", "HOLA WITH")
    doc.export_xlsx("borrar.xlsx")
    
with ODT_Standard() as doc:
    doc.addParagraph("Hello world","Standard")
    doc.export_pdf("borrar.pdf")