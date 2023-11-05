from datetime import datetime
from unogenerator import ODS_Standard
from unogenerator.commons import ColorsNamed


number=1000
row=[]
for i in range(number):
    row.append(f"Column {i}")


start=datetime.now()
with ODS_Standard() as doc:
    doc.createSheet("Benchmark")
    doc.addRowWithStyle("A1", [1, 2, 3], [ColorsNamed.Orange, ColorsNamed.Yellow, ColorsNamed.Blue], ["BoldCenter", "Normal", "Integer"], cellbycell=False)
    print(f"UnoGenerator creates {number} cells (cellbycell) in {datetime.now()-start}")
    doc.save("addrowwithstyle_normal.ods")
    
start=datetime.now()
with ODS_Standard() as doc:
    doc.createSheet("Benchmark")
    doc.addRowWithStyle("A1", row, ColorsNamed.Orange, "BoldCenter", cellbycell=True)
    print(f"UnoGenerator creates {number} cells (cellbycell) in {datetime.now()-start}")
    doc.save("addrowwithstyle_unogenerator_guessing_styles.ods")

start=datetime.now()
with ODS_Standard() as doc:
    doc.createSheet("Benchmark")
    doc.addRowWithStyle("A1", row, ColorsNamed.Orange, "BoldCenter")
    print(f"UnoGenerator creates {number} cells in {datetime.now()-start}")
    doc.save("addrowwithstyle_unogenerator_using_massive.ods")

start=datetime.now()
with ODS_Standard() as doc:
    doc.createSheet("Benchmark")
    doc.addColumnWithStyle("A1", row, ColorsNamed.Orange, "BoldCenter", cellbycell=True)
    print(f"UnoGenerator creates {number} cells (cellbycell) in {datetime.now()-start}")
    doc.save("addcolumnwithstyle_unogenerator_guessing_styles.ods")

start=datetime.now()
with ODS_Standard() as doc:
    doc.createSheet("Benchmark")
    doc.addColumnWithStyle("A1", row, ColorsNamed.Orange, "BoldCenter")
    print(f"UnoGenerator creates {number} cells in {datetime.now()-start}")
    doc.save("addcolumnwithstyle_unogenerator_using_massive.ods")
