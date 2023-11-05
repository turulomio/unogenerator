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
    doc.addRowWithStyle("A1", row, ColorsNamed.Orange, "BoldCenter", cellbycell=True)
    print(f"UnoGenerator creates {number} cells (cellbycell) in {datetime.now()-start}")
    doc.save("addrowwithstyle_unogenerator_guessing_styles.ods")

start=datetime.now()
with ODS_Standard() as doc:
    doc.createSheet("Benchmark")
    doc.addRowWithStyle("A1", row, ColorsNamed.Orange, "BoldCenter")
    print(f"UnoGenerator creates {number} cells in {datetime.now()-start}")
    doc.save("addrowwithstyle_unogenerator_using_massive.ods")
