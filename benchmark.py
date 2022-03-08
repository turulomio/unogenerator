from datetime import datetime
from os import path
from unogenerator import ODS_Standard
from unogenerator.commons import Coord, ColorsNamed
from sys import exit

number=5000

start=datetime.now()
doc=ODS_Standard()
doc.createSheet("Benchmark")
doc.addCellWithStyle("A1", "N", ColorsNamed.Orange, "BoldCenter")
for i in range(number):
    doc.addCellWithStyle(Coord("A2").addRow(i), i)
print(f"Unogenerator creates {number} cells guessing styles in {datetime.now()-start}")
doc.save("benchmark_unogenerator_guessing_styles.ods")
doc.close()


start=datetime.now()
doc=ODS_Standard()
doc.createSheet("Benchmark")
doc.addCellWithStyle("A1", "N", ColorsNamed.Orange, "BoldCenter")
for i in range(number):
    doc.addCellWithStyle(Coord("A2").addRow(i), i, ColorsNamed.White, "Integer")
print(f"Unogenerator creates {number} cells defining styles in {datetime.now()-start}")
doc.save("benchmark_unogenerator_defining_styles.ods")
doc.close()