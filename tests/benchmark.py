from datetime import datetime, timedelta, date
from decimal import Decimal
from unogenerator import ODS_Standard
from unogenerator.commons import Coord, ColorsNamed
from unogenerator.reusing.percentage import Percentage
from unogenerator.reusing.currency import Currency


number=2000
lor=[]
for i in range(number):
    lor.append([
        i, 
        Coord("B2").addRow(i).string(), 
        date.today()+timedelta(days=i) , 
        datetime.now()+timedelta(hours=i), 
        1.123*i, 
        Decimal(1.123*i), 
        Percentage(1, 1.123*i), 
        Currency(1.123*i, "EUR"), 
        bool(i % 2), 
    ])
headers=["N", "String", "Date", "Datetime", "Float", "Decimal", "Percentage", "Currency", "Boolean"]


start=datetime.now()
doc=ODS_Standard()
doc.createSheet("Benchmark")
doc.setColumnsWidth([4]*20)
doc.addRowWithStyle("A1", headers, ColorsNamed.Orange, "BoldCenter")
doc.addListOfRowsWithStyle(Coord("A2"), lor, cellbycell=True)
print(f"UnoGenerator creates {number} cells (cellbycell) in {datetime.now()-start}")
doc.save("benchmark_unogenerator_guessing_styles.ods")
doc.close()

start=datetime.now()
doc=ODS_Standard()
doc.createSheet("Benchmark")
doc.setColumnsWidth([4]*20)
doc.addRowWithStyle("A1", headers, ColorsNamed.Orange, "BoldCenter")

doc.addListOfRowsWithStyle(Coord("A2"), lor)
print(f"UnoGenerator creates {number} cells in {datetime.now()-start}")
doc.save("benchmark_unogenerator_using_massive.ods")
doc.close()
