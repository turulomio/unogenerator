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

row=[]
for i in range(number):
    row.append(i)


with ODS_Standard() as doc:
    doc.createSheet("Row with list of colors and styles")
    start=datetime.now()
    doc.addRowWithStyle("A1", row , [ColorsNamed.Orange, ColorsNamed.Yellow, ColorsNamed.Blue, ColorsNamed.Green]*int(number/4), ["BoldCenter", "Integer"]*int(number/2))
    print(f"Row with {number } different colors and styles in {datetime.now()-start}")
    

    doc.createSheet("Row with one color and style")
    start=datetime.now()
    doc.addRowWithStyle("A1", row, ColorsNamed.Orange, "BoldCenter")
    print(f"Row with {number} cells with one color and style in {datetime.now()-start}")

    doc.createSheet("Column with one color and style")
    start=datetime.now()
    doc.addColumnWithStyle("A1", row, ColorsNamed.Orange, "BoldCenter")
    print(f"Column with {number} cells with one color and style in {datetime.now()-start}")


    doc.createSheet("List of rows")
    start=datetime.now()
    doc.addListOfRowsWithStyle(Coord("A2"), lor)
    print(f"List of {number} rows in {datetime.now()-start}")

    doc.createSheet("List of columns")
    start=datetime.now()
    doc.addListOfColumnsWithStyle(Coord("A2"), lor)
    print(f"List of {number} columns in {datetime.now()-start}")
