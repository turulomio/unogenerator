from datetime import datetime
from unogenerator import ODS_Standard, ODS, Coord, Range

number=1000
range_=Range(f"A1:{Coord.from_index(number-1, number-1)}")
print(range_)

with ODS_Standard() as doc:
    doc.createSheet("Get Values")
    lor=[]
    for row in range(number):
        lor_row=[]
        for column in range(number):
            lor_row.append(str(Coord.from_index(column, row)))
        lor.append(lor_row)
    doc.addListOfRowsWithStyle("A1",  lor)
    doc.save("get_values.ods")

with ODS("get_values.ods") as doc:
    start=datetime.now()
    r=doc.getValue("A1", detailed=True)
    print("getValue()", 1,   datetime.now()-start)
    
    start=datetime.now()
    r=doc.getValueByPosition(0, 0)
    print("getValueByPosition()", 1,   datetime.now()-start)
    
    start=datetime.now()
    r=doc.getValues(skip_up=0, skip_down=0, skip_left=0, skip_right=0, cast=[str]*10, detailed=True)
    print("getValues()", number*number,   datetime.now()-start)
    
    start=datetime.now()
    r=doc.getValuesByRange(range_)
    print("getValuesByRange()", number*number,   datetime.now()-start)
    
    start=datetime.now()
    r=doc.getValuesByColumn("A", skip_up=0, skip_down=0)
    print("getValuesByColumn()", number,   datetime.now()-start)
#
    start=datetime.now()
    r=doc.getValuesByRow("1",  skip_left=0, skip_right=0)
    print("getValuesByRow()", number,   datetime.now()-start)
