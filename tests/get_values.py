from unogenerator import ODS_Standard, ODS

with ODS_Standard() as doc:
    doc.createSheet("Get Values")
    doc.addListOfRowsWithStyle("A1",  [[1, 2, 3], [4, 5, 6], [7, 8, 9]])
    doc.save("get_values.ods")

with ODS("get_values.ods") as doc:
    print("All",  doc.getValues())
    print("Skip up 1",  doc.getValues(skip_up=1))
    print("Skip down 1",  doc.getValues(skip_down=1))
