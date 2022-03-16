from unogenerator import Range

range=Range("A1:C3")
print(range)

range=range.addRowBefore(-1)
print("Removing row before",  range)


range=range.addRowAfter(-1)
print("Removing row after",  range)
