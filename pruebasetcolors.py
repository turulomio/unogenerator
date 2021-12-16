from unogenerator import ODS
from unogenerator.unogenerator import createUnoService
from com.sun.star.beans import PropertyValue
from com.sun.star.sheet.ConditionEntryType import COLORSCALE
from uno import createUnoStruct
doc=ODS("colorscale.ods")

myCells=doc.sheet.getCellRangeByName("A1:A8")
##        print(unorange, dir(unorange))
myConditionalFormats = doc.sheet.ConditionalFormats.getConditionalFormats()
mcf=myConditionalFormats[0]
print("CONDITIONAL FORMATS", myConditionalFormats, dir(myConditionalFormats))
print()
print("CONDITIONAL FORMAT [0]", mcf, dir(mcf))

print(mcf.Count)
for o in mcf:
    print("O", o, dir(o))
    for cse in o.ColorScaleEntries:
        print("cse",cse, dir(cse))
        print(cse.getFormula(), cse.getColor(),cse.getType())


# cse=mcf
# print("CSE", cse, dir(cse))

##print(myCells.ConditionalFormat, dir(myCells.ConditionalFormat),  myCells.ConditionalFormat.getTypes())
##print(myCells.ConditionalFormat.getSupportedServiceNames())
##for e in myCells.ConditionalFormat.createEnumeration():
##    print(e, dir(e))

#myNewCells
#doc.addColumnWithStyle("B1", (1, 2, 4, 7, 5, 3, 6, 1))
#myNewCells=doc.sheet.getCellRangeByName("B1:B8")
#print(myNewCells)

#print("SHEET",  doc.sheet,  dir(doc.sheet))
#print()
#cursor=doc.sheet.createCursorByRange(doc.sheet.getCellRangeByName("A1:A8"))
#print("CURSOR", cursor, dir(cursor))
#print()
##print("CF", cursor.ConditionalFormat,  dir(cursor.ConditionalFormat), cursor.ConditionalFormat.Count)
#print(cursor.getRangeAddress())

# oRangos = doc.document.createInstance("com.sun.star.sheet.SheetCellRanges")
# oRangos.addRangeAddress( doc.sheet.getCellRangeByName( "B1:B8" ).getRangeAddress() ,False )
# doc.document.getCurrentController().select(oRangos)

# print("oRangos", oRangos, dir(oRangos))
# orcf=oRangos.ConditionalFormat
# print("oRangosConditionalFormat", orcf, dir(orcf))

# arr=(
#     PropertyValue('FilterName',0,'writer8',0),
# )


# new=orcf.addNew()
# print("new", new, dir(new))



# cfs=doc.sheet.ConditionalFormats.createByRange(oRangos)
# #cfo=doc.sheet.ConditionalFormats.getConditionalFormats()[0]
# #print("CFO", cfo, dir(cfo), len(cfo))
# #
# #for e in range(cfo.Count):
# #    p=cfo.getByIndex(e)
# #    print("ORIGINAL", p,  dir(p))
# #    cse=p.ColorScaleEntries[0]
# #    print("COLORSCALEENTRIES",  cse ,  dir(cse))
# #    print("COLORSCALEENTRIESDETAILS", cse.getColor() ,  cse.getFormula(), cse.getType(), cse.getTypes())

# cf=doc.sheet.ConditionalFormats.getConditionalFormats()[cfs-1]
# print("CF", cf, dir(cf), len(cf))
#cf.createEntry(1, 0)#1 es colorscale

#entry=cf.getByIndex(0)
#print("ENTRY",  entry, dir(entry))
#
#new = createUnoStruct("com.sun.star.sheet.XColorScaleEntry")
#new.acquire()
##print("NEW", new, dir(new))
##newi=new.__init__()
##print(newi)
#
#
#print("Calorscaleentryies", entry.ColorScaleEntries,  dir(entry.ColorScaleEntries))
#entry.ColorScaleEntries=(new, )
#print("Calorscaleentryies", entry.ColorScaleEntries,  dir(entry.ColorScaleEntries))
#se_uno=createUnoStruct("com.sun.star.sheet.ColorScale")
#print(se_uno,  dir(se_uno), se_uno.__class__)
#print(se_uno.getFormula())
#p=PropertyValue('FilterName',0,'writer8',0),
#se=entry.queryInterface()
#print("SE",  se,  dir(se), se.__class__)

# doc.export_pdf("salida.pdf")
# doc.save("salida.ods")
doc.close()
