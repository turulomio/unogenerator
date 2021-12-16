from os import path
from uno import getComponentContext, createUnoStruct, systemPathToFileUrl
from com.sun.star.beans import PropertyValue
localContext = getComponentContext()

resolver = localContext.ServiceManager.createInstance('com.sun.star.bridge.UnoUrlResolver')
ctx = resolver.resolve('uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext')
desktop = ctx.ServiceManager.createInstance('com.sun.star.frame.Desktop')


args=(
    PropertyValue('AsTemplate',0,True,0),
)
template=systemPathToFileUrl(path.abspath("colorscale.ods"))
document=desktop.loadComponentFromURL(template,'_blank', 8, args)

sheet=document.getSheets().getByIndex(0)

# I create and select a range of cells
oRangos = document.createInstance("com.sun.star.sheet.SheetCellRanges")
oRangos.addRangeAddress( sheet.getCellRangeByName( "B1:B8" ).getRangeAddress() ,False )
document.getCurrentController().select(oRangos)
        
#I create a conditional format range in the sheet
cfs=sheet.ConditionalFormats.createByRange(oRangos)
cf=sheet.ConditionalFormats.getConditionalFormats()[cfs-1]

# I Create a colorscale conditional format

cf.createEntry(1, 0)   #1 iS COLORSCALE
print(cf,dir(cf))
entry=cf[0]
print("entry", entry,dir(entry),entry.getPropertySetInfo())
# # pep=entry.queryInterface("com.sun.star.sheet.XColorScaleEntry")
# # print(pep,dir(pep))
cse=entry.ColorScaleEntries
print("CSENTRIES",cse,dir(cse),cse.__class__)


from com.sun.star.beans import PropertyValue

from com.sun.star.sheet import XColorScaleEntry
celladdress= createUnoStruct("com.sun.star.table.CellAddress")
celladdress= createUnoStruct("com.sun.star.sheet.XColorScaleEntry")
print("CELLADDRESS", celladdress,dir(celladdress))
a=XColorScaleEntry()
print("A",a,dir(a))
a.setType(1)


# colorentry=list()

# entry= list()
# entry.append(PropertyValue("Formula",0,0,0))
# entry.append(PropertyValue("Color",0,16711680,0))
# entry.append(PropertyValue("Type",0,0,0))
# colorentry.append(tuple(entry))
# entry= list()
# entry.append(PropertyValue("Formula",0,50,0))
# entry.append(PropertyValue("Color",0,16777215,0))
# entry.append(PropertyValue("Type",0,2,0))
# colorentry.append(tuple(entry))
# entry= list()
# entry.append(PropertyValue("Formula",0,0,0))
# entry.append(PropertyValue("Color",0,43315,0))
# entry.append(PropertyValue("Type",0,1,0))
# colorentry.append(tuple(entry))

# colorentry=tuple(colorentry)
# cse=colorentry

# for o in cse:
#     print("cse",o, dir(o))


# I don't know how to follow
#new = createUnoStruct()


save_args=(
    PropertyValue('FilterName', 0, 'calc8', 0),
)
document.storeAsURL(systemPathToFileUrl(path.abspath("output.ods")), (save_args))
document.dispose()
