from os import path
from uno import getComponentContext, systemPathToFileUrl
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

# I don't know how to follow
#new = createUnoStruct("com.sun.star.sheet.XColorScaleEntry")


save_args=(
    PropertyValue('FilterName', 0, 'calc8', 0),
)
document.storeAsURL(systemPathToFileUrl(path.abspath("output.ods")), (save_args))
document.dispose()
