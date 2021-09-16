import uno

from os import path


from com.sun.star.beans import PropertyValue
localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',localContext)
ctx = resolver.resolve('uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext')
desktop = ctx.ServiceManager.createInstance('com.sun.star.frame.Desktop')
wb=desktop.loadComponentFromURL('private:factory/scalc','_blank',0,())
sheet=wb.getSheets().getByIndex(0)
sheet.setName('Prueba')


sheet.getCellByPosition(0,0).setString('Idioma')
sheet.getCellByPosition(1,0).setString('Cadena a traducir')
sheet.getCellByPosition(1,1).setString("Hola")
p = PropertyValue()
fileroot=path.abspath("./examples_localc")

p.Name = "ReadOnly"
p.Value= "False"
wb.storeToURL(f"file://{fileroot}.ods", (p,))

args=(PropertyValue('FilterName',0,'Calc MS Excel 2007 XML',0),)
wb.storeToURL(f"file://{fileroot}.xlsx", args)

args=(PropertyValue('FilterName',0,'calc_pdf_Export',0),)
wb.storeToURL(f"file://{fileroot}.pdf", args)

wb.dispose()