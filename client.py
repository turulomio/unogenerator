import uno

from com.sun.star.beans import PropertyValue
localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',localContext)
ctx = resolver.resolve('uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
desktop = ctx.ServiceManager.createInstance('com.sun.star.frame.Desktop')
wb=desktop.loadComponentFromURL('private:factory/scalc','_blank',0,())
sheet=wb.getSheets().getByIndex(0)
sheet.setName('Prueba')


sheet.getCellByPosition(0,0).setString('Idioma')
sheet.getCellByPosition(1,0).setString('Cadena a traducir')
sheet.getCellByPosition(1,1).setString("Hola")
p = PropertyValue()
p.Name = "ReadOnly"
p.Value= "False"
args=(PropertyValue('FilterName',0,'MS Word 2007 XML',0),)
print(dir(wb))
sUrl=uno.systemPathToFileUrl('/home/keko/Proyectos/unogenerator/prueba.ods')
wb.storeAsURL(sUrl, (p,))

