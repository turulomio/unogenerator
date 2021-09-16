import uno

from os import path


from com.sun.star.beans import PropertyValue
localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',localContext)
ctx = resolver.resolve('uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext')
desktop = ctx.ServiceManager.createInstance('com.sun.star.frame.Desktop')
document=desktop.loadComponentFromURL('private:factory/swriter','_blank',0,())
cursor=document.Text.createTextCursor()
document.Text.insertString(cursor, "This is a probe.",0)

p = PropertyValue()
fileroot=path.abspath("./examples_lowriter")
print(fileroot)
args=(
    PropertyValue('FilterName',0,'writer8',0),
    PropertyValue('Overwrite',0,True,0),
)
document.storeAsURL(f"file://{fileroot}.odt", args)
print(document.URL)
args=(PropertyValue('FilterName',0,'MS Word 2007 XML',0),)
document.storeToURL(f"file://{fileroot}.docx", args)

args=(PropertyValue('FilterName',0,'writer_pdf_Export',0),)
document.storeToURL(f"file://{fileroot}.pdf", args)

document.dispose()
