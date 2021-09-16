import uno

from os import path


from com.sun.star.beans import PropertyValue
from com.sun.star.text import ControlCharacter


localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',localContext)
ctx = resolver.resolve('uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext')
desktop = ctx.ServiceManager.createInstance('com.sun.star.frame.Desktop')
document=desktop.loadComponentFromURL('private:factory/swriter','_blank',0,())
cursor=document.Text.createTextCursor()

stylefam=document.StyleFamilies
print (stylefam.getElementNames())
styles_page=stylefam.getByName("PageStyles")
styles_paragraph=stylefam.getByName("ParagraphStyles")
print(styles_paragraph.getElementNames())


cursor.setPropertyValue("ParaStyleName", "Heading 1")
document.Text.insertString(cursor, "Text examples", False)
document.Text.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)

document.Text.insertString(cursor, "This is an inserted string.", False)
document.Text.insertString(cursor, "This is another inserted string with absorb.", False)
document.Text.insertString(cursor, "This is another inserted string with new line in the same paragraph\n as you can see", False)
document.Text.insertControlCharacter(cursor, ControlCharacter.PARAGRAPH_BREAK, False)
document.Text.insertString(cursor, "This is another paragraph.", False)

document.Text.insertString(cursor, "This is another paragraph without inserting a control character.\r as you can see ", False)
document.Text.insertString(cursor, "\r\t\tThis is an indented paragraph as you can see ", False)

cursor.setPropertyValue("CharHeight", 20)
document.Text.insertString(cursor, "\rThis is another changing height without styles.", False)


p = PropertyValue()
fileroot=path.abspath("./examples_lowriter")

args=(
    PropertyValue('FilterName',0,'writer8',0),
    PropertyValue('Overwrite',0,True,0),
)
document.storeAsURL(f"file://{fileroot}.odt", args)
print(document.URL)
args=(PropertyValue('FilterName',0,'MS Word 2007 XML',0),
    PropertyValue('Overwrite',0,True,0),)
document.storeToURL(f"file://{fileroot}.docx", args)

args=(PropertyValue('FilterName',0,'writer_pdf_Export',0),
    PropertyValue('Overwrite',0,True,0),)
document.storeToURL(f"file://{fileroot}.pdf", args)

document.dispose()
