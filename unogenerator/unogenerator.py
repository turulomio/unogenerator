## @namespace unogenerator.unogenerator
## @brief Package that allows to read and write Libreoffice ods and odt files

from datetime import datetime
from os import path
from uno import getComponentContext

from com.sun.star.beans import PropertyValue
from com.sun.star.text import ControlCharacter
from com.sun.star.awt import Size
from com.sun.star.style.BreakType import PAGE_AFTER
#from unogenerator.reusing.casts import object2value

class ODF:
    def __init__(self, filename,  template=None, loserver_port=2002):
        self.filename=f"file://{path.abspath(filename)}"
        self.template=None if template is None else f"file://{path.abspath(template)}"
        self.init=datetime.now()
        
        localContext = getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',localContext)
        ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={loserver_port};urp;StarOffice.ComponentContext')
        self.desktop = ctx.ServiceManager.createInstance('com.sun.star.frame.Desktop')

    def close(self):
        self.document.dispose()
        
    def print_styles(self):
        stylefam=self.document.StyleFamilies
        
        print(f"Document '{self.filename}' styles")
        for sf_index,  sf in enumerate(stylefam):
            print (f"  * Family '{stylefam.getElementNames()[sf_index]}'")
            styles=list(sf.getElementNames())
            styles.sort()
            for style in styles:
                print ( f"    - {style}")
                    
class ODT(ODF):
    def __init__(self, filename, template=None):
        ODF.__init__(self, filename, template)
        if self.template is None:
            self.document=self.desktop.loadComponentFromURL('private:factory/swriter','_blank',0,())
        else:
            self.document=self.desktop.loadComponentFromURL(self.template,'_blank',0,())
        self.cursor=self.document.Text.createTextCursor()

        
    def save(self):
        ## SAVE FILE
        args=(
            PropertyValue('FilterName',0,'writer8',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeAsURL(self.filename, args)
        
                
    def export_pdf(self):
        args=(
            PropertyValue('FilterName',0,'writer_pdf_Export',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeToURL(self.filename[:-4]+".pdf", args)
        
        
    def export_docx(self):
        args=(
            PropertyValue('FilterName',0,'MS Word 2007 XML',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeToURL(self.filename[:-4]+".docx", args)
        
    def addParagraph(self,  text,  style):
        self.cursor.setPropertyValue("ParaStyleName", style)
        self.document.Text.insertString(self.cursor, text, False)
        self.document.Text.insertControlCharacter(self.cursor, ControlCharacter.PARAGRAPH_BREAK, False)
        
    ## @param list Pueden ser str, o textcontent objects
    def addParagraphComplex(self,  list,  style):
        self.cursor.setPropertyValue("ParaStyleName", style)
        for o in list:
            if o.__class__==str:
                self.document.Text.insertString(self.cursor, o, False)
            elif o.ImplementationName=="SwXTextGraphicObject":## Es un pyuno object raro no mostraba bvien
                self.document.Text.insertTextContent(self.cursor, o, False)               
        self.document.Text.insertControlCharacter(self.cursor, ControlCharacter.PARAGRAPH_BREAK, False)
        
    ## Returns a text content that can be inserted with document.Text.insertTextContent(cursor,image, False)
    ## @param anchortype AS_CHARACTER, AT_PARAGRAPH
    ## @param name None if we want to use lo default name
    def textcontentImage(self, filename, width=2000,  height=2000, anchortype="AS_CHARACTER", name=None, ):
        image=self.document.createInstance("com.sun.star.text.TextGraphicObject")
        image.AnchorType=anchortype
        if name is not None:
            image.setName(name)
        image.GraphicURL=f"file://{path.abspath(filename)}"
        image.Size=Size(width, height)
        return image
        
    def addImageParagraph(self, filename_list, width=2000,  height=2000, name=None, style="Illustration"):
        self.cursor.setPropertyValue("ParaStyleName", style)
        for filename in filename_list:
            self.document.Text.insertTextContent(self.cursor,self.textcontentImage(filename, width, height, "AS_CHARACTER"), False)
        self.document.Text.insertControlCharacter(self.cursor, ControlCharacter.PARAGRAPH_BREAK, False)
        
##    table.setPropertyValue( "BackTransparent", uno.Bool(0) )
##    table.setPropertyValue( "BackColor", 13421823 )
##    row = rows.getByIndex(0)
##    row.setPropertyValue( "BackTransparent", uno.Bool(0) )
##    row.setPropertyValue( "BackColor", 6710932 )
##    table.getCellByName("C4").setValue(415.7)
##    table.getCellByName("D4").setFormula("sum <A4:C4>")
 
##def insertTextIntoCell( table, cellName, text, color ):
##    tableText = table.getCellByName( cellName )
##    cursor = tableText.createTextCursor()
##    cursor.setPropertyValue( "CharColor", color )
##    tableText.setString( text )
    def addTableParagraph(self, data, paragraph_style="Standard", table_style="Elegant"):
        ## Table by cellname
        num_rows=len(data)
        if num_rows==0:
            return
        num_columns=len(data[0])
        self.cursor.setPropertyValue("ParaStyleName", paragraph_style)
        table=self.document.createInstance("com.sun.star.text.TextTable")
#        table.setPropertyValue("TableStyleName", table_style)
        table.initialize(num_rows, num_columns)
        self.document.Text.insertTextContent(self.cursor,table,False)
        for row, row_data in enumerate(data):
            for column, cell_data in enumerate(row_data):
                cell=table.getCellByPosition(row, column)
                cell.setString(str(cell_data))
        #cell.insertTextContent(cell,image,False)
        self.document.Text.insertControlCharacter(self.cursor, ControlCharacter.PARAGRAPH_BREAK, False)
        
    def addListPlain(self, arr, list_style="List_2", paragraph_style="Puntitos"):
#        def get_items(list_o, list_style, paragraph_style):
#            r=[]
#            for o in list_o:
#                it=ListItem()
#                if o.__class__==str:
#                    it.addElement(P(stylename=paragraph_style, text=o))
#                else:
#                    it.addElement(get_list(o, list_style, paragraph_style))
#                r.append(it)
#            return r
#        def get_list(arr, list_style, paragraph_style):
#            ls=List(stylename=list_style)
#            for listitem in get_items(arr, list_style, paragraph_style):
#                ls.addElement(listitem)
#            return ls
#        # #########################
        for s in arr:
            self.addParagraph(s, paragraph_style)
        
    def pageBreak(self, style="Standard"):
        self.cursor.setPropertyValue("ParaStyleName", style)
        self.document.Text.insertString(self.cursor, "", False)
        self.cursor.BreakType = PAGE_AFTER
        self.document.Text.insertControlCharacter(self.cursor.End, ControlCharacter.PARAGRAPH_BREAK, False)

class ODS(ODF):
    def __init__(self, filename, template=None):
        ODF.__init__(self, filename, template)
        
        if self.template is None:
            self.document=self.desktop.loadComponentFromURL('private:factory/scalc','_blank',0,())
        else:
            self.document=self.desktop.loadComponentFromURL(self.template,'_blank',0,())

    def save(self):
        ## SAVE FILE
        args=(
            PropertyValue('FilterName',0,'calc8',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeAsURL(self.filename, args)

    def export_pdf(self):
        args=(
            PropertyValue('FilterName',0,'calc_pdf_Export',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeToURL(self.filename[:-4]+".pdf", args)

    def export_xlsx(self):
        args=(
            PropertyValue('FilterName',0,'Calc MS Excel 2007 XML',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeToURL(self.filename[:-4]+".xlsx", args)
