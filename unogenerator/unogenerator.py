## @namespace unogenerator.unogenerator
## @brief Package that allows to read and write Libreoffice ods and odt files

from datetime import datetime
from os import path
from uno import getComponentContext, createUnoStruct

from com.sun.star.beans import PropertyValue
from com.sun.star.text import ControlCharacter
from com.sun.star.awt import Size
from com.sun.star.style.BreakType import PAGE_AFTER
from pkg_resources import resource_filename
from unogenerator.commons import Coord as C, Colors,  Range as R, datetime2uno, row2index, column2index
from unogenerator.reusing.currency import Currency
from unogenerator.reusing.datetime_functions import string2dtnaive, string2date
from unogenerator.reusing.percentage import Percentage

class ODF:
    def __init__(self, template=None, loserver_port=2002):
        self.template=None if template is None else f"file://{path.abspath(template)}"
        self.init=datetime.now()
        
        localContext = getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',localContext)
        ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={loserver_port};urp;StarOffice.ComponentContext')
        self.desktop = ctx.ServiceManager.createInstance('com.sun.star.frame.Desktop')

    def calculateAll(self):
        self.document.calculateAll()

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

    def setLanguage(self, language, country):
        print("FALTA")
        self.language="es"
        self.country="ES"

    def setMetadata(self, title="",  subject="", creator="", description="", keywords=[], creationdate=datetime.now()):
        self.document.DocumentProperties.Author=creator
        self.document.DocumentProperties.Generator=creator
        self.document.DocumentProperties.Description=description
        self.document.DocumentProperties.Subject=subject
        self.document.DocumentProperties.Keywords=keywords
        self.document.DocumentProperties.CreationDate=datetime2uno(creationdate)
        self.document.DocumentProperties.ModificationDate=datetime2uno(creationdate)
        self.document.DocumentProperties.Title=title
                    
class ODT(ODF):
    def __init__(self, template=None, loserver_port=2002):
        ODF.__init__(self, template, loserver_port)
        if self.template is None:
            self.document=self.desktop.loadComponentFromURL('private:factory/swriter','_blank',0,())
        else:
            self.document=self.desktop.loadComponentFromURL(self.template,'_blank',0,())
        self.cursor=self.document.Text.createTextCursor()

        
    def save(self, filename):
        filename=f"file://{path.abspath(filename)}"
        ## SAVE FILE
        args=(
            PropertyValue('FilterName',0,'writer8',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeAsURL(filename, args)
        
                
    def export_pdf(self, filename):
        filename=f"file://{path.abspath(filename)}"
        args=(
            PropertyValue('FilterName',0,'writer_pdf_Export',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeToURL(filename, args)
        
    def find_and_remove_and_setcursorposition(self, find):
        search=self.document.createSearchDescriptor()
        search.SearchString=find
        found=self.document.findFirst(search)
        found.String.replace(find,  "")

    def export_docx(self, filename):
        filename=f"file://{path.abspath(filename)}"
        args=(
            PropertyValue('FilterName',0,'MS Word 2007 XML',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeToURL(filename, args)
        
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
    def __init__(self, template=None, loserver_port=2002):
        ODF.__init__(self, template, loserver_port)
        
        if self.template is None:
            self.document=self.desktop.loadComponentFromURL('private:factory/scalc','_blank',0,())
        else:
            self.document=self.desktop.loadComponentFromURL(self.template,'_blank',0,())
        self.sheet_index=0
        self.sheet=self.setActiveSheet(self.sheet_index)
            
    def createSheet(self, name, index):
        self.document.getSheets().insertNewByName(name, index)
        self.setActiveSheet(index)
        
    def removeSheet(self, index):
        current=self.sheet.Name
        self.setActiveSheet(index)
        self.document.getSheets().removeByName(self.sheet.Name)
        self.document.getSheets().getByName(current)

    def setActiveSheet(self,  index):
        self.sheet_index=index
        self.sheet=self.document.getSheets().getByIndex(index)
    
    def setColumnsWidth(self, l):
        columns=self.sheet.getColumns()
        for i, width in enumerate(l):
            column=columns.getByIndex(i)
            column.Width=l[i]
            
    def setComment(self, coord, comment):
        coord=C.assertCoord(coord)
#        print(cell.supportsService("com.sun.star.sheet.XSheetAnnotation"))
        celladdress= createUnoStruct("com.sun.star.table.CellAddress")
        celladdress.Sheet=self.sheet_index
        celladdress.Column=coord.letterIndex()
        celladdress.Row=coord.numberIndex()
        self.sheet.Annotations.insertNew(celladdress, comment)
            
    def addCell(self, coord, o, color_dict=Colors["White"], outlined=1, alignment="left", decimals=2, bold=False):
        coord=C.assertCoord(coord)
        cell=self.sheet.getCellByPosition(coord.letterIndex(), coord.numberIndex())
        self.__set_value_to_cell(cell, o)
            
        border_prop = createUnoStruct("com.sun.star.table.BorderLine2")
        border_prop.LineWidth = outlined
        cell.setPropertyValue("TopBorder", border_prop)
        cell.setPropertyValue("LeftBorder", border_prop)
        cell.setPropertyValue("RightBorder", border_prop)
        cell.setPropertyValue("BottomBorder", border_prop)
        cell.setPropertyValue("CellBackColor", color_dict["color"])

    def addCellWithStyle(self, coord, o, color_dict=Colors["White"], style=None):
        coord=C.assertCoord(coord)
        cell=self.sheet.getCellByPosition(coord.letterIndex(), coord.numberIndex())
        self.__set_value_to_cell(cell, o)
        cell.setPropertyValue("CellStyle", style)
        cell.setPropertyValue("CellBackColor", color_dict["color"])
        
    def __set_value_to_cell(self, cell, o):
        if o.__class__.__name__  == "datetime":
            cell.setString(str(o))
        elif o.__class__.__name__ in ("date", "timedelta" , "time"):
            cell.setString(str(o))
        elif o.__class__.__name__ in ("Percentage", "Money"):
            cell.setValue(float(o.value))
        elif o.__class__.__name__ in ("Currency", "Money"):
            cell.setValue(float(o.amount))
        elif o.__class__.__name__ in ("Decimal",  "float"):
            cell.setValue(float(o))
        elif o.__class__.__name__ in ("int", ):
            cell.setValue(int(o))
        elif o.__class__.__name__ in ("str", ):
            if o.startswith("=") or o.startswith("+"):
                cell.setFormula(o)
            else:
                cell.setString(o)
        elif o.__class__.__name__ in ("bool", ):
            cell.setValue(int(o))
        else:
            cell.setString(str(o))
            print("MISSING", o.__class__.__name__)
        
    def addCellMerged(self, range, o, color_dict=Colors["White"], style=None):
        range=R.assertRange(range)
        cell=self.sheet.getCellByPosition(range.start.letterIndex(), range.start.numberIndex())
        cellrange=self.sheet.getCellRangeByName(range.string())
        cellrange.merge(True)
        self.__set_value_to_cell(cell, o)
        cell.setPropertyValue("CellStyle", style)
        cell.setPropertyValue("CellBackColor", color_dict["color"])

    def freezeAndSelect(self, freeze, current=None, topleft=None):
        freeze=C.assertCoord(freeze) 
        current=None if current is None else C.assertCoord(current)
        topleft=None if topleft is None else C.assertCoord(topleft)
        self.document.getCurrentController().setActiveSheet(self.sheet)
        self.document.getCurrentController().freezeAtPosition(freeze.letterIndex(), freeze.numberIndex())

        if current is not None:
            currentcell=self.sheet.getCellByPosition(current.letterIndex(), current.numberIndex())
            self.document.getCurrentController().select(currentcell)
        
        if topleft is not None:
            self.document.getCurrentController().setFirstVisibleColumn(topleft.letterIndex())
            self.document.getCurrentController().setFirstVisibleRow(topleft.numberIndex())
        
    def getValue(self, coord, standard=True):
        coord=C.assertCoord(coord)
        return self.getValueByPosition(coord.letterIndex(), coord.numberIndex(), standard)
        
    def getValueByPosition(self, letter_index, number_index, standard=True):
        return self.__cell_to_object(self.sheet.getCellByPosition(letter_index, number_index), standard)


    ## Returns a list of rows with the values of the sheet
    ## @param sheet_index Integer index of the sheet
    ## @param skip_up int. Number of rows to skip at the begining of the list of rows (lor)
    ## @param skip_down int. Number of rows to skip at the end of the list of rows (lor)
    ## @return Returns a list of rows of object values
    def getValues(self, skip_up=0, skip_down=0, standard=True):
        if skip_up>0 or skip_down>0:
            print("FALTA, REMOVE FROM RANGE")
        return self.getValuesByRange(self.getSheetRange(), standard)

    ## @param sheet_index Integer index of the sheet
    ## @param range_ Range object to get values. If None returns all values from sheet
    ## @return Returns a list of rows of object values
    def getValuesByRange(self, range_, standard=True):
        r=[]
        for row_indexes in range_.indexes_list():
            rrow=[]
            for column_index,  row_index in row_indexes:
                rrow.append(self.getValueByPosition(column_index, row_index, standard))
            r.append(rrow)
        return r
    
    ## @param sheet_index Integer index of the sheet
    ## @param column_letter Letter of the column to get values
    ## @param skip Integer Number of top rows to skip in the result
    ## @return List of values
    def getValuesByColumn(self, column_letter, skip_up=0, skip_down=0, standard=True):
        r=[]
        for row in range(skip_up, self.rowNumber()-skip_down):
            r.append(self.getValueByPosition(column2index(column_letter), row, standard))
        return r    

    ## @param sheet_index Integer index of the sheet
    ## @param row_number String Number of the row to get values
    ## @param skip Integer Number of top rows to skip in the result
    ## @return List of values
    def getValuesByRow(self, row_number, skip_left=0, skip_right=0, standard=True):
        r=[]
        for column in range(skip_left, self.columnNumber()-skip_right):
            r.append(self.getValueByPosition(column, row2index(row_number), standard))
        return r    


    ## @standard Bool Uses standard templates styles if true
    def __cell_to_object(self, cell, standard=True):
        if standard is True:
            if cell.CellStyle=="Bool":
                return bool(cell.getValue())
            elif cell.CellStyle in ["Datetime"]:
                return string2dtnaive(cell.getString(),"%Y-%m-%d %H:%M:%S.")
            elif cell.CellStyle in ["Date"]:
                return string2date(cell.getString())
            elif cell.CellStyle in ["Float2", "Float6"]:
                return cell.getValue()
            elif cell.CellStyle in ["Integer"]:
                return int(cell.getValue())
            elif cell.CellStyle in ["EUR", "USD"]:
                return Currency(cell.getValue(), cell.CellStyle)
            elif cell.CellStyle in ["Percentage"]:
                return Percentage(cell.getValue(), 1)
#            else:
                #print("Missing", cell.CellStyle)
            return cell.getString()
        else: #standad is false
            return cell.getString()

    ## Return a Range object with the limits of the index sheet
    def getSheetRange(self):
        endcoord=C("A1").addRow(self.rowNumber()-1).addColumn(self.columnNumber()-1)
        return R("A1:"+endcoord.string())
        
        
    ## Esta función puede que sea costosa
    def rowNumber(self):
        return len(self.sheet.getData())
        
    ## Esta función puede que sea costosa
    def columnNumber(self):
        if len(self.sheet.getData())==0:
            return 0
        else:
            return len(self.sheet.getData()[0])
            
    def save(self, filename):
        filename=f"file://{path.abspath(filename)}"
        ## SAVE FILE
        args=(
            PropertyValue('FilterName',0,'calc8',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeAsURL(filename, args)

    def export_pdf(self, filename):
        filename=f"file://{path.abspath(filename)}"
        args=(
            PropertyValue('FilterName',0,'calc_pdf_Export',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeToURL(filename, args)

    def export_xlsx(self, filename):
        filename=f"file://{path.abspath(filename)}"
        args=(
            PropertyValue('FilterName',0,'Calc MS Excel 2007 XML',0),
            PropertyValue('Overwrite',0,True,0),
        )
        self.document.storeToURL(filename, args)


class ODS_Standard(ODS):
    def __init__(self, loserver_port=2002):
        ODS.__init__(self, resource_filename(__name__, 'templates/standard.ods'), loserver_port)

class ODT_Standard(ODT):
    def __init__(self, loserver_port=2002):
        ODT.__init__(self, resource_filename(__name__, 'templates/standard.odt'), loserver_port)
