## @namespace unogenerator.unogenerator
## @brief Package that allows to read and write Libreoffice ods and odt files

from datetime import datetime
from os import path, makedirs
from uno import getComponentContext, createUnoStruct, systemPathToFileUrl
from com.sun.star.beans import PropertyValue
from com.sun.star.text import ControlCharacter
from com.sun.star.awt import Size
from com.sun.star.style.ParagraphAdjust import RIGHT,  LEFT
from com.sun.star.style.BreakType import PAGE_BEFORE, PAGE_AFTER
from gettext import translation
from logging import warning, debug
from pkg_resources import resource_filename
from shutil import copyfile
from tempfile import TemporaryDirectory
from unogenerator.commons import Coord as C, ColorsNamed,  Range as R, datetime2uno, guess_object_style, row2index, column2index, datetime2localc1989, date2localc1989,  time2localc1989, Coord_from_letters, Coord_from_index, next_port, get_from_process_numinstances_and_firstport
from unogenerator.reusing.currency import Currency
from unogenerator.reusing.datetime_functions import string2dtnaive, string2date, string2time
from unogenerator.reusing.decorators import timeit
from unogenerator.reusing.percentage import Percentage

def createUnoService(serviceName):
#        resolver = localContext.ServiceManager.createInstance('com.sun.star.bridge.UnoUrlResolver')
  return getComponentContext().ServiceManager.createInstance(serviceName)

        
try:
    t=translation('unogenerator', resource_filename("unogenerator","locale"))
    _=t.gettext
except:
    _=str


class ODF:
    def __init__(self, template=None, loserver_port=2002):
        self.template=None if template is None else systemPathToFileUrl(path.abspath(template))
        self.loserver_port=loserver_port
        self.init=datetime.now()
        self.num_instances, self.first_port=get_from_process_numinstances_and_firstport()
        for i in range(self.num_instances):
            try:
                localContext = getComponentContext()
                resolver = localContext.ServiceManager.createInstance('com.sun.star.bridge.UnoUrlResolver')
                ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={loserver_port};urp;StarOffice.ComponentContext')
                self.desktop = ctx.ServiceManager.createInstance('com.sun.star.frame.Desktop')
                self.graphicsprovider=ctx.ServiceManager.createInstance("com.sun.star.graphic.GraphicProvider")                   
                args=(
                    PropertyValue('AsTemplate',0,True,0),
                )
                if self.__class__  in (ODS, ODS_Standard):
                    if self.template is None:
                        self.document=self.desktop.loadComponentFromURL('private:factory/scalc','_blank',8,())
                    else:
                        self.document=self.desktop.loadComponentFromURL(self.template,'_blank', 8, args)
                    self.sheet=self.setActiveSheet(0)
                    self.numcells=0
                    self.numgetvalues=0
                else: #ODT
                    if self.template is None:
                        self.document=self.desktop.loadComponentFromURL('private:factory/swriter','_blank', 8, ())
                    else:
                        self.document=self.desktop.loadComponentFromURL(self.template,'_blank', 8, args)
                    self.cursor=self.document.Text.createTextCursor()
                break
            except Exception as e:
                print(e)
                old=self.loserver_port
                self.loserver_port=next_port(self.loserver_port, self.first_port, self.num_instances)
                print(_(f"Changing port {old} to {self.loserver_port} {i} times"))
                if i==self.num_instances-1:
                    print(_("This process died"))


    def calculateAll(self):
        self.document.calculateAll()

    def close(self):
        self.document.dispose()
        
    def print_styles(self):
        stylefam=self.document.StyleFamilies
        
        print("Document styles")
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

    def deleteAll(self):
        self.executeDispatch(".uno:SelectAll")
        self.executeDispatch(".uno:Delete")
        
    def executeDispatch(self, command):
        oDisp = createUnoService("com.sun.star.frame.DispatchHelper")
        oDisp.executeDispatch(self.document.getCurrentController().Frame, command, "", 0, ())
        
    ## Poner el tray en el resover y cambiar el puerto cuando except
    
    def loadStylesFromFile(self, filename, overwrite=True):
        styleoptions=list(self.document.StyleFamilies.StyleLoaderOptions)#it's a tuple
        styleoptions.pop()
        styleoptions.append(PropertyValue("OverwriteStyles",0,overwrite,0))
        self.document.StyleFamilies.loadStylesFromURL(systemPathToFileUrl(filename), tuple(styleoptions))

        
                   
class ODT(ODF):
    def __init__(self, template=None, loserver_port=2002):
        ODF.__init__(self, template, loserver_port)

    @timeit
    def save(self, filename, overwrite_template=False):
        if filename==self.template and overwrite_template is False:
            print(_("You can't use the same filename as your template or you will overwrite it."))
            print(_("You can force it, setting overwrite_template paramter to True"))
            print(_("Document hasn't been saved."))
            return
        if filename.endswith(".odt") is False:
            print(_("Filename extension must be 'odt'."))
            print(_("Document hasn't been saved."))
            return
        ## SAVE FILE
        args=(
            PropertyValue('FilterName',0,'writer8',0),
            PropertyValue('Overwrite',0,True,0),
        )
    

        with TemporaryDirectory() as tmpdirname:
                tempfile=f"{tmpdirname}/{path.basename(filename)}"
                self.document.storeAsURL(systemPathToFileUrl(tempfile), args)
                makedirs(path.dirname(path.abspath(filename)), exist_ok=True)
                copyfile(tempfile, filename)
        
                
    def export_pdf(self, filename):
        if filename.endswith(".pdf") is False:
            print(_("Filename extension must be 'pdf'."))
            print(_("Document hasn't been saved."))
            return
        args=(
            PropertyValue('FilterName',0,'writer_pdf_Export',0),
            PropertyValue('Overwrite',0,True,0),
        )
        with TemporaryDirectory() as tmpdirname:
            tempfile=f"{tmpdirname}/{path.basename(filename)}"
            self.document.storeToURL(systemPathToFileUrl(tempfile), args)
            
            makedirs(path.dirname(path.abspath(filename)), exist_ok=True)
            copyfile(tempfile, filename)
         
    def export_docx(self, filename):
        if filename.endswith(".docx") is False:
            print(_("Filename extension must be 'docx'."))
            print(_("Document hasn't been saved."))
            return
        args=(
            PropertyValue('FilterName',0,'MS Word 2007 XML',0),
            PropertyValue('Overwrite',0,True,0),
        )

        with TemporaryDirectory() as tmpdirname:
            tempfile=f"{tmpdirname}/{path.basename(filename)}"
            self.document.storeToURL(systemPathToFileUrl(tempfile), args)
            makedirs(path.dirname(path.abspath(filename)), exist_ok=True)
            copyfile(tempfile, filename)

    ## This method has problems with ().\ *// especial characters. Findall and replace works fine
    def find_and_replace(self, find, replace="", log=False):
        search=self.document.createSearchDescriptor()
        search.SearchString=find
        found=self.document.findFirst(search)
        if found is not None:
            found.setString("")
            self.cursor=found
            found.Text.insertString(self.cursor, replace, False)
            return True
        else:
            if log is True:
                warning(f"'{find}' was not found in the document'")
            return False

    def findall_and_replace(self, find, replace="", log=False):
        search=self.document.createReplaceDescriptor()
        search.setSearchString(find)
        search.setReplaceString(replace)
        found=self.document.findFirst(search)
        if found is not None:
            number=self.document.replaceAll(search)
            debug(f"Replaced {number} times '{find}' by '{replace}'")
            self.cursor=found
            return number
        else:
            if log is True:
                debug(f"'{find}' was not found in the document'")
            return 0

    ## Sets the cursor after found string
    def find(self, find, log=False):
        self.first_and_replace(find, find, log)

    ## Find a string and deleteAll untill the end of the document
    ## Usefull to delete document part with styles in templates
    def find_and_delete_until_the_end_of_document(self, find):
        search=self.document.createSearchDescriptor()
        search.SearchString=find
        found=self.document.findFirst(search)
        if found is not None:
            found.setString("")
            self.cursor=found
            
            oVC = self.document.getCurrentController().getViewCursor()#		'Create View Cursor oVC at the end of document by default?
            oVC.gotoEnd(False)
            self.cursor.gotoRange(oVC,True)			#'Move Text Cursor to same location as oVC while selecting text in between (True)
            self.cursor.setString("")
        else:
            warning(f"'{find}' was not found in the document'")
        
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
    def textcontentImage(self, filename, width=2,  height=2, anchortype="AS_CHARACTER", name=None, linked=False ):
        oProps=(
            PropertyValue('URL',0,systemPathToFileUrl(filename),0),
            PropertyValue('LoadAsLink',0, linked,0),
        )        
        graphic=self.graphicsprovider.queryGraphic(oProps)
        image = self.document.createInstance("com.sun.star.text.GraphicObject")
        image.Graphic=graphic
        if name is not None:
            image.setName(name)
        image.AnchorType=anchortype
        image.Size=Size(width*1000, height*1000)
        return image
        
    ## @param width float in cm
    ## @param height float in cm
    def addImageParagraph(self, filename_list, width=2,  height=2, name=None, style="Illustration", linked=False):
        self.cursor.setPropertyValue("ParaStyleName", style)
        for filename in filename_list:
            self.document.Text.insertTextContent(self.cursor,self.textcontentImage(filename, width, height, "AS_CHARACTER", linked=linked), False)
        self.document.Text.insertControlCharacter(self.cursor, ControlCharacter.PARAGRAPH_BREAK, False)


    ## Table data, must be all strings (Not a spreadsheet)
    ## @params magings TRBL
    ## @params columnssize_percentage Por ejemplo para table 30,70, solo debo poner [30,70]
    ##   para table [10,10,80], solo debo poner [10,10]
    ## if none , rangos iguales
    def addTableParagraph(self, 
        data,  
        columnssize_percentages=None, 
        margins_top=0.3,
        margins_bottom=0.3,
        size=10,  
        width_percentage=100, 
        alignment="center",  
        paragraph_style="Standard", 
        style=None, 
        name=None
    ):        
        #Tabla definition
        table=self.document.createInstance("com.sun.star.text.TextTable")
        num_rows=len(data)
        if num_rows==0:
            return
        num_columns=len(data[0])
        table.initialize(num_rows, num_columns)
        if name is not None:
            table.TableName=name

        #Párrafo de la table
        self.cursor.setPropertyValue("ParaStyleName", paragraph_style)
        self.document.Text.insertTextContent(self.cursor,table,False)
        
        #Table data, must be all strings (Not a spreadsheet)
        if style is not None:
            table.TableTemplateName=style #MUST BE BEFORE TO OVERRIDE SOME PROPERTIES
        for row, row_data in enumerate(data):
            for column, cell_data in enumerate(row_data):
                cell=table.getCellByPosition( column, row)
                cursor = cell.createTextCursor()
                cursor.CharHeight=size
                if str(cell_data).startswith("-"):
                    cursor.setPropertyValue( "CharColor", 0xDD0000 )
                if style in ["Table0", "Table1", "Table1Total"] and row==0:
                    pass # Centered text for styles headers
                else:
                    if cell_data.__class__.__name__ in ["str"]:
                        cursor.ParaAdjust=LEFT
                    else:
                        cursor.ParaAdjust=RIGHT
                        cursor.ParaRightMargin=100
                cell.insertString(cursor, str(cell_data), False)
        
        #TAble width and style
        table.HoriOrient=2 #Centered
        table.TopMargin=margins_top*1000
        table.BottomMargin=margins_bottom*1000
        table.RelativeWidth=width_percentage #PARECE QUE ES SOLO DESCRIPTIVO
        #print("RelativeSum", table.TableColumnRelativeSum)
        #Columns width
        if columnssize_percentages is not None:
            separators=[]
            #print("Before", *self.printSeparators(table, table.TableColumnSeparators))
            for i, sep in enumerate(columnssize_percentages[:-1]):
                separator= createUnoStruct("com.sun.star.text.TableColumnSeparator")
                separator.Position=sum(columnssize_percentages[0:i+1])*100
                separator.IsVisible=True
                separators.append(separator)
            #print("Setting",  *self.printSeparators(table, separators))
            table.TableColumnSeparators=separators
            #print("After", *self.printSeparators(table, table.TableColumnSeparators))
            #print()

    def printSeparators(self, table, separators):
            r=[]
            for i in range(len(separators)):
                r.append(separators[i].Position)
            return (table.TableColumnRelativeSum, str(r))
        
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
            
    def paragraphBreak(self):
        self.document.Text.insertControlCharacter(self.cursor, ControlCharacter.PARAGRAPH_BREAK, False)
        
    ## Default styles are Landscape
    def pageBreak(self, style="Standard", page_before=True):
        if page_before is True:
            self.cursor.BreakType = PAGE_BEFORE 
        else:
            self.cursor.BreakType = PAGE_AFTER
            
        self.cursor.setPropertyValue("PageDescName", style)
        self.document.Text.insertString(self.cursor, "", False)

class ODS(ODF):
    def __init__(self, template=None, loserver_port=2002):
        ODF.__init__(self, template, loserver_port)
        self._remove_default_sheet=True
        
    def getRemoveDefaultSheet(self):
        return self._remove_default_sheet
        
    ## @param b Boolean Sets if thiss class remove default sheet if it's empty
    def setRemoveDefaultSheet(self, b):
        self._remove_default_sheet=b
        
    ## Creates a new sheet at the end of the sheets
    ## @param if index is None it creates sheet at the end of the existing sheets
    @timeit
    def createSheet(self, name, index=None):
        sheets=self.document.getSheets()
        if index is None:
            index=len(sheets)
        sheets.insertNewByName(name, index)
        self.setActiveSheet(index)
        
    @timeit
    def removeSheet(self, index):
        current=self.sheet.Name
        self.setActiveSheet(index)
        self.document.getSheets().removeByName(self.sheet.Name)
        self.document.getSheets().getByName(current)

    def setActiveSheet(self,  index):
        self.sheet_index=index
        self.sheet=self.document.getSheets().getByIndex(index)
        return self.sheet
    
    ## l measures are in cm can be float
    @timeit
    def setColumnsWidth(self, l):
        columns=self.sheet.getColumns()
        for i, width in enumerate(l):
            column=columns.getByIndex(i)
            width=width*1000
            column.Width=width ## Are in 1/100th of mm
            
    def setComment(self, coord, comment):
        coord=C.assertCoord(coord)
        celladdress= createUnoStruct("com.sun.star.table.CellAddress")
        celladdress.Sheet=self.sheet_index
        celladdress.Column=coord.letterIndex()
        celladdress.Row=coord.numberIndex()
        self.sheet.Annotations.insertNew(celladdress, comment)
            
    def addCell(self, coord, o, color_dict=ColorsNamed.White, outlined=1, alignment="left", decimals=2, bold=False):
        coord=C.assertCoord(coord)
        cell=self.sheet.getCellByPosition(coord.letterIndex(), coord.numberIndex())
        self.__object_to_cell(cell, o)
            
        border_prop = createUnoStruct("com.sun.star.table.BorderLine2")
        border_prop.LineWidth = outlined
        cell.setPropertyValue("TopBorder", border_prop)
        cell.setPropertyValue("LeftBorder", border_prop)
        cell.setPropertyValue("RightBorder", border_prop)
        cell.setPropertyValue("BottomBorder", border_prop)
        cell.setPropertyValue("CellBackColor", color_dict["color"])

    ## @param colors If None uses Wh
    ## @param styles If None uses guest style. Else an array of styles
    def addRowWithStyle(self, coord_start, list_o, colors=ColorsNamed.White,styles=None):
        coord_start=C.assertCoord(coord_start)
        if styles is None:
            styles=[]
            for o in list_o:
                styles.append(guess_object_style(o))
        elif styles.__class__.__name__!="list":
            styles=[styles]*len(list_o)

        if colors.__class__.__name__!="list":
            colors=[colors]*len(list_o)

        for i,o in enumerate(list_o):
            self.addCellWithStyle(coord_start.addColumnCopy(i),o,colors[i],styles[i])

    ## @param colors If None uses Wh
    ## @param styles If None uses guest style. Else an array of styles
    def addColumnWithStyle(self, coord_start, list_o, colors=ColorsNamed.White,styles=None):
        coord_start=C.assertCoord(coord_start)
        if styles is None:
            styles=[]
            for o in list_o:
                styles.append(guess_object_style(o))
        elif styles.__class__.__name__!="list":
            styles=[styles]*len(list_o)

        if colors.__class__.__name__!="list":
            colors=[colors]*len(list_o)

        for i,o in enumerate(list_o):
            self.addCellWithStyle(coord_start.addRowCopy(i),o,colors[i],styles[i])

    ## @param style If None tries to guess it
    def addListOfRowsWithStyle(self, coord_start, list_rows, colors=ColorsNamed.White, styles=None):
        coord_start=C.assertCoord(coord_start)
        for i, row in enumerate(list_rows):
            self.addRowWithStyle(coord_start.addRowCopy(i), row, colors=colors,styles=styles)

    ## @param style If None tries to guess it
    def addListOfColumnsWithStyle(self, coord_start, list_columns, colors=ColorsNamed.White, styles=None):
        coord_start=C.assertCoord(coord_start)
        for i, column in enumerate(list_columns):
            self.addColumnWithStyle(coord_start.addColumnCopy(i), column, colors=colors,styles=styles)

    ## @param style If None tries to guess it
#    @timeit
    def addCellWithStyle(self, coord, o, color=ColorsNamed.White, style=None):
        coord=C.assertCoord(coord)
        if style is None:
            style=guess_object_style(o)
        cell=self.sheet.getCellByPosition(coord.letterIndex(), coord.numberIndex())
        self.__object_to_cell(cell, o)
        cell.setPropertyValue("CellStyle", style)
        cell.setPropertyValue("CellBackColor", color)
        self.numcells=self.numcells+1
        if self.numcells % 500==0:
            debug(f"Wrote {self.numcells} cells in {datetime.now()-self.init}")
        
    ## All of this names are document names
    def setCellName(self, coord, name):
        coord=C.assertCoord(coord)
        cell=self.sheet.getCellRangeByName(coord.string()).getCellAddress()
        self.document.NamedRanges.addNewByName(name, coord.string(), cell, 0)
        
    def __object_to_cell(self, cell, o):
        if o.__class__.__name__  == "datetime":
            cell.setValue(datetime2localc1989(o))
        elif o.__class__.__name__  == "date":
            cell.setValue(date2localc1989(o))
        elif o.__class__.__name__  == "time":
            cell.setValue(time2localc1989(o))
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
        elif o is None:
            cell.setString("")
        else:
            cell.setString(str(o))
            print("MISSING", o.__class__.__name__)

    def showStatistics(self):
        debug(f"- Number of cells {self.numcells}")
        debug(f"- Number of get values {self.numgetvalues}")
        debug(f"- Total time {datetime.now()-self.init}")

    @timeit
    def addCellMergedWithStyle(self, range, o, color=ColorsNamed.White, style=None):
        range=R.assertRange(range)
        cell=self.sheet.getCellByPosition(range.start.letterIndex(), range.start.numberIndex())
        cellrange=self.sheet.getCellRangeByName(range.string())
        cellrange.merge(True)
        self.__object_to_cell(cell, o)
        if style is None:
            style=guess_object_style(o)
        cell.setPropertyValue("CellStyle", style)
        cell.setPropertyValue("CellBackColor", color)
        self.numcells=self.numcells+1

    @timeit
    def freezeAndSelect(self, freeze, selected=None, topleft=None):
        freeze=C.assertCoord(freeze) 
        num_columns, num_rows=self.getSheetSize()
        self.document.getCurrentController().setActiveSheet(self.sheet)
        self.document.getCurrentController().freezeAtPosition(freeze.letterIndex(), freeze.numberIndex())

        if selected is None:
            selected=Coord_from_index(num_columns-1, num_rows-1)
        else:
            selected=C.assertCoord(selected)
        selectedcell=self.sheet.getCellByPosition(selected.letterIndex(), selected.numberIndex())
        self.document.getCurrentController().select(selectedcell)
        
        if topleft is None:
            #Get the select coord - 20 rows and - 10 columns
            minus_coord=selected.addRowCopy(-20).addColumnCopy(-10)
            if minus_coord.letterIndex()<=freeze.letterIndex():#letter
                letter=freeze.letter
            else:
                letter=minus_coord.letter
            if minus_coord.numberIndex()<=freeze.numberIndex():#number
                number=freeze.number
            else:
                number=minus_coord.number
            topleft=Coord_from_letters(letter, number)
        else:
            topleft=C.assertCoord(topleft)
        self.document.getCurrentController().setFirstVisibleColumn(topleft.letterIndex())
        self.document.getCurrentController().setFirstVisibleRow(topleft.numberIndex())
        
    def getValue(self, coord, standard=True):
        coord=C.assertCoord(coord)
        self.numgetvalues=self.numgetvalues+1
        return self.getValueByPosition(coord.letterIndex(), coord.numberIndex(), standard)
        
    @timeit
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
            elif cell.CellStyle in ["Time"]:
                return string2time(cell.getString(), "HH:MM:SS")
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
        return R("A1:" + endcoord.string())

    ## Returns (columnsNumber, rowsNumber
    @timeit
    def getSheetSize(self):
        data=self.sheet.getData()
        if len(data)==0:
            return 0, 0
        else:
            return len(data[0]), len(data)
            
    ## Method to remove default sheet if empty
    def removeDefaultSheet(self):
        for sheet in self.document.getSheets():
            if sheet.getName() in ("Hoja1", "Sheet1") :
                data=sheet.getData()
                if len(data)==1 and len(data[0])==1:
                    self.document.getSheets().removeByName(sheet.getName())
            
    def save(self, filename, overwrite_template=False):
        if self.getRemoveDefaultSheet() is True:
            self.removeDefaultSheet()
            
        if filename==self.template and overwrite_template is False:
            print(_("You can't use the same filename as your template or you will overwrite it."))
            print(_("You can force it, setting overwrite_template paramter to True"))
            print(_("Document hasn't been saved."))
            return
        if filename.endswith(".ods") is False:
            print(_("Filename extension must be 'ods'."))
            print(_("Document hasn't been saved."))
            return
        ## SAVE FILE
        args=(
            PropertyValue('FilterName', 0, 'calc8', 0),
            PropertyValue('Overwrite', 0, True, 0),
        )
        with TemporaryDirectory() as tmpdirname:
                tempfile=f"{tmpdirname}/{path.basename(filename)}"
                self.document.storeAsURL(systemPathToFileUrl(tempfile), args)
                makedirs(path.dirname(path.abspath(filename)), exist_ok=True)
                copyfile(tempfile, filename)
        self.showStatistics()

    def export_pdf(self, filename):
        if self.getRemoveDefaultSheet() is True:
            self.removeDefaultSheet()
        
        if filename.endswith(".pdf") is False:
            print(_("Filename extension must be 'pdf'."))
            print(_("Document hasn't been saved."))
            return
        args=(
            PropertyValue('FilterName',0,'calc_pdf_Export',0),
            PropertyValue('Overwrite',0,True,0),
        )
        with TemporaryDirectory() as tmpdirname:
            tempfile=f"{tmpdirname}/{path.basename(filename)}"
            self.document.storeToURL(systemPathToFileUrl(tempfile), args)
            makedirs(path.dirname(path.abspath(filename)), exist_ok=True)
            copyfile(tempfile, filename)


    def export_xlsx(self, filename):
        if self.getRemoveDefaultSheet() is True:
            self.removeDefaultSheet()

        if filename.endswith(".xlsx") is False:
            print(_("Filename extension must be 'xlsx'."))
            print(_("Document hasn't been saved."))
            return
        args=(
            PropertyValue('FilterName',0,'Calc MS Excel 2007 XML',0),
            PropertyValue('Overwrite',0,True,0),
        )
        with TemporaryDirectory() as tmpdirname:
            tempfile=f"{tmpdirname}/{path.basename(filename)}"
            self.document.storeToURL(systemPathToFileUrl(tempfile), args)
            makedirs(path.dirname(path.abspath(filename)), exist_ok=True)
            copyfile(tempfile, filename)


class ODS_Standard(ODS):
    def __init__(self, loserver_port=2002):
        ODS.__init__(self, resource_filename(__name__, 'templates/standard.ods'), loserver_port)

class ODT_Standard(ODT):
    def __init__(self, loserver_port=2002):
        ODT.__init__(self, resource_filename(__name__, 'templates/standard.odt'), loserver_port)
        self.deleteAll()


