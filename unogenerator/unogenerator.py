## @namespace unogenerator.unogenerator
## @brief Package that allows to read and write Libreoffice ods and odt files

from datetime import datetime
from os import path, makedirs
from uno import getComponentContext, createUnoStruct, systemPathToFileUrl, Any, ByteSequence
from com.sun.star.beans import PropertyValue
from com.sun.star.text import ControlCharacter
from com.sun.star.awt import Size
from com.sun.star.sheet.ConditionEntryType import COLORSCALE
from com.sun.star.style.ParagraphAdjust import RIGHT,  LEFT
from com.sun.star.style.BreakType import PAGE_BEFORE, PAGE_AFTER
from gettext import translation
from logging import warning, debug
from importlib.resources import files
from os import system
from pydicts import lol, casts
from shutil import copyfile
from socket import socket, AF_INET, SOCK_STREAM
from subprocess import Popen, PIPE
from tempfile import TemporaryDirectory
from time import sleep
from unogenerator import __version__, exceptions
from unogenerator.commons import Coord, ColorsNamed,  Range as R, datetime2uno, guess_object_style, datetime2localc1989, date2localc1989,  time2localc1989,  is_formula, uno2datetime, string_float2object
from pydicts.currency import Currency
from pydicts.percentage import Percentage

def createUnoService(serviceName):
#        resolver = localContext.ServiceManager.createInstance('com.sun.star.bridge.UnoUrlResolver')
  return getComponentContext().ServiceManager.createInstance(serviceName)
  
try:
    t=translation('unogenerator', files("unogenerator") / 'locale')
    _=t.gettext
except:
    _=str

class LibreofficeServer:
    def __init__(self):
        self.pid=None
        self.start()
        
    ## This method allows to use with statement. 
    ## with LibreofficeServer() as server:
    ##      with ODS_Standard
    ## First calls __init__ with None and 2002, then enter, then exit
    def __enter__(self):
        return self
        
    ## Exit function to use with with statement. __enter__ defines enter in with
    def __exit__(self, *args, **kwargs):
        self.stop()

    def start(self):
        # Gets an unued port
        with socket(AF_INET, SOCK_STREAM) as s:
            s.bind(('', 0))  # Bind to port 0 to let the OS assign a free port
            self.port=s.getsockname()[1]

        command=f'loffice --accept="socket,host=localhost,port={self.port};urp;StarOffice.ServiceManager" -env:UserInstallation=file:///tmp/unogenerator{self.port} --headless  --nologo  --norestore'
        process=Popen(command, stdout=PIPE, stderr=PIPE, shell=True)       
        self.pid=process.pid
        
    def stop(self):
        system(f'pkill -f socket,host=localhost,port={self.port};urp;StarOffice.ServiceManager')
        system(f'rm -Rf /tmp/unogenerator{self.port}')

class ODF:
    def __init__(self, template=None,  server=None):
        """
            Common class for ODF instances

            @param template path to template
            @type string
            @param server Server object to use
            @type LibreofficeServer
        """        
        self.start=datetime.now()
        self.server=LibreofficeServer() if server is None else server #Assigns server or auto launch if None
        self.autoserver=server==None
        self.template=None if template is None else systemPathToFileUrl(path.abspath(template))
        maxtries=300
        
        for i in range(maxtries):
            try:
                localContext = getComponentContext()
                resolver = localContext.ServiceManager.createInstance('com.sun.star.bridge.UnoUrlResolver')
                ## self.ctx parece que es mi contexto para servicios
                self.ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={self.server.port};urp;StarOffice.ComponentContext')
                self.desktop = self.ctx.ServiceManager.createInstance('com.sun.star.frame.Desktop')
                self.graphicsprovider=self.ctx.ServiceManager.createInstance("com.sun.star.graphic.GraphicProvider")                   
                args=(
                    PropertyValue('AsTemplate',0,True,0),
                )
                if self.__class__  in (ODS, ODS_Standard):
                    if self.template is None:
                        self.document=self.desktop.loadComponentFromURL('private:factory/scalc','_blank',8,())
                    else:
                        self.document=self.desktop.loadComponentFromURL(self.template,'_blank', 8, args)
                    self.sheet=self.setActiveSheet(0)
                else: #ODT
                    if self.template is None:
                        self.document=self.desktop.loadComponentFromURL('private:factory/swriter','_blank', 8, ())
                    else:
                        self.document=self.desktop.loadComponentFromURL(self.template,'_blank', 8, args)
                    self.cursor=self.document.Text.createTextCursor()
                self.dict_stylenames=self.dictionary_of_stylenames()
                break
            except Exception as e:
                sleeptime=0.25
                sleep(sleeptime)
                if i==maxtries - 1:
                    print(_("This process died after trying to connect to port {0} during {1} seconds. Error: {2}").format(self.loserver_port, maxtries*sleeptime, e))

    ## This method allows to use with statement. 
    ## with ODS() as doc:
    ##      doc.createSheet("WITH")
    ## First calls __init__ with None and 2002, then enter, then exit
    def __enter__(self):
        return self
        
    ## Exit function to use with with statement. __enter__ defines enter in with
    def __exit__(self, *args, **kwargs):
        self.close()

    def calculateAll(self):
        self.document.calculateAll()

    def close(self):
        try:
            self.document.dispose()
        except:
            print (_("Error closing ODF instance"))
        finally:
            if self.autoserver is True:
                self.server.stop()

    ## Generate a dictionary_of_styles with families as key, and a list of string styles as value
    def dictionary_of_stylenames(self):
        stylefam=self.document.StyleFamilies
        d={}
        for sf_index,  sf in enumerate(stylefam):
            r=[]
            styles=list(sf.getElementNames())
            styles.sort()
            for style in styles:
                r.append(style)
            d[stylefam.getElementNames()[sf_index]]=r
        return d

    # Returns a list of page styles objects
    def getPageStylesObjects(self):
        r=[]
        family=self.document.StyleFamilies.getByName('PageStyles')
        for i in range(family.Count):
            r.append(family.getByIndex(i))
        return r
        

    def setLanguage(self, language, country):
        self.language="es"
        self.country="ES"

    def getMetadata(self):
        print("Author",  self.document.DocumentProperties.Author)
        print(dir(self.document.DocumentProperties))
        return {
            "Author": self.document.DocumentProperties.Author, 
            "Description": self.document.DocumentProperties.Description, 
            "Subject": self.document.DocumentProperties.Subject, 
            "Keywords": self.document.DocumentProperties.Keywords, 
            "CreationDate": uno2datetime(self.document.DocumentProperties.CreationDate), 
            "ModificationDate": uno2datetime(self.document.DocumentProperties.ModificationDate), 
            "Title": self.document.DocumentProperties.Title, 
        }
        
    ##Only sets a value when it's different of [] or ""
    def setMetadata(self, title="",  subject="", author="", description="", keywords=[], creationdate=datetime.now()):
        try:
            if author!="":
                self.document.DocumentProperties.Author=author
            self.document.DocumentProperties.Generator=f"UnoGenerator-{__version__}"
            if description!="":
                self.document.DocumentProperties.Description=description
            if subject!="":
                self.document.DocumentProperties.Subject=subject
            if keywords!="":
                self.document.DocumentProperties.Keywords=keywords
            self.document.DocumentProperties.CreationDate=datetime2uno(creationdate)
            self.document.DocumentProperties.ModificationDate=datetime2uno(creationdate)
            if title!="":
                self.document.DocumentProperties.Title=title
        except:
            print("Error setting metadata. Sometimes fails with concurrent process")

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
    def __init__(self, template=None, server=None):
        ODF.__init__(self, template, server)

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
    ## Last is the descriptor returned after first found
    ## @return None if  not found. Descriptor(can be used again with last) of the last found
    def find_and_replace(self, find, replace="", last=None, log=False):
        search=self.document.createSearchDescriptor()
        search.SearchString=find
        if last is None:
            found=self.document.findFirst(search)
        else:
            found=self.document.findNext(last,search)
        
        if found is not None:
            if found.HyperLinkTarget != "" or found.HyperLinkName != "" or found.HyperLinkURL != "":
                url=found.HyperLinkURL
                target=found.HyperLinkTarget
                found.setString(replace)
                found.HyperLinkURL=url
                found.HyperLinkTarget=target
            else:
                found.setString(replace)
            self.cursor=found
        else:
            if log is True:
                warning(f"'{find}' was not found in the document'")
        return found

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
        found=self.find_and_replace(find, "", None, False)
        if found is not None:
#            found.setString("")
            self.cursor=found
            
            oVC = self.document.getCurrentController().getViewCursor()#		'Create View Cursor oVC at the end of document by default?
            oVC.gotoEnd(False)
            self.cursor.gotoRange(oVC,True)			#'Move Text Cursor to same location as oVC while selecting text in between (True)
            self.cursor.setString("")
        else:
            warning(f"'{find}' was not found in the document'")
        
    def addString(self, text, style=None, paragraphBreak=False):
        if style is not None:
            self.cursor.setPropertyValue("ParaStyleName", style)
        self.document.Text.insertString(self.cursor, text, False)
        if paragraphBreak is True:
            self.document.Text.insertControlCharacter(self.cursor, ControlCharacter.PARAGRAPH_BREAK, False)
        
        
        
    def addStringHyperlink(self,  name,  url, paragraphBreak=False):

        oVC = self.document.getCurrentController().getViewCursor()#		'Create View Cursor oVC at the end of document by default?
        text=oVC.getText()
        text.insertString(oVC, name,  True)
        oVC.HyperLinkTarget=url
        oVC.HyperLinkURL=url
#        
        oVC.goRight(len(name), False)
        self.cursor.gotoRange(oVC,False)			#'Move Text Cursor to same location as oVC while selecting text in between (True)
#        self.document.Text.insertString(self.cursor, " ",  False)

        if paragraphBreak is True:
            self.document.Text.insertControlCharacter(self.cursor, ControlCharacter.PARAGRAPH_BREAK, False)
        
    def addParagraph(self,  text,  style="Standard"):
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
        
        
    def addHTMLBlock(self, htmlcode):  
        oStream = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", self.ctx)
        oStream.initialize((ByteSequence(htmlcode.encode()),))

        prop1 = PropertyValue()
        prop1.Name  = "FilterName"
        prop1.Value = "HTML (StarWriter)"
        prop2 = PropertyValue()
        prop2.Name = "InputStream" 
        prop2.Value = oStream
        oVC = self.document.getCurrentController().getViewCursor()#		'Create View Cursor oVC at the end of document by default?
        self.cursor.insertDocumentFromURL("private:stream", (prop1, prop2))
        oVC.goRight(len(htmlcode), False)
        self.cursor.gotoRange(oVC,False)		

        
    def pixelsto100thmm(self, pixels):
        ##                #OOo uses, 1440 twips per inch, which equals 56.7 twips per mm,
        ##The functions TwipsPerPixelX() and TwipsPerPixelY() (both of which return 15 on my system),
        ##Shape is in 100ths of a mm,
        ##Therefore,
        ##PixelsWidth = (1/15) * 56.7 * (ShapeSize.Width/100) [ (pixels/twips) * (twips/mm) * ((mm*100)/100) = pixels ]
        ##PixelsWidth = 0.0378 * ShapeSize.Width
        cm=pixels*15*1/56.7*100#pixels* [twips/pixels]*[mm/twips] * [*100thmm/1mm]
        #print("Converting pixels ",  pixels, "to thmm", cm)

        return cm  
            
        
    ## Returns a text content that can be inserted with document.Text.insertTextContent(cursor,image, False)
    
    ## @param filename_or_bytessequence, Can be a filename path or a bytesequence
    ## @param anchortype AS_CHARACTER, AT_PARAGRAPH
    ## @param name None if we want to use lo default name
    ## @param width 
    ## @param height Puede ser None or a value. If a value, witdth will be set  automatically
    ##      -Width:None -Heigth:None. Usa tamaño por defecto de la imagen (Opción por defecto)
    ##      - Width: 5  -Height:None. Pone 5cm de ancho y la altura automática
    ##      - Width: None  -Height:None5. Pone la anchura automática y 5cm de alto
    ##      - Width: 5  -Height:5. Pone 5cm de ancho de alto
    
    def textcontentImage(self, filename_or_bytessequence, width=None,  height=None, anchortype="AS_CHARACTER", name=None, linked=False ):
        if filename_or_bytessequence.__class__.__name__=="bytes":
            bytes_stream = self.ctx.ServiceManager.createInstanceWithContext('com.sun.star.io.SequenceInputStream', self.ctx)
            bytes_stream.initialize((ByteSequence(filename_or_bytessequence),)) ##Método para inicializar un servicio
            oProps=(
                PropertyValue('InputStream',0, bytes_stream,0),
            )        
            
        else:
            oProps=(
                PropertyValue('URL',0,systemPathToFileUrl(str(filename_or_bytessequence)),0),
                PropertyValue('LoadAsLink',0, linked,0),
            )        
        graphic=self.graphicsprovider.queryGraphic(oProps)
        image = self.document.createInstance("com.sun.star.text.GraphicObject")
        image.Graphic=graphic
        if name is not None:
            image.setName(name)
        image.AnchorType=anchortype
        #Calculate sizes
        if graphic.getType()==1:#Pixel
            ##Size100thMM sometimes is 0 in some images, but no Size
            if graphic.Size100thMM.Height==0 or graphic.Size100thMM.Height==0:
                debug(_("Problems to detect size in 100thmm, I'm trying to calculate it"))
            height100thmm=graphic.Size100thMM.Height if graphic.Size100thMM.Height!=0 else self.pixelsto100thmm(graphic.Size.Height)
            width100thmm=graphic.Size100thMM.Width if graphic.Size100thMM.Width!=0 else self.pixelsto100thmm(graphic.Size.Width)
            #print("100thmm problem ",  graphic.Size100thMM.Width,  graphic.Size100thMM.Height, width100thmm, height100thmm)
            
            
            
            if width is None and height is None:
                width=width100thmm/1000
                height=height100thmm/1000
            elif width is None and height is not None:
                width=round(height*width100thmm/height100thmm, 3)
            elif height is None and width is not None:
                height=round(width*height100thmm/width100thmm, 3)
            image.Size=Size(width*1000, height*1000)
        elif graphic.getType()==2:#Vector
            print("This is a vector graphic. TODO")
        elif graphic.getType()==0:#Empty
            print("This is a empty graphic. TODO")
        #print("Final size", width, height)
        return image
        
    ## @param filename_or_bytessequence_list List of filename path or byte_sequence
    ## @param width float in cm
    ## @param height float in cm
    def addImageParagraph(self, filename_or_bytessequence_list, width=None,  height=None, name=None, style="Illustration", linked=False):
        self.cursor.setPropertyValue("ParaStyleName", style)
        for filename_or_bytessequence in filename_or_bytessequence_list:
            self.document.Text.insertTextContent(self.cursor,self.textcontentImage(filename_or_bytessequence, width, height, "AS_CHARACTER", linked=linked), False)
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
    def __init__(self, template=None, server=None):
        ODF.__init__(self, template, server)
        self._remove_default_sheet=True

    def getRemoveDefaultSheet(self):
        return self._remove_default_sheet
        
    ## @param b Boolean Sets if thiss class remove default sheet if it's empty
    def setRemoveDefaultSheet(self, b):
        self._remove_default_sheet=b
        
    ## Sets sheet pagestyle. If you use ODS_Standard you can set Default (Landscape) and Portrait
    ## Usefull when exporting to pdf
    def setSheetStyle(self, style):
        self.sheet.PageStyle=style
        
    def getSheetNames(self):        
        return self.document.Sheets.ElementNames

    ## Creates a new sheet at the end of the sheets
    ## @param if index is None it creates sheet at the end of the existing sheets
    def createSheet(self, name, index=None):
        for sheet in self.getSheetNames():
            if sheet.upper()==name.upper():
                raise exceptions.UnogeneratorException(_("ERROR: You can't create '{0}' sheet, because it already exists.").format(name)) 
        
        sheets=self.document.getSheets()
        if index is None:
            index=len(sheets)
        sheets.insertNewByName(name, index)
        self.setActiveSheet(index)
        
    def removeSheet(self, index):
        current=self.sheet.Name
        self.setActiveSheet(index)
        self.document.getSheets().removeByName(self.sheet.Name)
        self.document.getSheets().getByName(current)

    def setActiveSheet(self,  index):
        self.sheet_index=index
        self.sheet=self.document.getSheets().getByIndex(index)
        debug(f"Sheet '{self.sheet.Name}' ({self.sheet_index}) is now active")
        return self.sheet
    
    ## l measures are in cm can be float
    def setColumnsWidth(self, l):
        columns=self.sheet.getColumns()
        for i, width in enumerate(l):
            column=columns.getByIndex(i)
            width=width*1000
            column.Width=width ## Are in 1/100th of mm
            
    def setComment(self, coord, comment):
        coord=Coord.assertCoord(coord)
        celladdress= createUnoStruct("com.sun.star.table.CellAddress")
        celladdress.Sheet=self.sheet_index
        celladdress.Column=coord.letterIndex()
        celladdress.Row=coord.numberIndex()
        self.sheet.Annotations.insertNew(celladdress, comment)
            
    ## If addCell is used, preserves styles properties
    ## If you want to change these properties you must add parameters. Not developed
    def addCell(self, coord, o, color=None, outlined=None, alignment=None, decimals=None, bold=None):
        coord=Coord.assertCoord(coord)
        cell=self.sheet.getCellByPosition(coord.letterIndex(), coord.numberIndex())
        self.__object_to_cell(cell, o)

    def addRow(self, coord_start, list_o, formulas=True):        
        """
            Parameters:
                - formulas Boolean. If true formulas will be written as formula. If false as string
                - colors Color: Use color for all array, List of colors one for each cell
                - styles If None uses guest style. Else an array of styles
            Return: Range
        """
        coord_start=Coord.assertCoord(coord_start)
        
        if len(list_o)==0:
            debug(_("addRow is empty. Nothing to write. Ignoring..."))
            return None

        #Convert list_rows to valid dataarray
        r=[]
        for o in list_o:
            r.append(self.__object_to_dataarray_element(o))
        
        #Writes data fast
        range_indexes=[coord_start.letterIndex(), coord_start.numberIndex(), coord_start.letterIndex()+len(list_o)-1, coord_start.numberIndex()]
        range_uno=self.sheet.getCellRangeByPosition(*range_indexes)
        if formulas is True:
            self.__setFormulaArray(range_uno, [r, ])
        else:
            self.__setDataArray(range_uno, [r, ])
        return R.from_uno_range(range_uno)


    def addRowWithStyle(self, coord_start, list_o, colors=ColorsNamed.White,styles=None, formulas=True):        
        """
            Parameters:
                - formulas Boolean. If true formulas will be written as formula. If false as string
                - colors Color: Use color for all array, List of colors one for each cell
                - styles If None uses guest style. Else an array of styles
            Return: Range
        """
        coord_start=Coord.assertCoord(coord_start)        
        range_=self.addRow(coord_start, list_o, formulas)
        if range_ is None:
            return None
        range_uno=range_.uno_range(self.sheet)
        #Fast color:
        if styles is None:
            styles=[]
            for o in list_o:
                styles.append(guess_object_style(o))
            
        
        if colors.__class__==list:
            for i in range(len(list_o)):
                cell=self.sheet.getCellByPosition(coord_start.letterIndex()+i, coord_start.numberIndex())
                cell.setPropertyValue("CellBackColor", colors[i])
        else:
            range_uno.setPropertyValue("CellBackColor", colors)
        #Fast style:
        if styles.__class__==list:
            for i in range(len(list_o)):
                cell=self.sheet.getCellByPosition(coord_start.letterIndex()+i, coord_start.numberIndex())
                cell.setPropertyValue("CellStyle", styles[i])
        else:
            range_uno.setPropertyValue("CellStyle", styles)
        return range_
            


    def addColumn(self, coord_start, list_o, formulas=True):
        """
            Parameters:
                - formulas Boolean. If true formulas will be written as formula. If false as string
                - colors If None uses Wh
                - styles If None uses guest style. Else an array of styles
            Return: Range
        """
        coord_start=Coord.assertCoord(coord_start)
        
        if len(list_o)==0:
            return None

        #Convert list_o to valid dataarray
        r=[]
        for o in list_o:
            r.append([self.__object_to_dataarray_element(o), ])
        
        #Writes data fast
        range_indexes=[coord_start.letterIndex(), coord_start.numberIndex(), coord_start.letterIndex(), coord_start.numberIndex()+len(list_o)-1]
        range_=self.sheet.getCellRangeByPosition(*range_indexes)
        if formulas is True:
            self.__setFormulaArray(range_, r)
        else:
            self.__setDataArray(range_, r)
        return R.from_coords_indexes(*range_indexes)
            
    def addColumnWithStyle(self, coord_start, list_o, colors=ColorsNamed.White,styles=None, formulas=True):
        """
            Parameters:
                - formulas Boolean. If true formulas will be written as formula. If false as string
                - colors Color: Use color for all array, List of colors one for each cell
                - styles If None uses guest style. Else an array of styles
            Return: Range
        """
        coord_start=Coord.assertCoord(coord_start)
        range_=self.addColumn(coord_start, list_o, formulas)
        
        if range_ is None:
            return None
            
        range_uno=range_.uno_range(self.sheet)
        # Guess styles if none
        if styles is None:
            styles=[]
            for o in list_o:
                styles.append(guess_object_style(o))
        #Fast color:
        if colors.__class__==list:
            for i in range(len(list_o)):
                cell=self.sheet.getCellByPosition(coord_start.letterIndex(), coord_start.numberIndex()+i)
                cell.setPropertyValue("CellBackColor", colors[i])
        else:
            range_uno.setPropertyValue("CellBackColor", colors)
        #Fast style:
        if styles.__class__==list:
            for i in range(len(list_o)):
                cell=self.sheet.getCellByPosition(coord_start.letterIndex(), coord_start.numberIndex()+i)
                cell.setPropertyValue("CellStyle", styles[i])
        else:
            range_uno.setPropertyValue("CellStyle", styles)
        return range_
        
        
    
            
    def __setDataArray(self,  unorange,  array_):
        def contains_special_start(iterable, special_chars=('+', '=')):
            if isinstance(iterable, str):
                # Check if the string starts with any of the special characters
                return any(iterable.startswith(char) for char in special_chars)
            try:
                # Try iterating over the iterable
                for item in iterable:
                    if contains_special_start(item, special_chars):
                        return True
            except TypeError:
                # Not an iterable
                pass
            return False
        
        if contains_special_start(array_):
            debug(_("You're trying to add formulas to cells using setDataArray and they will be treated as strings. If you want formula's value use formulas=True in ODS methods"))
            debug(_("  + Range: {0}").format(unorange.AbsoluteName))
        unorange.setDataArray(array_)


    def __setFormulaArray(self,  unorange,  array_):
        unorange.setFormulaArray(array_)


    def addListOfRows(self, coord_start, list_rows, formulas=True):
        """
            Function used to add a big amount of cells to paste quickly
            This method is used when we want to add data without styles, or because we are using a styled template

            Parameters:
                - formulas Boolean. If true formulas will be written as formula. If false as string

            Return: Range
        """
        coord_start=Coord.assertCoord(coord_start) 
        
        rows=len(list_rows)
        if rows==0:
            columns=0
        else:
            columns=len(list_rows[0])
            
            
        if rows==0 or columns==0:
            debug(_("addListOfRowsWithStyle has {0} rows and {1} columns. Nothing to write. Ignoring...").format(rows, columns))
            return 
            
        #Convert list_rows to valid dataarray
        r=[]
        for row in list_rows:
            r_row=[]
            for o in row:
                r_row.append(self.__object_to_dataarray_element(o))
            r.append(r_row)

        #Writes data fast
        range_indexes=[coord_start.letterIndex(), coord_start.numberIndex(), coord_start.letterIndex()+columns-1, coord_start.numberIndex()+rows-1]
        range=self.sheet.getCellRangeByPosition(coord_start.letterIndex(), coord_start.numberIndex(), coord_start.letterIndex()+columns-1, coord_start.numberIndex()+rows-1)
        if formulas is True:
            self.__setFormulaArray(range, r)
        else:
            self.__setDataArray(range, r)
        return R.from_coords_indexes(*range_indexes)

    ## Function used to add a big amount of cells to paste quickly
    ## @param colors. List of column colors, one color, or None to use white
    ## @param styles. List of styles (columns) or None to guess them from first row
    ## @return range of the list_of_rows
    def addListOfRowsWithStyle(self, coord_start, list_rows, colors=ColorsNamed.White, styles=None,  formulas=True):
        """
            Function used to add a big amount of cells to paste quickly
            
            Parameters:
                - formulas Boolean. If true formulas will be written as formula. If false as string
                - colors. List of column colors, one color, or None to use white
                - styles. List of styles (columns) or None to guess them from first row
                
            Return: Range
        """
        coord_start=Coord.assertCoord(coord_start) 
        
        range_=self.addListOfRows(coord_start, list_rows, formulas)
        if range_ is None:
            return
        
        columns=range_.numColumns()
        rows=range_.numRows()

        # Parse colors.
        if colors.__class__.__name__=="list":
            colors=colors
        elif colors is None:
            colors=[ColorsNamed.White]*columns
        else: #one ColorsNamed
            colors=[colors]*columns
        
        # Parse styles
        if styles is None and rows>0:
            styles=[]
            for o in list_rows[0]:
                styles.append(guess_object_style(o))
        elif styles.__class__.__name__=="list":
            styles=styles
        else:
            styles=[styles]*columns
                
                
        if len(colors)!=columns:
            raise exceptions.UnogeneratorException(_("Colors must have the same number of items as data columns"))
        if len(styles)!=columns:
            raise exceptions.UnogeneratorException(_("Styles must have the same number of items as data columns"))
            
        #Create styles by columns cellranges
        if rows>0:
            for c, o in enumerate(list_rows[0]):
                columnrange=self.sheet.getCellRangeByPosition(coord_start.letterIndex()+c, coord_start.numberIndex(), coord_start.letterIndex()+c, coord_start.numberIndex()+rows-1)
                columnrange.setPropertyValue("CellStyle", styles[c])
                columnrange.setPropertyValue("CellBackColor", colors[c])                    
        return range_


    def addListOfColumns(self, coord_start, list_columns, formulas=True):
        """
            Parameters:
                - formulas Boolean. If true formulas will be written as formula. If false as string
                
            Return: Range
        """
        coord_start=Coord.assertCoord(coord_start) 
        list_rows=lol.lol_transposed(list_columns)
        return self.addListOfRows(coord_start, list_rows, formulas)


    ## @param style If None tries to guess it
    def addListOfColumnsWithStyle(self, coord_start, list_columns, colors=ColorsNamed.White, styles=None,  formulas=True):
        """
            Colors and styles are the colors of the first column. Code is different
            
            Parameters:
                - formulas Boolean. If true formulas will be written as formula. If false as string
            
            Return: Range
        """
        coord_start=Coord.assertCoord(coord_start) 
        
        range_=self.addListOfColumns(coord_start, list_columns, formulas)
        if range_ is None:
            return
        
        columns=range_.numColumns()
        rows=range_.numRows()

        # Parse colors.
        if colors.__class__.__name__=="list":
            colors=colors
        elif colors is None:
            colors=[ColorsNamed.White]*rows
        else: #one ColorsNamed
            colors=[colors]*rows
        
        # Parse styles
        if styles is None and rows>0:
            styles=[]
            for o in list_columns[0]:
                styles.append(guess_object_style(o))
        elif styles.__class__.__name__=="list":
            styles=styles
        else:
            styles=[styles]*rows
                
                
        if len(colors)!=rows:
            raise exceptions.UnogeneratorException(_("Colors must have the same number of items as data columns"))
        if len(styles)!=rows:
            raise exceptions.UnogeneratorException(_("Styles must have the same number of items as data columns"))
            
        #Create styles by columns cellranges
        if rows>0:
            for c, o in enumerate(list_columns[0]):
                columnrange=self.sheet.getCellRangeByPosition(coord_start.letterIndex(), coord_start.numberIndex()+c, coord_start.letterIndex()+columns-1, coord_start.numberIndex()+c)
                columnrange.setPropertyValue("CellStyle", styles[c])
                columnrange.setPropertyValue("CellBackColor", colors[c])                    
        return range_
            
    ## @param style If None tries to guess it
    ## @param rewritewrite If color is ColorsNamed.White, rewrites the color to White instead of ignoring it. Ignore it gains 0.200 ms
    ## THIS IS THE WAY TO CREATE FORMULAS
    def addCellWithStyle(self, coord, o, color=ColorsNamed.White, style=None):
        coord=Coord.assertCoord(coord)
        
        if style is None:
            style=guess_object_style(o)
                
        cell=self.sheet.getCellByPosition(coord.letterIndex(), coord.numberIndex())
        self.__object_to_cell(cell, o)
        cell.setPropertyValue("CellStyle", style)
        cell.setPropertyValue("CellBackColor", color)
        
    def setCellName(self, reference, name):
        """
            All of this names are document names
            Parameters:
                - reference -> str: Examples:   $Sheet2.$A$1 
        """
        cell=self.sheet.getCellRangeByName(reference).getCellAddress()
        self.document.NamedRanges.addNewByName(name, reference, cell, 0)
        
    def __object_to_cell(self, cell, o):
        if o.__class__.__name__ in ("int", ):
            cell.setValue(int(o))
        elif o.__class__.__name__ in ("str", ):
            if is_formula(o):
                cell.setFormula(o)
            else:
                cell.setString(o)
        elif o.__class__.__name__  == "datetime":
            cell.setValue(datetime2localc1989(o))
        elif o.__class__.__name__  == "date":
            cell.setValue(date2localc1989(o))
        elif o.__class__.__name__  == "time":
            cell.setValue(time2localc1989(o))
        elif o.__class__.__name__ in ("Percentage", ):
            if o.value is None:
                cell.setString("")
            else:
                cell.setValue(float(o.value))
        elif o.__class__.__name__ in ("Currency", "Money"):
            cell.setValue(float(o.amount))
        elif o.__class__.__name__ in ("Decimal",  "float"):
            cell.setValue(float(o))
        elif o.__class__.__name__ in ("bool", ):
            cell.setValue(int(o))
        elif o.__class__.__name__ in ("timedelta", ):
            cell.setValue(o.total_seconds())
        elif o is None:
            cell.setString("")
        else:
            cell.setString(str(o))
            print("MISSING", o.__class__.__name__)
            
    ## Function used to massive data_array creation
    def __object_to_dataarray_element(self, o):
        if o.__class__.__name__ in ("int", "float"):
            return o
        elif o.__class__.__name__ in ("timedelta", ):
            return o.total_seconds()
        elif o.__class__.__name__ in ("str", ):
            if is_formula(o):
                return o
            else:
                return o
        elif o.__class__.__name__  == "datetime":
            return datetime2localc1989(o)
        elif o.__class__.__name__  == "date":
            return date2localc1989(o)
        elif o.__class__.__name__  == "time":
            return time2localc1989(o)
        elif o.__class__.__name__ in ("Percentage", ):
            if o.value is None:
                 return ""
            else:
                return float(o.value)
        elif o.__class__.__name__ in ("Currency", "Money"):
            return float(o.amount)
        elif o.__class__.__name__ in ("Decimal", ):
            return float(o)
        elif o.__class__.__name__ in ("bool", ):
            return int(o)
        elif o is None:
            return ""
        else:
            return str(o)
            print("MISSING", o.__class__.__name__)

            
    def sortRange(self, range,  sortindex, ascending=True, casesensitive=True):
        range=R.assertRange(range)
        unorange=self.sheet.getCellRangeByName(range.string())
        sortfield=createUnoStruct("com.sun.star.table.TableSortField")
        sortfield.Field=sortindex
        sortfield.IsAscending=ascending
        sortfield.IsCaseSensitive=casesensitive
        sortDescr = unorange.createSortDescriptor()
        

        for prop in sortDescr:
            if prop.Name == 'SortFields':
                prop.Value = Any('[]com.sun.star.table.TableSortField', (sortfield,))
        # sort ...
        unorange.sort(sortDescr)

    def addCellMerged(self, range, o):
        range=R.assertRange(range)
        cell=self.sheet.getCellByPosition(range.c_start.letterIndex(), range.c_start.numberIndex())
        cellrange=self.sheet.getCellRangeByName(range.string())
        cellrange.merge(True)
        self.__object_to_cell(cell, o)
        return cell

    def addCellMergedWithStyle(self, range, o, color=ColorsNamed.White, style=None):
        cell=self.addCellMerged(range, o)
        if style is None:
            style=guess_object_style(o)
        cell.setPropertyValue("CellStyle", style)
        cell.setPropertyValue("CellBackColor", color)

    def freezeAndSelect(self, freeze, selected=None, topleft=None):
        freeze=Coord.assertCoord(freeze) 
        num_columns, num_rows=self.getSheetSize()
        self.document.getCurrentController().setActiveSheet(self.sheet)
        self.document.getCurrentController().freezeAtPosition(freeze.letterIndex(), freeze.numberIndex())

        if selected is None:
            selected=Coord.from_index(num_columns-1, num_rows-1)
        else:
            selected=Coord.assertCoord(selected)
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
            topleft=Coord.from_letters(letter, number)
        else:
            topleft=Coord.assertCoord(topleft)
        self.document.getCurrentController().setFirstVisibleColumn(topleft.letterIndex())
        self.document.getCurrentController().setFirstVisibleRow(topleft.numberIndex())
        
        
    def getValue(self, coord,  detailed=False):
        """
            Gets the value from a coord
            If detailed is True returns a dict with detailed information
        """
        coord=Coord.assertCoord(coord)
        return self.getValueByPosition(coord.letterIndex(), coord.numberIndex(), detailed)
        
    def getValueByPosition(self, letter_index, number_index,  detailed=False):
        """
            Gets the value from a position A1= (0,0)
            If detailed is True returns a dict with detailed information
        """
        r=self.__cell_to_object(self.sheet.getCellByPosition(letter_index, number_index), detailed)
        return r

    ## Returns a list of rows with the values of the sheet
    ## @param sheet_index Integer index of the sheet
    ## @param skip_up int. Number of rows to skip at the begining of the list of rows (lor)
    ## @param skip_down int. Number of rows to skip at the end of the list of rows (lor)
    ## @param detailed Returns a dict {'value': 'A1', 'string': 'A1', 'style': 'Normal', 'class': 'str', 'is_formula': False, 'formula': None}
    ## @param casts List of strings with style names, to cast string and float to objects. Allowed values
    ##      "int", "str", "Decimal", "float", "Percentage", "USD", "EUR",
    ##      The cast columns, so you need to have the same cast items as columns in range
    ## If detailed is True cast is ignored
    ## @return Returns a list of rows of object values
    def getValues(self, skip_up=0, skip_down=0, skip_left=0, skip_right=0, cast=None, detailed=False):
        range_=self.getSheetRange()
        if skip_up>0:
           range_=range_.addRowBefore(-skip_up)
        if skip_down>0:
           range_=range_.addRowAfter(-skip_down)
        if skip_left>0:
           range_=range_.addColumnBefore(-skip_left)
        if skip_right>0:
           range_=range_.addColumnAfter(-skip_right)
        return self.getValuesByRange(range_, cast,  detailed)

    ## Get values by range. If detailed is False, it uses getDataArray method: Dates and datetimes are float. You can convert them using commons.localc1989 methods
    ## @param sheet_index Integer index of the sheet
    ## @param range_ Range object to get values. If None returns all values from sheet
    ## @param detailed Returns a dict {'value': 'A1', 'string': 'A1', 'style': 'Normal', 'class': 'str', 'is_formula': False, 'formula': None}
    ## @param casts List of strings with style names, to cast string and float to objects. Allowed values
    ##      "int", "str", "Decimal", "float", "Percentage", "USD", "EUR",
    ##      The cast columns, so you need to have the same cast items as columns in range
    ## If detailed is True cast is ignored
    ## @return Returns a list of rows of object values
    def getValuesByRange(self, range_,  cast=None, detailed=False):
        range_=R.assertRange(range_)
        
        if detailed is True:#position by position
            r=[]
            for row_indexes in range_.indexes_list():
                rrow=[]
                for column_index,  row_index in row_indexes:
                    rrow.append(self.getValueByPosition(column_index, row_index, detailed))
                r.append(rrow)
            return r
        else:
            range_=R.assertRange(range_)
            range_uno=self.sheet.getCellRangeByPosition(range_.c_start.letterIndex(), range_.c_start.numberIndex(), range_.c_end.letterIndex(), range_.c_end.numberIndex())
            tupleoftuples=range_uno.getDataArray()
            if cast is not None:
                #Reads data fast
                lor=[]
                for tuple_ in tupleoftuples:
                    lor_row=[]
                    for i, tuple_value in enumerate(tuple_):
                        lor_row.append(string_float2object(tuple_value, cast[i]))
                    lor.append(lor_row)
                return lor
            else: #cast false
                return tupleoftuples
       
    ## @param sheet_index Integer index of the sheet
    ## @param column_letter Letter of the column to get values
    ## @param skip Integer Number of top rows to skip in the result
    ## @param detailed Returns a dict {'value': 'A1', 'string': 'A1', 'style': 'Normal', 'class': 'str', 'is_formula': False, 'formula': None}
    ## @param casts List of strings with style names, to cast string and float to objects. Allowed values
    ##      "int", "str", "Decimal", "float", "Percentage", "USD", "EUR",
    ##      The cast columns, so you need to have the same cast items as columns in range
    ## If detailed is True cast is ignored
    ## @return List of values
    def getValuesByColumn(self, column_letter, skip_up=0, skip_down=0,  cast=None, detailed=False):
        range_=self.getSheetRange()
        range_.c_start.letter=column_letter
        range_.c_end.letter=column_letter
        if skip_up>0:
           range_=range_.addRowBefore(-skip_up)
        if skip_down>0:
           range_=range_.addRowAfter(-skip_down)
        lor=self.getValuesByRange(range_, cast, detailed)
        #Transform to list
        r=[]
        for o in lor:
            r.append(o[0])
        return r

    ## @param sheet_index Integer index of the sheet
    ## @param row_number String Number of the row to get values
    ## @param skip Integer Number of top rows to skip in the result
    ## @param detailed Returns a dict {'value': 'A1', 'string': 'A1', 'style': 'Normal', 'class': 'str', 'is_formula': False, 'formula': None}
    ## @param casts List of strings with style names, to cast string and float to objects. Allowed values
    ##      "int", "str", "Decimal", "float", "Percentage", "USD", "EUR",
    ##      The cast columns, so you need to have the same cast items as columns in range
    ## If detailed is True cast is ignored
    ## @return List of values
    def getValuesByRow(self, row_number, skip_left=0, skip_right=0,  cast=None,  detailed=False):
        range_=self.getSheetRange()
        if skip_left>0:
           range_=range_.addColumnBefore(-skip_left)
        if skip_right>0:
           range_=range_.addColumnAfter(-skip_right)
        lor=self.getValuesByRange(range_, cast, detailed)
        return lor[0]



    ## If formula  or without known style returns a tuple (value, formula)
    ## if detailed return a tuple (value, formula, style)
    def __cell_to_object(self, cell, detailed=False):
        isformula=is_formula(cell.getFormula())
        if isformula:
            value=cell.getValue()
        if cell.CellStyle=="Bool":
            value=bool(cell.getValue())
        elif cell.CellStyle in ["Datetime"]:
            value=casts.str2dtnaive(cell.getString(),"%Y-%m-%d %H:%M:%S.")
        elif cell.CellStyle in ["Date"]:
            value=casts.str2date(cell.getString())
        elif cell.CellStyle in ["Time"]:
            value=casts.str2time(cell.getString(), "HH:MM:SS")
        elif cell.CellStyle in ["Float2", "Float6"]:
            value=cell.getValue()
        elif cell.CellStyle in ["Integer"]:
            value=int(cell.getValue())
        elif cell.CellStyle in ["EUR", "USD"]:
            value=Currency(cell.getValue(), cell.CellStyle)
        elif cell.CellStyle in ["Percentage"]:
            value=Percentage(cell.getValue(), 1)
        else:
            value=cell.getString()
        
        if detailed is False:
            return value
        else:
            formula = cell.getFormula() if isformula else None
            return {"value":value, "string":  cell.getString(),  "style":cell.CellStyle, "class": value.__class__.__name__, "is_formula": isformula, "formula": formula}

    ## Return a Range object with the limits of the index sheet
    def getSheetRange(self):
        columns,  rows=self.getSheetSize()
        endcoord=Coord("A1").addRow(rows-1).addColumn(columns-1)
        return R("A1:" + endcoord.string())
        

    ## Returns (columnsNumber, rowsNumber
    def getSheetSize(self):
        data=self.sheet.getData()
        if len(data)==0:
            return 0, 0
        else:
            return len(data[0]), len(data)
            
    ## Method to remove default sheet if empty
    def removeDefaultSheet(self):
        sheets=self.document.getSheets()
        for sheet in sheets:
            if sheet.getName() in ("Hoja1", "Sheet1"):
                data=sheet.getData()
                d=sheet.getCellByPosition(0, 0)
                if len(data)==1 and len(data[0])==1 and d.Value==0 and len(sheets)>1: #Empty and A1 value 0 and number sheets >1
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

    ## If you want to set one page to portrait use setSheetStyle
    ## @filename
    ## @page_per_sheet If True exports pdf with a pdf page per sheet. If False makes a standard export
    def export_pdf(self, filename, page_per_sheet=True):
        if self.getRemoveDefaultSheet() is True:
            self.removeDefaultSheet()
        
        if filename.endswith(".pdf") is False:
            print(_("Filename extension must be 'pdf'."))
            print(_("Document hasn't been saved."))
            return
            
        if page_per_sheet is True:
            oStyleFamilies =self.document.getStyleFamilies()
            oObj_1 = oStyleFamilies.getByName("PageStyles")
            oObj_2 = oObj_1.getByName("Default")
            oObj_2.ScaleToPagesX= 1
            oObj_2.ScaleToPagesY = 1
            args=(
                PropertyValue('FilterName',0,'calc_pdf_Export',0),
                PropertyValue('Overwrite',0,True,0),
            )
        else: #page_per_sheet False
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
            
    def setColorScale(self, range):
        range=R.assertRange(range)
        myCells=self.sheet.getCellRangeByName(range.string())
#        print(unorange, dir(unorange))
#        print(self.sheet, dir(self.sheet))
        myConditionalFormat = myCells.ConditionalFormat
        args=(
            PropertyValue('Operator',0, COLORSCALE,0),
        )
        myConditionalFormat.clear()
        myConditionalFormat.addNew(args)
        myCells.ConditionalFormat = myConditionalFormat
#        self.ws_current.conditional_formatting.add(range, 
#                            openpyxl.formatting.rule.ColorScaleRule(
#                                                start_type='percentile', start_value=0, start_color='00FF00',
#                                                mid_type='percentile', mid_value=50, mid_color='FFFFFF',
#                                                end_type='percentile', end_value=100, end_color='FF0000'
#                                                )
#                                            )

    def toDictionaryOfDetailedValues(self):
        """
            Converts document to a dictionary with detailed values of every sheets
            Has two keys:
                - sheets. A list of dictionaries
                - dictionary. A dictionary with keys (sheet_name, coord_string)
            
            For example: {
                'dictionary': {
                    ('One', 'A1'): {'value': 2, 'string': '2', 'style': 'Integer', 'class': 'int', 'is_formula': False, 'formula': None}, 
                    ('One', 'A2'): {'value': '', 'string': '', 'style': 'Default', 'class': 'str', 'is_formula': False, 'formula': None}, 
                    ('One', 'A3'): {'value': 4, 'string': '4', 'style': 'Integer', 'class': 'int', 'is_formula': True, 'formula': '=2*hola'}, 
                    ('Two', 'A1'): {'value': 5, 'string': '5', 'style': 'Integer', 'class': 'int', 'is_formula': False, 'formula': None}
                }, 
                'sheets': [
                    {'name': 'One', 'columns': 1, 'rows': 3}, 
                    {'name': 'Two', 'columns': 1, 'rows': 1}
                ]
            }

            This method retains more exact values (types) with ODS_Standard although woks fine with ODS
        """
        r={"dictionary":{}, "sheets":[]}
        for i, sheet_name in enumerate(self.getSheetNames()):
            self.setActiveSheet(i)
            sheet_values=self.getValues(detailed=True)
            for number_index, row in enumerate(sheet_values):
                for letter_index, cell in enumerate(row):
                    r["dictionary"][(sheet_name, Coord.from_index(letter_index, number_index).string())]=cell
            r["sheets"].append({"name":sheet_name, "columns": letter_index+1,  "rows": number_index+1})
        return r

class ODS_Standard(ODS):
    def __init__(self, server=None):
        ODS.__init__(self, files('unogenerator') / 'templates/standard.ods', server)

class ODT_Standard(ODT):
    def __init__(self, server=None):
        ODT.__init__(self, files('unogenerator') / 'templates/standard.odt', server)
        self.deleteAll()


