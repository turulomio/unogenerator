## @namespace unogenerator.demo
## @brief Generate ODF example files
from uno import getComponentContext
getComponentContext()
import argparse
import gettext
import pkg_resources
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime, date, timedelta
from multiprocessing import cpu_count

from unogenerator.commons import __version__, addDebugSystem, argparse_epilog, Colors, Coord as C
from unogenerator.reusing.currency import Currency
from unogenerator.reusing.percentage import Percentage
from unogenerator.unogenerator import ODT, ODS
from os import remove

try:
    t=gettext.translation('unogenerator',pkg_resources.resource_filename("unogenerator","locale"))
    _=t.gettext
except:
    _=str

def remove_without_errors(filename):
    try:
        remove(filename)
    except OSError as e:
        print(_("Error deleting: {} -> {}".format(filename, e.strerror)))

## If arguments is None, launches with sys.argc parameters. Entry point is toomanyfiles:main
## You can call with main(['--pretend']). It's equivalento to os.system('program --pretend')
## @param arguments is an array with parser arguments. For example: ['--argument','9']. 
def main(arguments=None):
    parser=argparse.ArgumentParser(prog='unogenerator', description=_('Create example files using unogenerator module'), epilog=argparse_epilog(), formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--debug', help="Debug program information", choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")
    group= parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--create', help="Create demo files", action="store_true",default=False)
    group.add_argument('--remove', help="Remove demo files", action="store_true", default=False)
    args=parser.parse_args(arguments)

    addDebugSystem(args.debug)

    if args.remove==True:
        remove_without_errors("unogenerator.ods")
        remove_without_errors("unogenerator.odt")

    if args.create==True:
        start=datetime.now()
        futures=[]
        with ProcessPoolExecutor(max_workers=cpu_count()+1) as executor:
            futures.append(executor.submit(demo_ods))
            futures.append(executor.submit(demo_ods_standard))
            futures.append(executor.submit(demo_odt))
            futures.append(executor.submit(demo_odt_standard))

        for future in as_completed(futures):
            print(future.result())
            
        print(demo_ods_standard_read())
        print("All process took {}".format(datetime.now()-start))


def demo_ods():
    doc=ODS("unogenerator.ods")
    doc.setMetadata(
        _("UnoGenerator ODS example"),  
        _("Demo with ODS class"), 
        "Turulomio", 
        _(f"This file have been generated with UnoGenerator-{__version__}. You can see UnoGenerator main page in http://github.com/turulomio/unogenerator"), 
        ["unogenerator", "demo", "files"]
    )
    doc.createSheet("Hard work", 1)
    doc.setColumnsWidth([2000, 5000, 2000,  2000,  2000,  2000,  5000,  5000,  5000,  5000,  5000,  5000])
    
    doc.addCell("A1", _("Style name"), Colors["Orange"])
    doc.addCell("B1", _("Date and time"), Colors["Orange"])
    doc.addCell("C1", _("Date"), Colors["Orange"])
    doc.addCell("D1", _("Integer"), Colors["Orange"])
    doc.addCell("E1", _("Euros"), Colors["Orange"])
    doc.addCell("F1", _("Dollars"), Colors["Orange"])
    doc.addCell("G1", _("Percentage"), Colors["Orange"])
    doc.addCell("H1", _("Number with 2 decimals"), Colors["Orange"])
    doc.addCell("I1", _("Number with 6 decimals"), Colors["Orange"])
    doc.addCell("J1", _("Time"), Colors["Orange"])
    doc.addCell("K1", _("Boolean"), Colors["Orange"])
    for row, color_key in enumerate(Colors.keys()):
        doc.addCell(C("A2").addRow(row), color_key, Colors[color_key])
        doc.addCell(C("B2").addRow(row), datetime.now(), Colors[color_key])
        doc.addCell(C("C2").addRow(row), date.today(), Colors[color_key])
        doc.addCell(C("D2").addRow(row), pow(-1, row)*-10000000, Colors[color_key])
        doc.addCell(C("E2").addRow(row), Currency(pow(-1, row)*12.56, "EUR"), Colors[color_key])
        doc.addCell(C("F2").addRow(row), Currency(pow(-1, row)*12345.56, "USD"), Colors[color_key])
        doc.addCell(C("G2").addRow(row), Percentage(pow(-1, row)*1, 3), Colors[color_key])
        doc.addCell(C("H2").addRow(row), pow(-1, row)*123456789.121212, Colors[color_key])
        doc.addCell(C("I2").addRow(row), pow(-1, row)*-12.121212, Colors[color_key])
        doc.addCell(C("J2").addRow(row), (datetime.now()+timedelta(seconds=3600*12*row)).time(), Colors[color_key])
        doc.addCell(C("K2").addRow(row), bool(row%2), Colors[color_key])
    doc.freezeAndSelect("B2", "K9", "E4")
    doc.removeSheet(0)
    doc.save()
    doc.export_xlsx()
    doc.close()
        
        
    return "demo_ods took {}".format(datetime.now()-doc.init)
        
def demo_ods_standard():
    doc=ODS("unogenerator_standard.ods", pkg_resources.resource_filename(__name__, 'templates/standard.ods'))
    doc.setMetadata(
        _("UnoGenerator ODS example"),  
        _("Demo with ODS class"), 
        "Turulomio", 
        _(f"This file have been generated with UnoGenerator-{__version__}. You can see UnoGenerator main page in http://github.com/turulomio/unogenerator"), 
        ["unogenerator", "demo", "files"]
    )
    doc.createSheet("Styles", 1)
    doc.setColumnsWidth([2000, 5000, 2000,  2000,  2000,  2000,  5000,  5000,  5000,  5000,  5000,  5000])
    
    doc.addCellWithStyle("A1", _("Style name"), Colors["Orange"], "BoldCenter")
    doc.addCellWithStyle("B1", _("Date and time"), Colors["Orange"], "BoldCenter")
    doc.addCellWithStyle("C1", _("Date"), Colors["Orange"], "BoldCenter")
    doc.addCellWithStyle("D1", _("Integer"), Colors["Orange"], "BoldCenter")
    doc.addCellWithStyle("E1", _("Euros"), Colors["Orange"], "BoldCenter")
    doc.addCellWithStyle("F1", _("Dollars"), Colors["Orange"], "BoldCenter")
    doc.addCellWithStyle("G1", _("Percentage"), Colors["Orange"], "BoldCenter")
    doc.addCellWithStyle("H1", _("Number with 2 decimals"), Colors["Orange"], "BoldCenter")
    doc.addCellWithStyle("I1", _("Number with 6 decimals"), Colors["Orange"], "BoldCenter")
    doc.addCellWithStyle("J1", _("Time"), Colors["Orange"], "BoldCenter")
    doc.addCellWithStyle("K1", _("Boolean"), Colors["Orange"], "BoldCenter")
    for row, color_key in enumerate(Colors.keys()):
        doc.addCellWithStyle(C("A2").addRow(row), color_key, Colors[color_key], "Bold")
        doc.addCellWithStyle(C("B2").addRow(row), datetime.now(), Colors[color_key], "Datetime")
        doc.addCellWithStyle(C("C2").addRow(row), date.today(), Colors[color_key], "Date")
        doc.addCellWithStyle(C("D2").addRow(row), pow(-1, row)*-10000000, Colors[color_key], "Integer")
        doc.addCellWithStyle(C("E2").addRow(row), Currency(pow(-1, row)*12.56, "EUR"), Colors[color_key], "EUR")
        doc.addCellWithStyle(C("F2").addRow(row), Currency(pow(-1, row)*12345.56, "USD"), Colors[color_key], "USD")
        doc.addCellWithStyle(C("G2").addRow(row), Percentage(pow(-1, row)*1, 3), Colors[color_key],  "Percentage")
        doc.addCellWithStyle(C("H2").addRow(row), pow(-1, row)*123456789.121212, Colors[color_key], "Float6")
        doc.addCellWithStyle(C("I2").addRow(row), pow(-1, row)*-12.121212, Colors[color_key], "Float2")
        doc.addCellWithStyle(C("J2").addRow(row), (datetime.now()+timedelta(seconds=3600*12*row)).time(), Colors[color_key], "Time")
        doc.addCellWithStyle(C("K2").addRow(row), bool(row%2), Colors[color_key], "Bool")
        
    doc.addCellWithStyle("E10","=sum(E2:E8)", Colors["GrayLight"], "EUR" )
    doc.addCellMerged("E12:K12", "Prueba de merge", Colors["Yellow"], style="BoldCenter")
    doc.setComment("B11", "This is nice comment")
    
    doc.freezeAndSelect("B2")
    doc.removeSheet(0)
#    doc.calculateAll()
    doc.save()
    doc.export_xlsx()
    doc.close()

    return "demo_ods_standard took {}".format(datetime.now()-doc.init)
    
    
def demo_ods_standard_read():
    doc=ODS("donotsaveme.ods", "./unogenerator_standard.ods")
    doc.setActiveSheet(0)
    print(doc.getValue("E9"))
    print(doc.getValue("E10"))
    doc.close()
    return "demo_ods_standard took {}".format(datetime.now()-doc.init)
   

    
def demo_odt():
    doc=ODT("unogenerator.odt")
    doc.setMetadata(
        _("UnoGenerator ODT example"),  
        _("Demo with ODT class"), 
        "Turulomio", 
        _(f"This file have been generated with UnoGenerator-{__version__}. You can see UnoGenerator main page in http://github.com/turulomio/unogenerator"), 
        ["unogenerator", "demo", "files"]
    )
    doc.save()
    doc.export_docx()
    doc.export_pdf()
    doc.close()
    return "demo_odt took {}".format(datetime.now()-doc.init)    
    
def demo_odt_standard():
    doc=ODT("unogenerator_standard.odt", pkg_resources.resource_filename(__name__, 'templates/standard.odt'))
    doc.setMetadata(
        _("UnoGenerator ODT example"),  
        _("Demo with ODT class"), 
        "Turulomio", 
        _(f"This file have been generated with UnoGenerator-{__version__}. You can see UnoGenerator main page in http://github.com/turulomio/unogenerator"), 
        ["unogenerator", "demo", "files"]
    )
    doc.addParagraph(_("Manual of UnoGenerator"), "Title")
    doc.addParagraph(_(f"Version: {__version__}"), "Subtitle")
    
    doc.addParagraph(_("ODT"), "Heading 1")
#    doc.print_styles()
    doc.addParagraph(
        _("ODT files can be quickly generated with OfficeGenerator.") + " " + 
        _("It create predefined styles that allows to create nice documents without worry about styles."),  "Standard"
    )
    doc.addParagraph(_("OfficeGenerator predefined paragraph styles"), "Heading 2")
    doc.addParagraph(
        _("OfficeGenerator has headers and titles as you can see in the document structure.") + " " + 
        _("Morever, it has the following predefined styles:"), "Standard"
    )
    
    doc.addParagraph(_("This is the 'Standard' style"), "Standard")
    doc.addParagraph(_("This is the 'StandardCenter' style"), 'Standard')
    doc.addParagraph(_("This is the 'StandardRight' style"), 'Standard')
    doc.addParagraph(_("This is the 'Illustration' style"), 'Illustration')
    doc.addParagraph(_("This is the 'Bold18Center' style"), 'Standard')
    doc.addParagraph(_("This is the 'Bold16Center' style"), 'Standard')
    doc.addParagraph(_("This is the 'Bold14Center' style"), 'Standard')
    doc.addParagraph(_("This is the 'Bold12Center' style"), 'Standard')
    doc.addParagraph(_("This is the 'Bold12Underline' style"), 'Standard')
    doc.pageBreak()
    
    doc.addParagraph(_("Tables"), "Heading 1")
    doc.addParagraph(_("We can create tables too, for example with size 11pt:"), "Standard")
    table_data=[
        [_("Concept"), _("Value") ], 
        [_("Text"), _("This is a text")], 
#        [_("Datetime"), datetime.now()], 
#        [_("Date"), date.today()], 
#        [_("Decimal"), Decimal("12.121")], 
#        [_("Currency"), Currency(12.12, "EUR")], 
#        [_("Percentage"), Percentage(1, 3)], 
    ]
    
    
    doc.addTableParagraph(table_data)
    doc.pageBreak()

    doc.addParagraph(_("Lists and numbered lists"), "Heading 2") 
    doc.addParagraph(_("Simple list"), "Standard")
    doc.addListPlain([
        "Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. ", 
        "Adios", 
        "Bienvenido"
    ])       


    doc.pageBreak()
    doc.addParagraph(_("Images"), "Heading 1")
    
    l=[]
    l.append( _("Este es un ejemplo de imagen as char: "))
    l.append(doc.textcontentImage(pkg_resources.resource_filename(__name__, 'images/crown.png'), 1000, 1000, "AS_CHARACTER"))
    l.append(". Ahora sigo escribiendo sin problemas.")
    doc.addParagraphComplex(l, "Standard")

    l=[]
    l.append( _("As you can see, I can reuse it one hundred times. File size will not be increased because I used reference names."))
    for i in range(100):
        l.append(doc.textcontentImage(pkg_resources.resource_filename(__name__, 'images/crown.png'), 500, 500, "AS_CHARACTER"))
    doc.addParagraphComplex(l, "Standard")


    doc.addParagraph(_("The next paragraph is generated with the illustration method"), "Standard")
    doc.addImageParagraph([pkg_resources.resource_filename(__name__, 'images/crown.png')]*5, 2500, 1500, "Illustration")

    
    doc.save()
    doc.export_docx()
    doc.export_pdf()
    doc.close()
    return "demo_odt_standard took {}".format(datetime.now()-doc.init)
    

if __name__ == "__main__":
    main()
