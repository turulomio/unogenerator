## @namespace unogenerator.demo
## @brief Generate ODF example files
from uno import getComponentContext
getComponentContext()
import argparse
import gettext
import pkg_resources
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime
from multiprocessing import cpu_count

from unogenerator.commons import __version__, addDebugSystem, argparse_epilog
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
            futures.append(executor.submit(demo_odt))
            futures.append(executor.submit(demo_odt_standard))

        for future in as_completed(futures):
            print(future.result())
        print("All process took {}".format(datetime.now()-start))


def demo_ods():
    doc=ODS("unogenerator.ods")
    doc.createSheet("Example", 1)
    doc.addCell("A1", "Name")
    doc.addCell("B1", "Value")
    doc.addCell("A2",  "Float")
    doc.addCell("B2", float(12.121))
    doc.freezeAndSelect("B1")
    doc.removeSheet(0)
    doc.save()
    doc.export_xlsx()
    doc.close()

    return "demo_ods took {}".format(datetime.now()-doc.init)
   

    
def demo_odt():
    doc=ODT("unogenerator.odt")
    doc.save()
    doc.export_docx()
    doc.export_pdf()
    doc.close()
    return "demo_odt took {}".format(datetime.now()-doc.init)    
    
def demo_odt_standard():
    doc=ODT("unogenerator_standard.odt", "templates/standard.odt")
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
