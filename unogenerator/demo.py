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
from unogenerator.unogenerator import ODT_Standard, ODS_Standard
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
            for language in ['es', 'en']:
                remove_without_errors(f"unogenerator_documentation_{language}.odt")
                remove_without_errors(f"unogenerator_documentation_{language}.docx")
                remove_without_errors(f"unogenerator_documentation_{language}.pdf")
                remove_without_errors(f"unogenerator_example_{language}.ods")
                remove_without_errors(f"unogenerator_example_{language}.xlsx")
                remove_without_errors(f"unogenerator_example_{language}.pdf")

    if args.create==True:
        start=datetime.now()
        futures=[]
        with ProcessPoolExecutor(max_workers=cpu_count()+1) as executor:
            for language in ['es', 'en']:
                futures.append(executor.submit(demo_ods_standard, language))
                futures.append(executor.submit(demo_odt_standard, language))

        for future in as_completed(futures):
            print(future.result())
        print("All process took {}".format(datetime.now()-start))

       
def demo_ods_standard(language):
    if language=="en":
        lang1=gettext.install('unogenerator', 'badlocale')
    else:
        lang1=gettext.translation('unogenerator', 'unogenerator/locale', languages=[language])
        lang1.install()
    
    doc=ODS_Standard()
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
    doc.save(f"unogenerator_example_{language}.ods")
    doc.export_xlsx(f"unogenerator_example_{language}.xlsx")
    doc.export_pdf(f"unogenerator_example_{language}.pdf")
    doc.close()

    return f"unogenerator_example_{language}.ods took {datetime.now()-doc.init}"
    
    
def demo_odt_standard(language):
    if language=="en":
        lang1=gettext.install('unogenerator', 'badlocale')
    else:
        lang1=gettext.translation('unogenerator', 'unogenerator/locale', languages=[language])
        lang1.install()

    doc=ODT_Standard()
    doc.setMetadata(
        _("UnoGenerator documentation"),  
        _("Unogenerator python module documentation"), 
        "Turulomio", 
        _(f"This file have been generated with UnoGenerator-{__version__}. You can see UnoGenerator main page in http://github.com/turulomio/unogenerator"), 
        ["unogenerator", "demo", "files"]
    )
    doc.addParagraph(_("UnoGenerator documentation"), "Title")
    doc.addParagraph(_(f"Version: {__version__}"), "Subtitle")
    
    doc.addParagraph(_("Introduction"),  "Heading 1")
    
    doc.addParagraph(
        _("UnoGenerator uses Libreoffice UNO API python bindings to generate documents.") +" " +
        _("So in order to use, you need to launch a --headless libreoffice instance.") + " "+
        _("You can easily launch server.sh script with bash."), 
        "Standard"
    )

    doc.addParagraph(
        _("UnoGenerator has standard templates to help you with edition, although you can use your own templates.") +" " + 
        _("You can edit this one or create your own.")  +" " +
        _("This document has been created with 'standard.odt' files that you can find inside this python module."), 
        "Standard"
    )
        
    doc.addParagraph(_("Installation"), "Heading 2")
    doc.addParagraph(_("You can use pip to install this python package:") ,  "Standard")
    doc.addParagraph("""pip install unogenerator"""    , "Code")
    doc.addParagraph(_("Hello World example"), "Heading 2")
    doc.addParagraph(_("This is a Hello World example. You get the example in odt, docx and pdf formats:") ,  "Standard")
    doc.addParagraph("""from unogenerator import ODT_Standard
doc=ODT_Standard()
doc.addParagraph("Hello World", "Heading 1")
doc.addParagraph("Easy, isn't it","Standard")
doc.save("hello_world.odt")
doc.export_docx("hello_world.docx")
doc.export_pdf("hello_world.pdf")
doc.close()"""    , "Code")
    doc.pageBreak()
    
    doc.addParagraph(_("ODT"), "Heading 1")
    doc.addParagraph(
        _("ODT files can be quickly generated with UnoGenerator.") + " " + 
        _("There is a predefined template in code called 'standard.odt' to help you with edition."),  
        "Standard"
    )
    
    
    doc.addParagraph(_("Calling the ODT constructor"), "Heading 2")
    doc.addParagraph(_("You can call ODT constructor in this ways:") , "Standard")
        
    doc.addParagraph(
        _("ODT with standard template (Recomended).") + " " + 
        _("There is a predefined template in code called 'standard.odt', inside this python module, to help you with edition, although you can use your own ones.") +" "+
        _("With this mode you can create new documents"), 
        "Puntitos"
    )
    
    doc.addParagraph("""from unogenerator import ODT_Standard
doc=ODT_Standard()"""    , "Code")

        
    doc.addParagraph(
        _("ODT with template or file (Recomended).") + " " + 
        _("With this mode you can read your files to overwrite them or use your file as a new template to create new documents"), 
        "Puntitos"
    )
    
    doc.addParagraph("""from unogenerator import ODT
doc=ODT('yourdocument.odt')"""    , "Code")
    
    doc.addParagraph(
        _("ODT without template.") + " " + 
        _("With this mode you can write your files with Libreoffice default styles.") +" " +
        _("If you want to create new ones, you should write them using Libreoffice API code"), 
        "Puntitos"
    )
    
    
    doc.addParagraph("""from unogenerator import ODT
doc=ODT()"""    , "Code")
    
    doc.addParagraph(_("Styles"), "Heading 2")
    doc.addParagraph(
        _("To call default Libreoffice paragraph styles you must use their english name.") + " " + 
        _("You can see their names with this method:"), "Standard"
    )
    doc.addParagraph("""doc.print_styles()"""    , "Code")

    

    doc.addParagraph(_("Tables"), "Heading 2")
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
    
    
    doc.addTableParagraph(table_data, columnssize_percentages=[30, ])
    doc.pageBreak()

    doc.addParagraph(_("Lists and numbered lists"), "Heading 2") 
    doc.addParagraph(_("Simple list"), "Standard")
    doc.addListPlain([
        "Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. Prueba hola no. ", 
        "Adios", 
        "Bienvenido"
    ])       


    doc.pageBreak()
    doc.addParagraph(_("Images"), "Heading 2")
    
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

    doc.pageBreak()
    doc.addParagraph(_("ODS"), "Heading 1")
    doc.addParagraph("""    TTO READ
def demo_ods_standard_read():
    doc=ODS("unogenerator_ods_standard.ods")
    doc.setActiveSheet(0)
    print(doc.getValuesByRow("4", standard=True))
    print(doc.getValuesByRow("4", standard=False))
    doc.close()
    return "demo_ods_standard took {}".format(datetime.now()-doc.init)""",  "Standard")
    
    doc.save(f"unogenerator_documentation_{language}.odt")
    doc.export_docx(f"unogenerator_documentation_{language}.docx")
    doc.export_pdf(f"unogenerator_documentation_{language}.pdf")
    doc.close()
    return f"unogenerator_documentation_{language}.ods took {datetime.now()-doc.init}"
    

if __name__ == "__main__":
    main()
