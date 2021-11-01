## @namespace unogenerator.demo
## @brief Generate ODF example files
from uno import getComponentContext
getComponentContext()
import argparse
from collections import OrderedDict
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime, date, timedelta
from gettext import translation
from multiprocessing import cpu_count
from pkg_resources import resource_filename

from unogenerator.commons import __version__, addDebugSystem, argparse_epilog, ColorsNamed, Coord as C, next_port, get_from_process_numinstances_and_firstport
from unogenerator.reusing.currency import Currency
from unogenerator.reusing.percentage import Percentage
from unogenerator.unogenerator import ODT_Standard, ODS_Standard
from unogenerator.helpers import helper_title_values_total_row,helper_title_values_total_column, helper_totals_row, helper_totals_column, helper_totals_from_range, helper_list_of_ordereddicts, helper_list_of_dicts, helper_list_of_ordereddicts_with_totals
from os import remove
from tqdm import tqdm

try:
    t=translation('unogenerator', resource_filename("unogenerator","locale"))
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
    parser.add_argument('--debug', help=_("Debug program information"), choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")
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
            future.result()
        print(_("All process took {}".format(datetime.now()-start)))


def main_concurrent(arguments=None):
    parser=argparse.ArgumentParser(prog='unogenerator', description=_('Create example files using unogenerator module'), epilog=argparse_epilog(), formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--debug', help="Debug program information", choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")
    group= parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--create', help="Create demo files", action="store_true",default=False)
    group.add_argument('--remove', help="Remove demo files", action="store_true", default=False)
    parser.add_argument('--workers', help=_("Number of workers to process this script. Default 4"), action="store", default=4,  type=int)
    parser.add_argument('--loops', help="Loops of documentation jobs", action="store", default=30,  type=int)
    args=parser.parse_args(arguments)

    num_instances, first_port=get_from_process_numinstances_and_firstport()
    addDebugSystem(args.debug)

    if args.remove==True:
            for i in range(args.loops):
                remove_without_errors(f"unogenerator_documentation_en.{i}.odt")
                remove_without_errors(f"unogenerator_documentation_en.{i}.docx")
                remove_without_errors(f"unogenerator_documentation_en.{i}.pdf")
                remove_without_errors(f"unogenerator_example_en.{i}.ods")
                remove_without_errors(f"unogenerator_example_en.{i}.xlsx")
                remove_without_errors(f"unogenerator_example_en.{i}.pdf")

    if args.create==True:
        print(_(f"Launching concurrent demo with {args.workers} workers to a daemon with {num_instances} instances from {first_port} port"))

        start=datetime.now()
        futures=[]
        port=first_port
        with ProcessPoolExecutor(max_workers=args.workers) as executor:
            with tqdm(total=args.loops*2) as progress:
                for i in range(args.loops):
                    port=next_port(port, first_port, num_instances)
                    future=executor.submit(demo_ods_standard, 'en', port, f".{i}")
                    future.add_done_callback(lambda p: progress.update())
                    futures.append(future)
                    port=next_port(port, first_port, num_instances)
                    future=executor.submit(demo_odt_standard, 'en', port, f".{i}")
                    future.add_done_callback(lambda p: progress.update())
                    futures.append(future)

                for future in as_completed(futures):
                    future.result()
            


        results = []
        for future in futures:
            result = future.result()
            results.append(result)
        print(_("All process took {}".format(datetime.now()-start)))

       
def demo_ods_standard(language, port=2002, suffix="",):
    if language=="en":
        lang1=translation('unogenerator' , resource_filename("unogenerator","locale"), languages=[language])
        lang1.install()
    else:
        lang1=translation('unogenerator', resource_filename("unogenerator","locale"), languages=[language])
        lang1.install()
    ##print(_("My language is "),language,lang1)
    
    doc=ODS_Standard(port)
    doc.setMetadata(
        _("UnoGenerator ODS example"),  
        _("Demo with ODS class"), 
        "Turulomio", 
        _(f"This file have been generated with UnoGenerator-{__version__}. You can see UnoGenerator main page in http://github.com/turulomio/unogenerator"), 
        ["unogenerator", "demo", "files"]
    )
    doc.createSheet("Styles")
    doc.setColumnsWidth([3.5, 5, 2, 2, 2, 2, 2, 5, 5, 3, 3])
    
    
    doc.setCellName("A1",  "MYNAME")
    
    doc.addCellWithStyle("A1", _("Style name"), ColorsNamed.Orange, "BoldCenter")
    doc.addCellWithStyle("B1", _("Date and time"), ColorsNamed.Orange, "BoldCenter")
    doc.addCellWithStyle("C1", _("Date"), ColorsNamed.Orange, "BoldCenter")
    doc.addCellWithStyle("D1", _("Integer"), ColorsNamed.Orange, "BoldCenter")
    doc.addCellWithStyle("E1", _("Euros"), ColorsNamed.Orange, "BoldCenter")
    doc.addCellWithStyle("F1", _("Dollars"), ColorsNamed.Orange, "BoldCenter")
    doc.addCellWithStyle("G1", _("Percentage"), ColorsNamed.Orange, "BoldCenter")
    doc.addCellWithStyle("H1", _("Number with 2 decimals"), ColorsNamed.Orange, "BoldCenter")
    doc.addCellWithStyle("I1", _("Number with 6 decimals"), ColorsNamed.Orange, "BoldCenter")
    doc.addCellWithStyle("J1", _("Time"), ColorsNamed.Orange, "BoldCenter")
    doc.addCellWithStyle("K1", _("Boolean"), ColorsNamed.Orange, "BoldCenter")
    colors_list=([a for a in dir(ColorsNamed()) if not a.startswith('__')])
    for row, color_str in enumerate(colors_list):
        color_key=getattr(ColorsNamed(), color_str)
        doc.addCellWithStyle(C("A2").addRow(row), color_str, color_key, "Bold")
        doc.addCellWithStyle(C("B2").addRow(row), datetime.now(), color_key, "Datetime")
        doc.addCellWithStyle(C("C2").addRow(row), date.today(), color_key, "Date")
        doc.addCellWithStyle(C("D2").addRow(row), pow(-1, row)*-10000000, color_key, "Integer")
        doc.addCellWithStyle(C("E2").addRow(row), Currency(pow(-1, row)*12.56, "EUR"), color_key, "EUR")
        doc.addCellWithStyle(C("F2").addRow(row), Currency(pow(-1, row)*12345.56, "USD"), color_key, "USD")
        doc.addCellWithStyle(C("G2").addRow(row), Percentage(pow(-1, row)*1, 3), color_key,  "Percentage")
        doc.addCellWithStyle(C("H2").addRow(row), pow(-1, row)*123456789.121212, color_key, "Float6")
        doc.addCellWithStyle(C("I2").addRow(row), pow(-1, row)*-12.121212, color_key, "Float2")
        doc.addCellWithStyle(C("J2").addRow(row), (datetime.now()+timedelta(seconds=3600*12*row)).time(), color_key, "Time")
        doc.addCellWithStyle(C("K2").addRow(row), bool(row%2), color_key, "Bool")

    doc.addCellWithStyle(C("E2").addRow(row+1),f"=sum(E2:{C('E2').addRow(row).string()})", ColorsNamed.GrayLight, "EUR" )
    doc.addCellMergedWithStyle("E15:K15", "Prueba de merge", ColorsNamed.Yellow, style="BoldCenter")
    doc.setComment("B14", "This is nice comment")
    
    doc.freezeAndSelect("B2")
    
##    # Un comment to see objjects from cell
##    for i in range(10):
##        o=doc.getValue(Coord_from_index(i, 3))
##        print(o, o.__class__)    

    ## HELPERS
    doc.createSheet("Helpers")
    doc.addCellMergedWithStyle("A1:E1","Helper values with total (horizontal)", ColorsNamed.Orange, "BoldCenter")
    helper_title_values_total_row(doc, "A2", "Suma 3", [1,2,3])

    doc.addCellMergedWithStyle("A4:E4","Helper values with total (vertical)", ColorsNamed.Orange, "BoldCenter")
    helper_title_values_total_column(doc, "A5", "Suma 3", [1,2,3])
    
    doc.addCellMergedWithStyle("A11:C11","List of rows", ColorsNamed.Orange, "BoldCenter")
    doc.addListOfRowsWithStyle("A12", [[1,2,3],[4,5,6],[7,8,9]], ColorsNamed.White)
    
    doc.addCellMergedWithStyle("E11:G11","List of columns", ColorsNamed.Orange, "BoldCenter")
    doc.addListOfColumnsWithStyle("E12", [[1,2,3],[4,5,6],[7,8,9]], ColorsNamed.White)

    helper_totals_row(doc, "A17", ["#SUM"]*3, styles=None, row_from="12", row_to="15")
    helper_totals_column(doc, "I12", ["#SUM"]*3, styles=None, column_from="E", column_to="G")


    doc.addCellMergedWithStyle("A19:D19","List of rows with totals", ColorsNamed.Orange, "BoldCenter")
    doc.addListOfRowsWithStyle("A20", [["A",12000,2,3],["B",1020,5,6],["C",20404,8,9]], ColorsNamed.White)
    helper_totals_from_range(doc, "B20:D22", totalcolumns=True, totalrows=True)

    
    doc.addCellMergedWithStyle("A25:B25","List of ordered dictionaries", ColorsNamed.Orange, "BoldCenter")
    lod=[]
    lod.append(OrderedDict({"Singer": "Elvis",  "Song": "Fever" }))
    lod.append(OrderedDict({"Singer": "Roy Orbison",  "Song": "Blue angel" }))
    helper_list_of_ordereddicts(doc, "A26",  lod, columns_header=1)
    
    doc.addCellMergedWithStyle("A30:B30","List of dictionaries", ColorsNamed.Orange, "BoldCenter")
    helper_list_of_dicts(doc, "A31",  lod, keys=["Song",  "Singer"])
    
    doc.addCellMergedWithStyle("A35:D35","List of ordered dictionaries one method with totals", ColorsNamed.Orange, "BoldCenter")
    lod=[]
    lod.append(OrderedDict({"Singer": "Elvis",  "Songs": 10000 , "Albums": 100}))
    lod.append(OrderedDict({"Singer": "Roy Orbison",  "Songs": 100,  "Albums": 20 }))
    helper_list_of_ordereddicts_with_totals(doc, "A36",  lod, columns_header=1)

    doc.save(f"unogenerator_example_{language}{suffix}.ods")
    doc.export_xlsx(f"unogenerator_example_{language}{suffix}.xlsx")
    doc.export_pdf(f"unogenerator_example_{language}{suffix}.pdf")
    doc.close()
    
    r= _(f"unogenerator_example_{language}{suffix}.ods took {datetime.now()-doc.init} in {port}")
    print(r)
    return r
    
    
def demo_odt_standard(language, port=2002, suffix=""):
    if language=="en":
        lang1=translation('unogenerator', resource_filename("unogenerator","locale"), languages=[language])
        lang1.install()
    else:
        lang1=translation('unogenerator', resource_filename("unogenerator","locale"), languages=[language])
        lang1.install()

    doc=ODT_Standard(port)
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
        _("You can easily launch unogenerator_start script with bash."), 
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
    doc.addParagraph(_("ODT 'Hello World' example"), "Heading 2")
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
        "BulletsLevel1"
    )
    
    doc.addParagraph("""from unogenerator import ODT_Standard
doc=ODT_Standard()"""    , "Code")

    doc.addParagraph(
        _("ODT with template or file (Recomended).") + " " + 
        _("With this mode you can read your files to overwrite them or use your file as a new template to create new documents"), 
        "BulletsLevel1"
    )
    
    doc.addParagraph("""from unogenerator import ODT
doc=ODT('yourdocument.odt')"""    , "Code")
    
    doc.addParagraph(
        _("ODT without template.") + " " + 
        _("With this mode you can write your files with Libreoffice default styles.") +" " +
        _("If you want to create new ones, you should write them using Libreoffice API code"), 
        "BulletsLevel1"
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
    table_data=[
        [_("Concept"), _("Value"), _("Commnet") ], 
        [_("Text"), _("This is a text"), _("Good")], 
        [_("Datetime"), datetime.now(), _("Good")], 
        [_("Date"), date.today(), _("Good")], 
        [_("Float"),  12.121, _("Good")], 
        [_("Currency"), Currency(-12.12, "EUR"), _("Good")], 
        [_("Percentage"), Percentage(1, 3), _("Good")], 
    ]
    
    columnspercentages=[15, 70, 15 ]
    doc.addParagraph(_("We can create tables with diferent font sizes and formats:") + str(columnspercentages), "Standard")
    doc.addTableParagraph(table_data, columnssize_percentages=columnspercentages, style="Table1")
    
    doc.addTableParagraph(table_data, columnssize_percentages=[30, 40,30],  size=6, style="Table1")
    

    doc.addParagraph(_("Lists and numbered lists"), "Heading 2") 
    doc.addParagraph(_("Simple list"), "BulletsLevel1")
    doc.addParagraph(_("Simple list"), "BulletsLevel2")
    doc.addParagraph(_("Simple list"), "BulletsLevel2")
    doc.addParagraph(_("Simple list"), "BulletsLevel1")
    doc.addParagraph(_("Simple list"), "BulletsLevel2")
    doc.addParagraph(_("Simple list"), "BulletsLevel1")


    doc.pageBreak()
    doc.addParagraph(_("Images"), "Heading 2")
    
    l=[]
    l.append( _("Este es un ejemplo de imagen as char: "))
    l.append(doc.textcontentImage(resource_filename(__name__, 'images/crown.png'), 1, 1, "AS_CHARACTER", "PRIMERA", linked=True))
    l.append(". Ahora sigo escribiendo sin problemas.")
    doc.addParagraphComplex(l, "Standard")

    l=[]
    l.append( _("As you can see, I can reuse it one hundred times. File size will not be increased because I used reference names."))
    for i in range(100):
        l.append(doc.textcontentImage(resource_filename(__name__, 'images/crown.png'), 0.5,  0.5, "AS_CHARACTER", linked=True))
    doc.addParagraphComplex(l, "Standard")


    doc.addParagraph(_("The next paragraph is generated with the illustration method"), "Standard")
    doc.addImageParagraph([resource_filename(__name__, 'images/crown.png')]*5, 2.5, 1.5, "Illustration", linked=True)


    doc.addParagraph(_("Search and Replace"), "Heading 2")
    doc.addParagraph(_("Below this paragraph is a paragraph with a % REPLACEME % (Without white spaces) text and it's going to be replaced after all document is been generated"), "Standard")
    doc.addParagraph("%REPLACEME%", "Standard")

    doc.pageBreak()
    doc.addParagraph(_("ODS"), "Heading 1")

    doc.addParagraph(_("ODS 'Hello World' example"), "Heading 2")
    doc.addParagraph(_("This is a Hello World example. You'll get the example in ods, xlsx and pdf formats:") ,  "Standard")
    doc.addParagraph("""from unogenerator import ODS_Standard
doc=ODS_Standard()
doc.addCellMergedWithStyle("A1:E1", "Hello world", style="BoldCenter")
doc.save("hello_world.ods")
doc.export_xlsx("hello_world.xlsx")
doc.export_pdf("hello_world.pdf")
doc.close()"""    , "Code")
    
    doc.find_and_replace("%REPLACEME%", _("This paragraph was set at the end of the code after a find and replace command."))
    doc.paragraphBreak()
    doc.addParagraph(_("This paragraph was set after replacement."), "Standard")
    doc.pageBreak()
    doc.addParagraph(_("This paragraph was set after a page break."), "Standard")
    doc.pageBreak("Landscape")
    doc.addParagraph(_("This paragraph was set after a page break with Landscape style."), "Standard")
    
    
#    doc.find_and_delete_until_the_end_of_document("This paragraph was set after replacement.")
    
    
    doc.save(f"unogenerator_documentation_{language}{suffix}.odt")
    doc.export_docx(f"unogenerator_documentation_{language}{suffix}.docx")
    doc.export_pdf(f"unogenerator_documentation_{language}{suffix}.pdf")
    doc.close()
    r=_(f"unogenerator_documentation_{language}{suffix}.odt took {datetime.now()-doc.init} in {port}")
    print(r)
    return r

if __name__ == "__main__":
    main()
