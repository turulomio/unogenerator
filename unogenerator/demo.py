## @namespace unogenerator.demo
## @brief Generate ODF example files
from uno import getComponentContext
getComponentContext()
import argparse
from collections import OrderedDict
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime, date, timedelta
from gettext import translation
from logging import info
from multiprocessing import cpu_count
from importlib.resources import files
from pydicts.currency import Currency
from pydicts.percentage import Percentage
from unogenerator import ODT_Standard, ODS_Standard, __version__,  commons, ColorsNamed, Coord, LibreofficeServer
from unogenerator.helpers import helper_title_values_total_row,helper_title_values_total_column, helper_totals_row, helper_totals_from_range, helper_list_of_ordereddicts, helper_list_of_dicts, helper_list_of_ordereddicts_with_totals, helper_ods_sheet_stylenames, helper_split_big_listofrows
from tqdm import tqdm

try:
    t=translation('unogenerator', files("unogenerator") / 'locale')
    _=t.gettext
except:
    _=str

## If arguments is None, launches with sys.argc parameters. Entry point is toomanyfiles:main
## You can call with main(['--pretend']). It's equivalento to os.system('program --pretend')
## @param arguments is an array with parser arguments. For example: ['--argument','9']. 
def main(arguments=None):
    parser=argparse.ArgumentParser(prog='unogenerator', description=_('Create example files using unogenerator module'), epilog=commons.argparse_epilog(), formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--debug', help=_("Debug program information"), choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")
    group= parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--create', help="Create demo files", action="store_true",default=False)
    group.add_argument('--remove', help="Remove demo files", action="store_true", default=False)
    parser.add_argument('--type', help="Debug program information", choices=["COMMONSERVER_SEQUENTIAL","COMMONSERVER_CONCURRENT_PROCESS","COMMONSERVER_CONCURRENT_THREADS",  "SEQUENTIAL",  "CONCURRENT_PROCESS",  "CONCURRENT_THREADS"],  default="COMMONSERVER_CONCURRENT_PROCESS")
    args=parser.parse_args(arguments)
    commons.addDebugSystem(args.debug)
        
    if args.remove==True:
            for language in ['es', 'en']:
                commons.remove_without_errors(f"unogenerator_documentation_{language}.odt")
                commons.remove_without_errors(f"unogenerator_documentation_{language}.docx")
                commons.remove_without_errors(f"unogenerator_documentation_{language}.pdf")
                commons.remove_without_errors(f"unogenerator_example_{language}.ods")
                commons.remove_without_errors(f"unogenerator_example_{language}.xlsx")
                commons.remove_without_errors(f"unogenerator_example_{language}.pdf")

#    if args.create==True:
#        start=datetime.now()
#        futures=[]
#        with ProcessPoolExecutor(max_workers=cpu_count()+1) as executor:
#            for language in ['es', 'en']:
#                futures.append(executor.submit(demo_ods_standard, language))
#                futures.append(executor.submit(demo_odt_standard, language))
#
#        for future in as_completed(futures):
#            future.result()
#        print(_("All process took {}").format(datetime.now()-start))
#
#
#def main_concurrent(arguments=None):
#    parser=argparse.ArgumentParser(prog='unogenerator', description=_('Create example files using unogenerator module'), epilog=commons.argparse_epilog(), formatter_class=argparse.RawTextHelpFormatter)
#    parser.add_argument('--version', action='version', version=__version__)
#    parser.add_argument('--debug', help="Debug program information", choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")
#    group= parser.add_mutually_exclusive_group(required=True)
#    group.add_argument('--create', help="Create demo files", action="store_true",default=False)
#    group.add_argument('--remove', help="Remove demo files", action="store_true", default=False)
#    parser.add_argument('--loops', help="Loops of documentation jobs", action="store", default=30,  type=int)
#    parser.add_argument('--instances', help="Loops of documentation jobs", action="store", default=4,  type=int)
#    parser.add_argument('--nocommonserver', help="If true launches a server for each document", action="store_true", default=False)
#    args=parser.parse_args(arguments)
#
#    commons.addDebugSystem(args.debug)
#
#    if args.remove==True:
#            for i in range(args.loops):
#                for language in ['es', 'en']:
#                    commons.remove_without_errors(f"unogenerator_documentation_{language}.{i}.odt")
#                    commons.remove_without_errors(f"unogenerator_documentation_{language}.{i}.docx")
#                    commons.remove_without_errors(f"unogenerator_documentation_{language}.{i}.pdf")
#                    commons.remove_without_errors(f"unogenerator_example_{language}.{i}.ods")
#                    commons.remove_without_errors(f"unogenerator_example_{language}.{i}.xlsx")
#                    commons.remove_without_errors(f"unogenerator_example_{language}.{i}.pdf")

    if args.create==True:
        start=datetime.now()
        instances=3
        languages=['es', 'en', 'es', 'en']
        total_documents=len(languages)*2
        
        if args.type=="CONCURRENT_PROCESS":            
            futures=[]
            print(_("Launching demo with {0} workers without common server using concurrent processes").format(instances))

            with ProcessPoolExecutor(max_workers=instances) as executor:
                with tqdm(total=total_documents) as progress:
                    for language in languages:
                        future=executor.submit(demo_ods_standard, language, "", None)
                        future.add_done_callback(lambda p: progress.update())
                        futures.append(future)
                        future=executor.submit(demo_odt_standard, language, "",  None)
                        future.add_done_callback(lambda p: progress.update())
                        futures.append(future)

                    for future in as_completed(futures):
                        future.result()

            results = []
            for future in futures:
                result = future.result()
                results.append(result)
                
        elif args.type=="COMMONSERVER_SEQUENTIAL":
            with LibreofficeServer() as server:
                print(_("Launching concurrent demo with one commons server sequentially").format(args.instances))
                with tqdm(total=total_documents) as progress:
                    for language in languages:
                        demo_ods_standard(language, "", server)
                        progress.update()
                        demo_odt_standard(language, "",  server)       
                        progress.update()     
            
        print(_("All process took {}".format(datetime.now()-start)))

       
def demo_ods_standard(language, suffix="", server=None):
    lang1=translation('unogenerator', files("unogenerator") / 'locale', languages=[language])
    lang1.install()
    _=lang1.gettext
    
    with ODS_Standard(server=server) as doc:
        doc.setMetadata(
            _("UnoGenerator ODS example"),  
            _("Demo with ODS class"), 
            "Turulomio", 
            _("This file have been generated with UnoGenerator-{0}. You can see UnoGenerator main page in https://github.com/turulomio/unogenerator").format(__version__), 
            ["unogenerator", "demo", "files"]
        )
        doc.createSheet("Styles")
        doc.setColumnsWidth([3.5, 5, 2, 2, 2, 2, 2, 5, 5, 3, 3])
        
        doc.setSheetStyle("Portrait")
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
            doc.addCellWithStyle(Coord("A2").addRow(row), color_str, color_key, "Bold")
            doc.addCellWithStyle(Coord("B2").addRow(row), datetime.now(), color_key, "Datetime")
            doc.addCellWithStyle(Coord("C2").addRow(row), date.today(), color_key, "Date")
            doc.addCellWithStyle(Coord("D2").addRow(row), pow(-1, row)*-10000000, color_key, "Integer")
            doc.addCellWithStyle(Coord("E2").addRow(row), Currency(pow(-1, row)*12.56, "EUR"), color_key, "EUR")
            doc.addCellWithStyle(Coord("F2").addRow(row), Currency(pow(-1, row)*12345.56, "USD"), color_key, "USD")
            doc.addCellWithStyle(Coord("G2").addRow(row), Percentage(pow(-1, row)*1, 3), color_key,  "Percentage")
            doc.addCellWithStyle(Coord("H2").addRow(row), pow(-1, row)*123456789.121212, color_key, "Float6")
            doc.addCellWithStyle(Coord("I2").addRow(row), pow(-1, row)*-12.121212, color_key, "Float2")
            doc.addCellWithStyle(Coord("J2").addRow(row), (datetime.now()+timedelta(seconds=3600*12*row)).time(), color_key, "Time")
            doc.addCellWithStyle(Coord("K2").addRow(row), bool(row%2), color_key, "Bool")

        doc.addCellWithStyle(Coord("E2").addRow(row+1),f"=sum(E2:{Coord('E2').addRow(row).string()})", ColorsNamed.GrayLight, "EUR" )
        doc.addCellMergedWithStyle("E15:K15", "Merge proof", ColorsNamed.Yellow, style="BoldCenter")
        doc.setComment("B14", "This is nice comment")
        
        doc.freezeAndSelect("B2")
        
        ## List of rows
        doc.createSheet("List of rows or columns")
        
                
        doc.addCellMergedWithStyle("A1:C1","List of rows with helper_totals_row", ColorsNamed.Orange, "BoldCenter")
        range_=doc.addListOfRowsWithStyle("A2", [[1,2,3],[4,5,6],[7,8,9]], ColorsNamed.White)
        helper_totals_row(doc, range_.c_start.addRowCopy(range_.numRows()), ["#SUM"]*3, styles=None, row_from="2", row_to="4")
        
        doc.addCellMergedWithStyle("A8:C8","List of columns with helper_totals_row", ColorsNamed.Orange, "BoldCenter")
        range_=doc.addListOfColumnsWithStyle("A9", [[1,2,3],[4,5,6],[7,8,9]], ColorsNamed.White)
        helper_totals_row(doc, range_.c_start.addRowCopy(range_.numRows()), ["#SUM"]*3, styles=None, row_from="9", row_to="11")

        doc.addCellMergedWithStyle("A15:E15","List of rows with helper_totals_from_range in rows and columns", ColorsNamed.Orange, "BoldCenter")
        range_=doc.addListOfRowsWithStyle("A16", [["A",12000,2,3, 6],["B",1020,5,6, 7],["C",20404,8,9, 8]], ColorsNamed.White)
        helper_totals_from_range(doc, range_, totalcolumns=True, totalrows=True)
        
        
        doc.addCellMergedWithStyle("A22:E22","List of rows with helper_totals_from_range in rows", ColorsNamed.Orange, "BoldCenter")
        range_=doc.addListOfRowsWithStyle("A23", [["A",12000,2,3, 6],["B",1020,5,6, 7],["C",20404,8,9, 8]], ColorsNamed.White)
        helper_totals_from_range(doc, range_.addColumnBefore(-1), totalcolumns=False, totalrows=True) #Removes one column to filter first alphanumerical column

        doc.addCellMergedWithStyle("A29:E29","List of rows with helper_totals_from_range in columns", ColorsNamed.Orange, "BoldCenter")
        range_=doc.addListOfRowsWithStyle("A30", [["A",12000,2,3, 6],["B",1020,5,6, 7],["C",20404,8,9, 8]], ColorsNamed.White)
        helper_totals_from_range(doc, range_.addColumnBefore(-1), totalcolumns=True, totalrows=False)
        
        doc.addCellMergedWithStyle("A35:E35","List of rows with helper_totals_from_range in rows showing", ColorsNamed.Orange, "BoldCenter")
        range_=doc.addListOfRowsWithStyle("A36", [["A",12000,2,3, 6],["B",1020,5,6, 7],["C",20404,8,9, 8]], ColorsNamed.White)
        helper_totals_from_range(doc, range_.addColumnBefore(-1), totalcolumns=False, totalrows=True, showing=True)

        doc.addCellMergedWithStyle("A42:E42","List of rows with helper_totals_from_range in columns showing", ColorsNamed.Orange, "BoldCenter")
        range_=doc.addListOfRowsWithStyle("A43", [["A",12000,2,3, 6],["B",1020,5,6, 7],["C",20404,8,9, 8]], ColorsNamed.White)
        helper_totals_from_range(doc, range_.addColumnBefore(-1), totalcolumns=True, totalrows=False, showing=True)

        

        ## HELPERS
        doc.createSheet("Helpers")
        doc.setSheetStyle("Portrait")
        doc.addCellMergedWithStyle("A1:E1","Helper values with total (horizontal)", ColorsNamed.Orange, "BoldCenter")
        helper_title_values_total_row(doc, "A2", "Suma 3", [1,2,3])

        doc.addCellMergedWithStyle("A4:A9","Helper values with total (vertical)", ColorsNamed.Orange, "VerticalBoldCenter")
        helper_title_values_total_column(doc, "B4", "Suma 3", [1,2,3, 4])

        
        doc.addCellMergedWithStyle("A23:B23","List of ordered dictionaries", ColorsNamed.Orange, "BoldCenter")
        lod=[]
        lod.append(OrderedDict({"Singer": "Elvis",  "Song": "Fever" }))
        lod.append(OrderedDict({"Singer": "Roy Orbison",  "Song": "Blue angel" }))
        helper_list_of_ordereddicts(doc, "A24",  lod, columns_header=1)
        
        doc.addCellMergedWithStyle("A28:B28","List of dictionaries", ColorsNamed.Orange, "BoldCenter")
        helper_list_of_dicts(doc, "A29",  lod, keys=["Song",  "Singer"])
        
        doc.addCellMergedWithStyle("A33:D33","List of ordered dictionaries one method with totals", ColorsNamed.Orange, "BoldCenter")
        lod=[]
        lod.append(OrderedDict({"Singer": "Elvis",  "Songs": 10000 , "Albums": 100}))
        lod.append(OrderedDict({"Singer": "Roy Orbison",  "Songs": 100,  "Albums": 20 }))
        helper_list_of_ordereddicts_with_totals(doc, "A34",  lod, columns_header=1)
        
        ##Sort
        doc.createSheet("Sort")
        l=[7, 3, 2, 5, 6, 0, 9, 4, 10]
        doc.addCellWithStyle("A1",  "Unsorted", ColorsNamed.Orange, "BoldCenter")
        doc.addCellWithStyle("B1", "Sorted ASC", ColorsNamed.Orange, "BoldCenter")
        doc.addCellWithStyle("C1", "Sorted DESC", ColorsNamed.Orange, "BoldCenter")
        doc.addColumnWithStyle("A2", l)
        doc.addColumnWithStyle("B2", l)
        doc.addColumnWithStyle("C2", l)
        doc.sortRange("B2:B10",  0)
        doc.sortRange("C2:C10",  0, False)
        
        ## Split big LOR
        lor=[]
        for i in range(1000):
            lor.append([i, _("String")+" "+ str(i), datetime.now()])
            
        helper_split_big_listofrows(doc, "Splits in 400 rows", lor, ["Integer", "String", "Datetime"], columns_width=[2, 5, 5],  max_rows=400)

        
        ## Sheet with all styles names
        helper_ods_sheet_stylenames(doc)

        doc.save(f"unogenerator_example_{language}{suffix}.ods")
        doc.export_xlsx(f"unogenerator_example_{language}{suffix}.xlsx")
        doc.export_pdf(f"unogenerator_example_{language}{suffix}.pdf")
    
    r= _("unogenerator_example_{0}{1}.ods took {2} in {3}").format(language, suffix, datetime.now()-doc.start, doc.server.port)
    info(r)
    return r
    
    
def demo_odt_standard(language, suffix="", server=None):
    lang1=translation('unogenerator', files("unogenerator") / 'locale', languages=[language])
    lang1.install()
    _=lang1.gettext

    with ODT_Standard(server=server) as doc:
        doc.setMetadata(
            _("UnoGenerator documentation"),  
            _("UnoGenerator python module documentation"), 
            "Turulomio", 
            _("This file have been generated with UnoGenerator-{0}. You can see UnoGenerator main page in https://github.com/turulomio/unogenerator").format(__version__), 
            ["unogenerator", "demo", "files"]
        )
        
        
        doc.addParagraph(_("UnoGenerator documentation"), "Title")
        doc.addParagraph(_("Version: {0}").format(__version__), "Subtitle")

        doc.addImageParagraph([files('unogenerator') / 'images/unogenerator.png', ], 4, 4, "Illustration", linked=False)


        
        doc.addParagraph(_("Introduction"),  "Heading 1")
        
        doc.addParagraph(
            _("UnoGenerator uses Libreoffice UNO API python bindings to generate documents.") +" " +
            _("So in order to use, you need to launch a --headless LibreOffice instance.") + " "+
            _("We recomend to launch then easily with 'unogenerator_start' script."), 
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
with ODT_Standard() as doc:
    doc.addParagraph("Hello World", "Heading 1")
    doc.addParagraph("Easy, isn't it","Standard")
    doc.save("hello_world.odt")
    doc.export_docx("hello_world.docx")
    doc.export_pdf("hello_world.pdf")"""    , "Code")
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
            [_("Concept"), _("Value"), _("Comment") ], 
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
        doc.addParagraph(_("Hyperlinks"), "Heading 2")
        
        doc.addString(_("If you want to go to Google, click on this "))
        doc.addStringHyperlink("link",  "https://www.google.com")
        doc.addString(". " + _("That's all folks!"), paragraphBreak=True)
        
        doc.addStringHyperlink("Other link",  "https://www.google.com",  paragraphBreak=True)
        doc.pageBreak()
        
        
        doc.addParagraph(_("HTML code"), "Heading 2")
        doc.addHTMLBlock("<ul><li>This is a html list</li></ul><p style='color:red;'>This is a html paragraph.</p>")
        doc.pageBreak()

        doc.addParagraph(_("Images"), "Heading 2")
        
        l=[]
        l.append( _("This is an 'image as char' example: "))
        l.append(doc.textcontentImage(files('unogenerator') / 'images/crown.png', 1, 1, "AS_CHARACTER", "PRIMERA", linked=False))
        l.append(". "+_("Now I keep writing without problems."))
        doc.addParagraphComplex(l, "Standard")
        
        l=[]
        l.append( _("This is an image loaded from bytes: "))
        
        with open(files('unogenerator') / 'images/crown.png', "rb") as f:
            bytes_crown=f.read()
        
        l.append(doc.textcontentImage(bytes_crown, 1, 1, "AS_CHARACTER", "d", linked=False))
        doc.addParagraphComplex(l, "Standard")

        l=[]
        l.append( _("As you can see, I can reuse it one hundred times. File size will not be increased because I used reference names."))
        for i in range(100):
            l.append(doc.textcontentImage(files('unogenerator') / 'images/crown.png', 0.5,  0.5, "AS_CHARACTER", linked=False))
        doc.addParagraphComplex(l, "Standard")


        doc.addParagraph(_("The next paragraph is generated with the illustration method"), "Standard")
        doc.addImageParagraph([files('unogenerator') / 'images/crown.png']*5, 2.5, 1.5, "Illustration", linked=False)
        
        doc.addParagraph(_("The next paragraph is generated with the illustration method"), "Standard")
        doc.addImageParagraph([files('unogenerator') / 'images/crown.png']*5, 2.5, 1.5, "Illustration", linked=False)
        
        
        doc.addParagraph(_("You can play with image width and height:"), "Standard")
        l=[]
        l.append(doc.textcontentImage(files('unogenerator') / 'images/icons.jpg', None, None,  "AS_CHARACTER", "PRIMERA", linked=False))
        l.append(" Image default size. Height and width are set to None.")
        doc.addParagraphComplex(l, "Standard")
        l=[]
        l.append(doc.textcontentImage(files('unogenerator') / 'images/icons.jpg', 3,  None,  "AS_CHARACTER", "PRIMERA", linked=False))
        l.append(" Image width 3cm of width and height automatically set.")
        doc.addParagraphComplex(l, "Standard")
        l=[]
        l.append(doc.textcontentImage(files('unogenerator') / 'images/icons.jpg',  None, 3,  "AS_CHARACTER", "PRIMERA", linked=False))
        l.append(" Image height 3cm of width and width automatically set.")
        doc.addParagraphComplex(l, "Standard")
        l=[]
        l.append(doc.textcontentImage(files('unogenerator') / 'images/icons.jpg', 3, 3,  "AS_CHARACTER", "PRIMERA", linked=False))
        l.append(" Image width and height set to 3cm.")
        doc.addParagraphComplex(l, "Standard")
        
        
        doc.addParagraph(_("You can trim image border white space when needed:"), "Standard")
        
        l=[]
        l.append("This is a camerawith gray space around it: ")
        l.append(doc.textcontentImage(files('unogenerator') / 'images/Imagewithborder.png', 1, None))
        l.append(". ")
        l.append(_("You can trim that gray space when needed with 'bytes_after_trim_image' method. To use this method you need Imagemagick installed to use 'convert' command. This is the result: "))
        bytes_=commons.bytes_after_trim_image(files('unogenerator') / 'images/Imagewithborder.png', "png")
        l.append(doc.textcontentImage(bytes_, 1, None,))
        doc.addParagraphComplex(l, "Standard")
        
        doc.addParagraph(_("Search and Replace"), "Heading 2")
        doc.addParagraph(_("Below this paragraph is a paragraph with a % REPLACEME % (Without white spaces) text and it's going to be replaced after all document is been generated"), "Standard")
        doc.addParagraph("%REPLACEME%", "Standard")

        doc.pageBreak()
        doc.addParagraph(_("ODS"), "Heading 1")

        doc.addParagraph(_("ODS 'Hello World' example"), "Heading 2")
        doc.addParagraph(_("This is a Hello World example. You'll get the example in ods, xlsx and pdf formats:") ,  "Standard")
        doc.addParagraph("""from unogenerator import ODS_Standard
with ODS_Standard() as doc:
    doc.addCellMergedWithStyle("A1:E1", "Hello world", style="BoldCenter")
    doc.save("hello_world.ods")
    doc.export_xlsx("hello_world.xlsx")
    doc.export_pdf("hello_world.pdf")"""    , "Code")
        
        doc.find_and_replace("%REPLACEME%", _("This paragraph was set at the end of the code after a find and replace command."))
        doc.paragraphBreak()
        doc.addParagraph(_("This paragraph was set after replacement."), "Standard")
        doc.pageBreak()
        doc.addParagraph(_("This paragraph was set after a page break."), "Standard")
        doc.pageBreak("Landscape")
        doc.addParagraph(_("This paragraph was set after a page break with Landscape style."), "Standard")
        
        
    #    doc.find_and_delete_until_the_end_of_document("This paragraph was set after replacement.")

        doc.addParagraph(_("This is a pair of brackets ()."), "Standard")
        doc.findall_and_replace(" ().", " )(.", True)
        
        doc.addParagraph(_("NOW)("), "Standard")
        doc.addParagraph(_("This is a set of symbols: .,:;?ºª-/()."), "Standard")
        doc.findall_and_replace(".,:;?ºª-/().", ".,:;?ºª-/(). REPLACED", True)
        doc.addParagraph(_("NOW)("), "Standard")
        
        
        
        doc.save(f"unogenerator_documentation_{language}{suffix}.odt")
        doc.export_docx(f"unogenerator_documentation_{language}{suffix}.docx")
        doc.export_pdf(f"unogenerator_documentation_{language}{suffix}.pdf")

    r= _("unogenerator_documentation_{0}{1}.ods took {2} in {3}").format(language, suffix, datetime.now()-doc.start, doc.server.port)
    info(r)
    return r

