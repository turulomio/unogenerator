from uno import getComponentContext

getComponentContext()
import argparse
from collections import OrderedDict
from concurrent.futures import ProcessPoolExecutor, as_completed,  ThreadPoolExecutor
from datetime import datetime, date, timedelta
from gettext import translation # Removed 'info'
import logging # Import logging module
from importlib.resources import files
from os import system
from pydicts.currency import Currency
from pydicts.percentage import Percentage
from unogenerator import ODT_Standard, ODS_Standard, __version__,  commons, ColorsNamed, Coord, LibreofficeServer, helpers, types, Range
from tqdm import tqdm

try:
    t=translation('unogenerator', files("unogenerator") / 'locale')
    _=t.gettext
except:
    _=str

type_choices=[ "SEQUENTIAL",  "CONCURRENT_PROCESS",  "CONCURRENT_THREADS", "COMMONSERVER_SEQUENTIAL","COMMONSERVER_CONCURRENT_PROCESS","COMMONSERVER_CONCURRENT_THREADS" ]

## If arguments is None, launches with sys.argc parameters. Entry point is toomanyfiles:main

logger = logging.getLogger(__name__) # Get logger for this module


lod_singers=[
    {"Singer": "Elvis",  "Songs": 10000 , "Albums": 100, "Best song": "Always on my mind"},
    {"Singer": "Roy Orbison",  "Songs": 100,  "Albums": 20, "Best song": "Crying"},
]

lod_widths = [
    OrderedDict({"Column 1": "Short", "Column 2": "This is a much longer string to measure", "Column 3": 100}),
    OrderedDict({"Column 1": "A medium string", "Column 2": "Short", "Column 3": 20000}),
    OrderedDict({"Column 1": "A very very very long string that should affect quantile 90", "Column 2": "Medium", "Column 3": 3}),
    ]

lod_singers_rows=len(lod_singers)
lod_singers_columns=len(lod_singers[0].keys())

lol_numbers=[
    ["One Two Three", "Four", "Ten"],
    ["One Two Three", "Four Two Three", "Ten"],
    ["One Two Three", "Four", "Ten Two Three"],
]

lol_numbers_headers=["A","B","C"]

lol_numbers_rows=len(lol_numbers)
lol_numbers_columns=len(lol_numbers[0])


lol_integers=[
    [1,2,3],
    [4,5,6],
    [7,8,9]
]

lol_thousands=[]
for i in range(1000):
    lol_thousands.append([i, _("String")+" "+ str(i), datetime.now()])

lol_thousands_rows=len(lol_thousands)
lol_thousands_columns=len(lol_thousands[0])



## You can call with main(['--pretend']). It's equivalento to os.system('program --pretend')
## @param arguments is an array with parser arguments. For example: ['--argument','9']. 
def demo(arguments=None):
    parser=argparse.ArgumentParser(prog='unogenerator', description=_('Create example files using unogenerator module'), epilog=commons.argparse_epilog(), formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--debug', help=_("Debug program information"), choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")
    group= parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--create', help="Create demo files", action="store_true",default=False)
    group.add_argument('--remove', help="Remove demo files", action="store_true", default=False)
    group.add_argument('--benchmark', help="Executes all types to compare its benchmark", action="store_true", default=False)
    parser.add_argument('--type', help="Debug program information", choices=type_choices,  default="COMMONSERVER_CONCURRENT_PROCESS")
    args=parser.parse_args(arguments)
    commons.addDebugSystem(args.debug)
    demo_command(args.create, args.remove, args.benchmark, args.type)

def demo_command(create, remove, benchmark, type):
    languages=['es', 'en',  'ro',  'fr']
        
    if benchmark is True:
        for type in type_choices:
            #demo_command(True, False,  False,  type)
            system(f"unogenerator_demo --create --type {type}")

    if remove==True:
            for language in languages:
                commons.remove_without_errors(f"unogenerator_documentation_{language}.odt")
                commons.remove_without_errors(f"unogenerator_documentation_{language}.docx")
                commons.remove_without_errors(f"unogenerator_documentation_{language}.pdf")
                commons.remove_without_errors(f"unogenerator_example_{language}.ods")
                commons.remove_without_errors(f"unogenerator_example_{language}.xlsx")
                commons.remove_without_errors(f"unogenerator_example_{language}.pdf")

    if create==True:
        start=datetime.now()
        instances=3
        total_documents=len(languages)*2
        
        if type=="CONCURRENT_PROCESS":            
            futures=[]
            print(_("Launching demo with {0} workers without common server using concurrent processes").format(instances))

            with ProcessPoolExecutor(max_workers=instances) as executor:
                with tqdm(total=total_documents) as progress:
                    for language in languages:
                        future=executor.submit(demo_ods_standard, language, None)
                        future.add_done_callback(lambda p: progress.update())
                        futures.append(future)
                        future=executor.submit(demo_odt_standard, language,  None)
                        future.add_done_callback(lambda p: progress.update())
                        futures.append(future)

                    for future in as_completed(futures):
                        future.result()

            results = []
            for future in futures:
                result = future.result()
                results.append(result)    

        elif type=="COMMONSERVER_CONCURRENT_PROCESS":            
            futures=[]
            print(_("Launching demo with {0} workers with common server using concurrent processes").format(instances))

            # Start a single LibreofficeServer in the main process.
            # This instance will manage the actual LibreOffice process.
            main_server = LibreofficeServer()
            main_server_port = main_server.port

            try:
                with ProcessPoolExecutor(max_workers=instances) as executor:
                    with tqdm(total=total_documents) as progress:
                            for language in languages:
                                # Pass only the port to the child processes.
                                # Each child process will create a new LibreofficeServer(port=main_server_port)
                                # which will connect to the main_server_port without starting a new LO process.
                                future=executor.submit(demo_ods_standard, language, main_server_port)
                                future.add_done_callback(lambda p: progress.update())
                                futures.append(future)
                                future=executor.submit(demo_odt_standard, language,  main_server_port)
                                future.add_done_callback(lambda p: progress.update())
                                futures.append(future)

                            for future in as_completed(futures):
                                future.result()

                results = []
                for future in futures:
                    result = future.result()
                    results.append(result)
            finally:
                main_server.stop() # Ensure the main server is stopped when done

        elif type=="CONCURRENT_THREADS":            
            futures=[]
            print(_("Launching demo with {0} workers without common server using concurrent threads").format(instances))

            with ThreadPoolExecutor(max_workers=instances) as executor:
                with tqdm(total=total_documents) as progress:
                    for language in languages:
                        future=executor.submit(demo_ods_standard, language, None)
                        future.add_done_callback(lambda p: progress.update())
                        futures.append(future)
                        future=executor.submit(demo_odt_standard, language,  None)
                        future.add_done_callback(lambda p: progress.update())
                        futures.append(future)

                    for future in as_completed(futures):
                        future.result()

                results = []
                for future in futures:
                    result = future.result()
                    results.append(result)

        elif type=="COMMONSERVER_CONCURRENT_THREADS":            
            futures=[]
            print(_("Launching demo with {0} workers with common server using concurrent threads").format(instances))

            with LibreofficeServer() as server: #FALLA POR PICCKING
                with ThreadPoolExecutor(max_workers=instances) as executor:
                    with tqdm(total=total_documents) as progress:
                        for language in languages:
                            future=executor.submit(demo_ods_standard, language, server)
                            future.add_done_callback(lambda p: progress.update())
                            futures.append(future)
                            future=executor.submit(demo_odt_standard, language,  server)
                            future.add_done_callback(lambda p: progress.update())
                            futures.append(future)

                        for future in as_completed(futures):
                            future.result()

                results = []
                for future in futures:
                    result = future.result()
                    results.append(result)

        elif type=="COMMONSERVER_SEQUENTIAL":
            with LibreofficeServer() as server:
                print(_("Launching demo with one common server sequentially"))
                with tqdm(total=total_documents) as progress:
                    for language in languages:
                        demo_ods_standard(language, server)
                        progress.update()
                        demo_odt_standard(language,  server)       
                        progress.update()
                        
        elif type=="SEQUENTIAL":
                print(_("Launching demo without one common server sequentially"))
                with tqdm(total=total_documents) as progress:
                    for language in languages:
                        demo_ods_standard(language, None)
                        progress.update()
                        demo_odt_standard(language,  None)       
                        progress.update()     
            
        print(_("All process took {}".format(datetime.now()-start)))

       
def demo_ods_standard(language, server):
    lang1=translation('unogenerator', files("unogenerator") / 'locale', languages=[language], fallback=True)
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
        
        demo_ods_sheet_styles(doc)
        demo_ods_sheet_sort(doc)
        demo_ods_sheet_word_wrap(doc)
        demo_ods_helpers_single(doc)
        demo_ods_block_from_lod(doc)
        demo_ods_block_from_lol(doc)
        demo_ods_block_from_lod_with_headers(doc)
        demo_ods_sheet_from_lod(doc)
        demo_ods_sheet_from_lol(doc)
        demo_ods_columns_width_modes(doc)
        demo_ods_sheet_split_with_big_lol(doc)
        helpers.sheet_stylenames(doc)

        doc.save(f"unogenerator_example_{language}.ods")
        doc.export_xlsx(f"unogenerator_example_{language}.xlsx")
        doc.export_pdf(f"unogenerator_example_{language}.pdf")
    
    r= _("unogenerator_example_{0}.ods took {1} in {2}").format(language, datetime.now()-doc.start, doc.server.port) # This is an application-level message
    logger.info(r)
    return r
    
    
def demo_odt_standard(language, server):
    lang1=translation('unogenerator', files("unogenerator") / 'locale', languages=[language], fallback=True)
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
            _("UnoGenerator make this unattended for you in each ODF (ODS and ODT) instance.") + " " +
            _("However if you wish, you can do this programmatically using LibreofficeServer class to reuse it in serveral ODF instances to improve documents generation speed."),
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
        
        
        
        doc.save(f"unogenerator_documentation_{language}.odt")
        doc.export_docx(f"unogenerator_documentation_{language}.docx")
        doc.export_pdf(f"unogenerator_documentation_{language}.pdf")

    r= _("unogenerator_documentation_{0}.ods took {1} in {2}").format(language, datetime.now()-doc.start, doc.server.port) # This is an application-level message
    logger.info(r)
    return r


def demo_ods_sheet_styles(doc):
    doc.createSheet("Styles")
    
    doc.setSheetStyle("Portrait")
    doc.setCellName("A1",  "MYNAME")


    headers=[_("Color name"), _("Hex"), _("Date and time"), _("Date"), _("Integer"), _("Euros"), _("Dollars"), _("Percentage"), _("Number with 2 decimals"), _("Number with 6 decimals"), _("Time"), _("Boolean")]
    doc.addRowWithStyle( "A1", headers, ColorsNamed.Orange, "BoldCenter")

    # Get colors and sort them by affinity (Hue)
    def rgb_to_hsv(rgb):
        r = ((rgb >> 16) & 0xff) / 255.0
        g = ((rgb >> 8) & 0xff) / 255.0
        b = (rgb & 0xff) / 255.0
        mx = max(r, g, b)
        mn = min(r, g, b)
        df = mx - mn
        if mx == mn:
            h = 0
        elif mx == r:
            h = (60 * ((g - b) / df) + 360) % 360
        elif mx == g:
            h = (60 * ((b - r) / df) + 120) % 360
        elif mx == b:
            h = (60 * ((r - g) / df) + 240) % 360
        if mx == 0:
            s = 0
        else:
            s = df / mx
        v = mx
        return h, s, v

    colors_list = [a for a in dir(ColorsNamed()) if not a.startswith('__')]
    # Decorate with hsv for sorting
    decorated = []
    for color_str in colors_list:
        val = getattr(ColorsNamed(), color_str)
        decorated.append((rgb_to_hsv(val), color_str, val))
    
    # Sort by Hue, then Saturation, then Value
    decorated.sort()
    
    for row, (hsv, color_str, color_key) in enumerate(decorated):
        hex_str = f"#{color_key:06X}"
        doc.addCellWithStyle(Coord("A2").addRow(row), color_str, color_key, "Bold")
        doc.addCellWithStyle(Coord("B2").addRow(row), hex_str, color_key, "Normal")
        doc.addCellWithStyle(Coord("C2").addRow(row), datetime.now(), color_key, "Datetime")
        doc.addCellWithStyle(Coord("D2").addRow(row), date.today(), color_key, "Date")
        doc.addCellWithStyle(Coord("E2").addRow(row), pow(-1, row)*-10000000, color_key, "Integer")
        doc.addCellWithStyle(Coord("F2").addRow(row), Currency(pow(-1, row)*12.56, "EUR"), color_key, "EUR")
        doc.addCellWithStyle(Coord("G2").addRow(row), Currency(pow(-1, row)*12345.56, "USD"), color_key, "USD")
        doc.addCellWithStyle(Coord("H2").addRow(row), Percentage(pow(-1, row)*1, 3), color_key,  "Percentage")
        doc.addCellWithStyle(Coord("I2").addRow(row), pow(-1, row)*123456789.121212, color_key, "Float6")
        doc.addCellWithStyle(Coord("J2").addRow(row), pow(-1, row)*-12.121212, color_key, "Float2")
        doc.addCellWithStyle(Coord("K2").addRow(row), (datetime.now()+timedelta(seconds=3600*12*row)).time(), color_key, "Time")
        doc.addCellWithStyle(Coord("L2").addRow(row), bool(row%2), color_key, "Bool")
   
    doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)
    doc.freezeAndSelect("C2")



def demo_ods_block_from_lod(doc):
        ## List of rows
        doc.createSheet("block_from_lod")
        helpers.block_from_lod(doc, "A1", lod_singers)
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

        doc.createSheet("block_from_lod title")
        helpers.block_from_lod(doc, "A1", lod_singers, title="block_from_lod with title")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

        doc.createSheet("block_from_lod other headers")
        helpers.block_from_lod(doc, "A1", lod_singers, columns_header=1, color_row_header=ColorsNamed.Red, title="block_from_lod (With other headers)")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

        doc.createSheet("block_from_lod empty")
        helpers.block_from_lod(doc, "A1", [], columns_header=1, color_row_header=ColorsNamed.Red, title="block_from_lod (Empty)")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

        doc.createSheet("block_from_lod column of totals")
        helpers.block_from_lod(doc, "A1", lod_singers, column_of_totals=True, title="block_from_lod (With total columns)", styles="Integer")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

        doc.createSheet("block_from_lod row of totals")
        helpers.block_from_lod(doc, "A1", lod_singers, row_of_totals=True, title="block_from_lod (With total rows)", styles="Integer")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)
        
        doc.createSheet("block_from_lod column of totals row of totals")
        helpers.block_from_lod(doc, "A1", lod_singers, column_of_totals=True, row_of_totals=True, title="block_from_lod (With total columns and rows)", styles="Integer")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)


def demo_ods_block_from_lod_with_headers(doc):
        # block_from_lod_with_headers
        doc.createSheet("block_from_lod_with_headers")
        helpers.block_from_lod_with_headers(doc, lod_singers, "A1", [
             ["Singer header", "Singer"],
             ["Song header", "Best song"]
        ], titulo="block_from_lod_with_headers")
        doc.setColumnsWidth(lod_singers, types.ColumnsWidthMode.FROM_LOD)


        doc.createSheet("block_from_lod_with_headers column_of_totals")
        helpers.block_from_lod_with_headers(doc, lod_singers, "A1", [
             ["Singer header", "Singer"],
             ["Song header", "Best song"]
        ], titulo="block_from_lod_with_headers (With total columns)", column_of_totals=True)
        doc.setColumnsWidth(lod_singers, types.ColumnsWidthMode.FROM_LOD)

        doc.createSheet("block_from_lod_with_headers row_of_totals")
        helpers.block_from_lod_with_headers(doc, lod_singers, "A1", [
             ["Singer header", "Singer"],
             ["Song header", "Best song"]
        ], titulo="block_from_lod_with_headers (With total rows)", row_of_totals=True)
        doc.setColumnsWidth(lod_singers, types.ColumnsWidthMode.FROM_LOD)


        doc.createSheet("block_from_lod_with_headers column_of_totals row_of_totals")
        helpers.block_from_lod_with_headers(doc, lod_singers, "A1", [
             ["Singer header", "Singer"],
             ["Song header", "Best song"]
        ], titulo="block_from_lod_with_headers (With total columns and rows)", column_of_totals=True, row_of_totals=True)
        doc.setColumnsWidth(lod_singers, types.ColumnsWidthMode.FROM_LOD)

        
def demo_ods_sheet_from_lol(doc):
        ## Sheet from LOL and LOD
        helpers.sheet_from_lol(doc, "sheet_from_lol", lol_numbers, lol_numbers_headers, column_of_totals=True, row_of_totals=True, titulo="LOL Sheet")
        
def demo_ods_sheet_from_lod(doc):
        helpers.sheet_from_lod(doc, "sheet_from_lod", lod_singers, column_of_totals=True, row_of_totals=True, title="LOD Sheet", styles="Integer")
        
def demo_ods_sheet_sort(doc):
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
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

def demo_ods_sheet_word_wrap(doc):
        ## Word Wrap
        doc.createSheet("Word Wrap")
        long_text = "This is a very long text that should be wrapped if the parameter is set to True, otherwise it should stay in a single line with fixed row height."
        doc.addCellWithStyle("A1", "Word Wrap False", ColorsNamed.Orange, "BoldCenter")
        doc.addCellWithStyle("A2", long_text, word_wrap=False)
        
        doc.addCellWithStyle("A4", "Word Wrap True", ColorsNamed.Orange, "BoldCenter")
        doc.addCellWithStyle("A5", long_text, word_wrap=True)


        helpers.block_from_lod(doc, "A7",  lod_singers, columns_header=1, word_wrap=False)

        helpers.block_from_lod(doc, "A12",  lod_singers, columns_header=1, word_wrap=True)


        
        doc.setColumnsWidth([2, 2, 2, 2], types.ColumnsWidthMode.MANUAL)

def demo_ods_sheet_split_with_big_lol(doc):
        helpers.sheet_split_with_big_lol(doc, "Splits in 400 rows", lol_thousands, ["Integer", "String", "Datetime"],  max_rows=400)




def demo_ods_block_from_lol(doc):
        headers = ["Product", "Qty", "Price"]
        data = [
            ["Item A", 10, 20.5],
            ["Item B", 5, 15.0],
            ["Item C", 2, 100.0],
        ]




        doc.createSheet("block_from_lol empty")
        helpers.block_from_lol(doc, "A1", [], headers=headers, title="block_from_lol (With total columns)", styles="Float2")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

        doc.createSheet("block_from_lol column of totals")
        helpers.block_from_lol(doc, "A1", data, headers=headers, column_of_totals=True, title="block_from_lol (With total columns)", styles="Float2")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

        doc.createSheet("block_from_lol row of totals")
        helpers.block_from_lol(doc, "A1", data, headers=headers, row_of_totals=True, title="block_from_lol (With total rows)", styles="Float2")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)
        
        doc.createSheet("block_from_lol both totals")
        helpers.block_from_lol(doc, "A1", data, headers=headers, column_of_totals=True, row_of_totals=True, title="block_from_lol (With both totals)", styles="Float2")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

def demo_ods_helpers_single(doc):
        ## row_title_values_total
        doc.createSheet("row_title_values_total")
        helpers.row_title_values_total(doc, "A1", "My Row Sum", [10, 20, 30])
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

        ## column_title_values_total
        doc.createSheet("column_title_values_total")
        helpers.column_title_values_total(doc, "A1", "My Col Sum", [1, 2, 3, 4, 5])
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

        ## row_totals
        doc.createSheet("row_totals")
        doc.addRowWithStyle("A1", [100, 200, 300, 400])
        helpers.row_totals(doc, "A2", ["#SUM"] * 4, row_from="1")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

        ## column_totals
        doc.createSheet("column_totals")
        doc.addColumnWithStyle("A1", [10, 20, 30, 40])
        helpers.column_totals(doc, "B1", ["#SUM"] * 4, column_from="A")
        doc.setColumnsWidth(doc, types.ColumnsWidthMode.FROM_SHEET_CELLS)

def demo_ods_columns_width_modes(doc):

    # 1. MANUAL
    helpers.sheet_from_lod(doc, "Width MANUAL", lod_widths, columns_width_mode=types.ColumnsWidthMode.MANUAL, value=[5, 10, 5])
    
    # 2. FROM_LOD (Default behavior: max of all)
    helpers.sheet_from_lod(doc, "Width FROM_LOD", lod_widths, columns_width_mode=types.ColumnsWidthMode.FROM_LOD)

    # 3. FROM_LOD_0 (Uses first dictionary)
    helpers.sheet_from_lod(doc, "Width FROM_LOD_0", lod_widths, columns_width_mode=types.ColumnsWidthMode.FROM_LOD_0)

    # 4. FROM_LOD_QUANTILE_90 (Uses 90th percentile of lengths)
    helpers.sheet_from_lod(doc, "Width FROM_LOD_Q90", lod_widths, columns_width_mode=types.ColumnsWidthMode.FROM_LOD_QUANTILE_90)

    # 5. FROM_SHEET_CELLS (Measures from generated cells)
    helpers.sheet_from_lod(doc, "Width FROM_SHEET_CELLS", lod_widths, columns_width_mode=types.ColumnsWidthMode.FROM_SHEET_CELLS)


    doc.createSheet("Width FROM_LOL")
    doc.addListOfRowsWithStyle("A1", lol_numbers)
    doc.setColumnsWidth(lol_numbers, types.ColumnsWidthMode.FROM_LOL)

    doc.createSheet("Width FROM_LIST")
    doc.addListOfRowsWithStyle("A1", [lol_numbers[0]])
    doc.setColumnsWidth(lol_numbers[0], types.ColumnsWidthMode.FROM_LIST)
