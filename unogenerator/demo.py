## @namespace unogenerator.demo
## @brief Generate ODF example files

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
    doc.print_styles()
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
    #doc.pageBreak()
    
    
    doc.save()
    doc.export_docx()
    doc.export_pdf()
    doc.close()
    return "demo_odt took {}".format(datetime.now()-doc.init)
    

if __name__ == "__main__":
    main()
