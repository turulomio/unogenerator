from argparse import ArgumentParser,  RawTextHelpFormatter
from gettext import translation
from os import path,  makedirs
from polib import POEntry,  POFile, pofile
from subprocess import run
from pkg_resources import resource_filename
from unogenerator.commons import __version__, argparse_epilog
from unogenerator import ODT
import xml.etree.ElementTree as ET


try:
    t=translation('unogenerator', resource_filename("unogenerator","locale"))
    _=t.gettext
except:
    _=str

def run_check(command, shell=False):
    p=run(command, shell=shell, capture_output=True);
    if p.returncode!=0:
        print(f"Error en comando. {command}")
        print("STDOUT:")
        print(p.stdout.decode('utf-8'))
        print("STDERR:")
        print(p.stderr.decode('utf-8'))
        print("Saliendo de la instalación")
        exit(2)


def innercontent(element):
    return (element.text or '') + ''.join(ET.tostring(e, 'unicode').replace("ns0:", "") for e in element)
            
def main():
    parser=ArgumentParser(prog='unogenerator', description=_('Generate a po and pot file from ODF'), epilog=argparse_epilog(), formatter_class=RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--from_language', action='store', help=_('Language to translate from. Example codes: es, fr, en, md'), required=True, metavar="CODE")
    parser.add_argument('--to_language', action='store', help=_('Language to translate to. Example codes: es, fr, en, md'), required=True,  metavar="CODE")
    parser.add_argument('--input', action='append', help=_('File to translate'), required=True,  metavar="FILE")
    parser.add_argument('--output_directory', action='store', help=_('Output directory'), required=True,  metavar="FILE")
    parser.add_argument('--undetected', action='append', help=_('Undetected strings to apped to translation'), default=[])
    parser.add_argument('--translate', action='store_true', help=_('Make translation too'), default=False)
    parser.add_argument('--fake', action='store_true', help=_('Sets fake to all strings'), default=False)
    args=parser.parse_args()
    
    command(args.from_language, args.to_language, args.input, args.output_directory, args.translate, args.undetected, args.fake)
def same_entries_to_ocurrences(l):
    l= sorted(l, key=lambda x: (x[0], x[1], x[2], x[3]))
    r=[]
    for filename, type, number,  position,  text in l:
        r.append((filename, f"{type}#{number}#{position}"))
    return r
    
## Returns a list of entrieses withh filename and type sorted by number and positionion
def filter_occurrences(entries, filename, type):
    entries= sorted(entries, key=lambda x: (x[0], x[1], x[2], x[3]))
    r=[]
    for e_filename, e_type, e_number,  e_position,  text in entries:
        if filename==e_filename and e_type==type:
            r.append((e_filename, e_type, e_number,  e_position,  text))
    return r
    
    
def getEntriesFromDocument(filename):
        r=[]
        doc=ODT(filename)
        print(dir(doc.document))

        #Extract strings from paragraphs
        r=r+entries_from_paragraph_enumeration("Paragraph", doc.cursor.Text.createEnumeration(), filename)

        for style in doc.getPageStyles():
#            print(style.getName())
            
            #Extract strings from headers
            ht=style.HeaderText
            if ht is None:
                continue
            r=r+entries_from_paragraph_enumeration("HeaderParagraph", ht.createEnumeration(), filename)
            #Extract strings from foot
            ft=style.FooterText
            if ft is None:
                continue
            r=r+entries_from_paragraph_enumeration("FooterParagraph", ft.createEnumeration(), filename)

        for table in doc.document.getTextTables():
            print

        doc.close()
        return r
        
def entries_from_paragraph_enumeration(title, enumeration, filename):
        r=[]
        for i,  par in enumerate(enumeration):
#            print(i, par, dir(par))
            if  par.supportsService("com.sun.star.text.Paragraph") :
                for position, element in enumerate(par.createEnumeration()):
                    text_=element.getString()
                    if text_ !="" and text_!=" " and text_!="  ":
                        entry=(filename, title,  i,  position, text_)
                        r.append(entry)

        return r
            
    #Extract strings from headers
#    ' Turn headers on and then make them shared!
#oPstyle.HeaderOn = True
#oPstyle.HeaderShared = True
#' The is also a RightText and a LeftText
#oHeader = oPstyle.RightPageHeaderContent
#oText = oHeader.CenterText
#' You may now set the text object to be anything you desire
#' Use setSTring() from the text object to set simple text.
#' Use a cursor to insert a field (such as the current sheet name).
#' First, clear any existing text!
#oText.setString("")
#oCursor = oText.createTextCursor()
#oText.insertString(oCursor, "Sheet: ", False)
#' This will have the sheet name of the current sheet!
#sService = "com.sun.star.text.TextField.SheetName"
#oField = oDoc.createInstance(sService)
#oText.insertTextContent(oCursor, oField, False)
## Input is a listst of strings

def generate_pot_file(potfilename, set_strings, entries):
    file_pot = POFile()
    file_pot.metadata = {
        'Project-Id-Version': '1.0',
        'Report-Msgid-Bugs-To': 'you@example.com',
        'POT-Creation-Date': '2007-10-18 14:00+0100',
        'PO-Revision-Date': '2007-10-18 14:00+0100',
        'Last-Translator': 'you <you@example.com>',
        'Language-Team': 'English <yourteam@example.com>',
        'MIME-Version': '1.0',
        'Content-Type': 'text/plain; charset=utf-8',
        'Content-Transfer-Encoding': '8bit',
    }
    for s in set_strings:
        same_entries=[] #Join seame text entries
        for filename, type, number, position, string_ in entries:
            if string_==s:
                same_entries.append((filename, type, number, position, string_))

        entry = POEntry(
            msgid=s,
            msgstr='', 
            occurrences=same_entries_to_ocurrences(same_entries)
        )
        file_pot.append(entry)
    file_pot.save(potfilename)

def command(from_language, to_language, input, output_directory, translate,  undetected_strings=[], fake=False):   
    makedirs(output_directory, exist_ok=True)
    makedirs(f"{output_directory}/{to_language}", exist_ok=True)
    
    pot=f"{output_directory}/catalogue.pot"
    po=f"{output_directory}/{to_language}/{to_language}.po"
        
    entries=[]#List of (filename,"type", numero, posicion) type=Paragraph, numero=numero parrafo y posición orden dentro del parrafo
    print(_("Extracting strings from:"))
    for filename in input:
        print(_(f"   - {filename}"))
        entries=entries+getEntriesFromDocument(filename)
        
    # Set with distinct strings of all entries
    set_strings=set()
    for filename, type, number, position, string_ in entries:
        set_strings.add(string_)

    #Generate pot file
    generate_pot_file(pot, set_strings, entries)
    
    #Merging pot with out po file
    if path.exists(po)==False:
        run_check(["msginit", "-i", pot,  "-o", po])
    run_check(["msgmerge","-N", "--no-wrap","-U", po, pot])
    
    print(f"{len(set_strings)} different strings detected")
    
        
    # Creates a dictionary of translations
    dict_po={}
    for i, entry in enumerate(pofile(po)):
        if fake is True:
            dict_po[entry.msgid]=f"{{{entry.msgid}}}"
        else:
            if entry.msgstr == "":
                dict_po[entry.msgid]=entry.msgid
            else:
                dict_po[entry.msgid]=entry.msgstr
    
    if translate is True:
        print(_("Translating files to:"))
        for filename in input:
            output=f"{output_directory}/{to_language}/{path.basename(filename)}"
            print(_(f"   - {output}"))
            write_translation(output,  dict_po, entries)
            
            
def write_translation(filename, dict_po, entries):
            doc=ODT(filename)
            search_descriptor=None
            for filename, type, number,  position,  text in filter_occurrences(entries, filename, "Paragraph"):
                search_descriptor=doc.find_and_replace_and_return_descriptor(text, dict_po[text], search_descriptor)
            
            doc.save(filename)
            doc.export_pdf(filename+".pdf")
            doc.close()
        

