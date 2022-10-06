from argparse import ArgumentParser,  RawTextHelpFormatter
from gettext import translation
from os import path,  makedirs
from polib import POEntry,  POFile, pofile
from subprocess import run
from importlib.resources import files
from unogenerator.commons import __version__, argparse_epilog, red
from unogenerator import ODT


try:
    t=translation('unogenerator', files("unogenerator") / 'locale')
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

            
def main():
    parser=ArgumentParser(prog='unogenerator', description=_('Generate a po and pot file from ODF'), epilog=argparse_epilog(), formatter_class=RawTextHelpFormatter)
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--from_language', action='store', help=_('Language to translate from. Example codes: es, fr, en, md'), required=True, metavar="CODE")
    parser.add_argument('--to_language', action='store', help=_('Language to translate to. Example codes: es, fr, en, md'), required=True,  metavar="CODE")
    parser.add_argument('--input', action='append', help=_('Files to translate. You can set several files.'), required=True,  metavar="FILE")
    parser.add_argument('--output_directory', action='store', help=_('Output directory with results and catalogues. Default: ./unogenerator_translation/') , metavar="FILE",  default="./unogenerator_translation/")
    parser.add_argument('--undetected', action='append', help=_('Undetected strings to append to translation. Default: []'), default=[])
    parser.add_argument('--fake', action='store_true', help=_('Sets a fake translation to all strings. Default: False'), default=False)
    parser.add_argument('--pdf', action='store_true', help=_('Creates translated file in pdf. Default: False'), default=False)
    args=parser.parse_args()
    
    command(args.from_language, args.to_language, args.input, args.output_directory, args.undetected, args.fake, args.pdf)

def same_entries_to_ocurrences(l):
    l= sorted(l, key=lambda x: (x[0], x[1], x[2], x[3]))
    r=[]
    for filename, type, number,  position,  text in l:
        r.append((filename, f"{type}#{number}#{position}"))
    return r

    
def getEntriesFromDocument(filename):
        r=[]
        doc=ODT(filename)
        #Extract strings from metadata
        metadata=doc.getMetadata()
        for key,  value in metadata.items():
            if value.__class__.__name__=="str" and len(value)>0:
                r.append((filename, f"Metadata_{key}",  0,  0, value))

        #Extract strings from paragraphs
        r=r+entries_from_paragraph_enumeration("Paragraph", doc.cursor.Text.createEnumeration(), filename)

        for style in doc.getPageStylesObjects():            
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
#                        print(entry)
                        r.append(entry)
            elif  par.supportsService("com.sun.star.text.TextTable") :
#                print("AQUQI", par, dir(par), par.getColumns().Count, par.getRows().Count)
                for column in range(par.getColumns().Count):
                    for row in range(par.getRows().Count):
                        try:
                            cell=par.getCellByPosition(column, row)
                            r=r+entries_from_paragraph_enumeration(f"{title}Table,paragraph{i},column{column},row{row}", cell.createEnumeration(), filename)
                        except:
                            print (_("Perhaps a merged string"))
                            continue
        return r
            
        

            

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

def command(from_language, to_language, input, output_directory="./unogenerator_translation/", undetected_strings=[], fake=False, pdf=False):   
    makedirs(output_directory, exist_ok=True)
    makedirs(f"{output_directory}/{to_language}", exist_ok=True)
    
    pot=f"{output_directory}/catalogue.pot"
    po=f"{output_directory}/{to_language}/{to_language}.po"
        
    entries=[]#List of (filename,"type", numero, posicion) type=Paragraph, numero=numero parrafo y posición orden dentro del parrafo
    print(_("Extracting strings from:"))
    
    goodinput=[]#To remove bad files and dont make errors deleting in list
    for filename in input:
        if filename.upper()[-3:]!="ODT":
            print(f"   - {filename}. " + red(_("Only ODT files are supported")))
            continue
        print(f"   - {filename}")
        goodinput.append(filename)
        entries=entries+getEntriesFromDocument(filename)
    input=goodinput
        
        
    # Set with distinct strings of all entries
    set_strings=set()
    for filename, type, number, position, string_ in entries:
        set_strings.add(string_)

    #Generate pot file
    print(_("Generating {0}").format(pot))
    generate_pot_file(pot, set_strings, entries)
    
    #Merging pot file to po (new or not)
    if path.exists(po) is False:
        run_check(["msginit", "-i", pot,  "-o", po])
    run_check(["msgmerge","-N", "--no-wrap","-U", "-F", po, pot])
    print(_("   - {0} different strings detected").format(len(set_strings)))
    
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
    
    print(_("Translating files to:"))
    for filename_input in input:
        output=f"{output_directory}/{to_language}/{path.basename(filename_input)}"
        print(f"   + {output}")
        write_translation(filename_input, output,  dict_po, entries, pdf)
            
            
def write_translation(original, filename, dict_po, entries, pdf):
    doc=ODT(original)
    search_descriptor=None
    replaced=0
    #Replaces all strings generic
    for filename_, type, number,  position,  text in entries:
        if filename_==path.basename(filename):
            #Translate metadata
            if type=="Metadata_Author":
                doc.setMetadata(author=dict_po[text])
                replaced=replaced+1
            elif type=="Metadata_Description":
                doc.setMetadata(description=dict_po[text])
                replaced=replaced+1
            elif type=="Metadata_Title":
                doc.setMetadata(title=dict_po[text])
                replaced=replaced+1
            elif type=="Metadata_Subject":
                doc.setMetadata(subject=dict_po[text])
                replaced=replaced+1
            else:#Translate rest of entries
                search_descriptor=doc.find_and_replace(text, dict_po[text], search_descriptor)
                if search_descriptor is not None:
                    replaced=replaced+1
                
    
    print(f"      - Translated {replaced} strings")
    doc.save(filename)
    if pdf is True:
        print("      - PDF translation generated")
        doc.export_pdf(filename[:-4] +".pdf")
    doc.close()
        

