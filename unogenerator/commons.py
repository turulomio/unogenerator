## @namespace unogenerator.commons
from colorama import Fore, Style
from datetime import datetime, date
from gettext import translation
from logging import info, ERROR, WARNING, INFO, DEBUG, CRITICAL, basicConfig, error
from pkg_resources import resource_filename
from psutil import process_iter
from uno import createUnoStruct
from unogenerator.reusing.listdict_functions import listdict_min
from time import sleep

__version__ = '0.18.0'
__versiondatetime__=datetime(2022, 2, 10, 8, 6)
__versiondate__=__versiondatetime__.date()

try:
    t=translation('unogenerator',resource_filename("unogenerator","locale"))
    _=t.gettext
except:
    _=str

class ColorsNamed:
    Black=0x111111
    White=0xFFFFFF
    Blue=0x9999ff
    Red=0xFF9999
    Green=0xc0FFc0
    Orange=0xffdca8
    Yellow=0xffffc0
    GrayLight=0xd2ced4
    GrayDark=0xa29ea4
    GrayVeryDark=0x726e74

def datetime2uno( dt):
    r=createUnoStruct("com.sun.star.util.DateTime")
    r.Year=dt.year
    r.Month=dt.month
    r.Day=dt.day
    r.Hours=dt.hour
    r.Minutes=dt.minute
    r.Seconds=dt.second
    return r

def date2uno( dt):
    r=createUnoStruct("com.sun.star.util.Date")
    r.Year=dt.year
    r.Month=dt.month
    r.Day=dt.day
    return r
    

## Function used in argparse_epilog
## @return String
def argparse_epilog():
    return _("Developed by Mariano Muñoz 2020-{}").format(__versiondate__.year)


## Allows to operate with columns letter names
## @param letter String with the column name. For example A or AA...
## @param number Columns to move
## @return String With the name of the column after movement
def columnAdd(letter, number):
    letter_value=column2number(letter)+number
    return number2column(letter_value)


def rowAdd(letter,number):
    return str(int(letter)+number)

## Convierte un número  con el numero de columna al nombre de la columna de hoja de datos
##
## Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA.
def number2column(n):
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

## Convierte una columna de hoja de datos a un número
##
## Excel-style column name to number, e.g., A = 1, Z = 26, AA = 27, AAA = 703.
def column2number(name):
    n = 0
    for c in name:
        n = n * 26 + 1 + ord(c) - ord('A')
    return n

## Converts a column name to a index position (number of column -1)
def column2index(name):
    return column2number(name)-1

## Convierte el nombre de la fila de la hoja de datos a un índice, es decir el número de la fila -1
def row2index(number):
    return int(number)-1

## Covierte el nombre de la fila de la hoja de datos a un  numero entero que corresponde con el numero de la fila
def row2number(strnumber):
    return int(strnumber)

## Convierte el numero de la fila al nombre de la fila en la hoja de datos , que corresponde con un string del numero de la fila
def number2row(number):
    return str(number)
    
## Convierte el indice de la fila al numero cadena de la hoja de datos
def index2row(index):
    return str(index+1)
    
## Convierte el indice de la columna a la cadena de letras de la columna de la hoja de datos
def index2column(index):
    return number2column(index+1)

## Class that manage spreadsheet coordinates (letter + number)
class Coord:
    def __init__(self, strcoord):
        self.letter, self.number=self.__extract(strcoord)

    def __repr__(self):
        return f"Coord <{self}>"
        
    def __str__(self):
        return self.string()

    def __eq__(self, b):
        b=Coord.assertCoord(b)
        if self.letter==b.letter and self.number==b.number:
            return True
        return False
        
    def __extract(self, strcoord):
        if strcoord.find(":")!=-1:
            print("I can't manage range coord")
            return
        letter=""
        number=""
        for l in strcoord:
            if l.isdigit()==False:
                letter=letter+l
            else:
                number=number+l
        return (letter,number)

    ## Returns Coord string
    ## @return string For example "Z1"
    def string(self):
        return self.letter+self.number

    ## Add a number of rows to the Coord and return itself
    ## @param num Integer Can be positive and negative. When num is negative, if Coord.number is less than 1, returns 1
    def addRow(self, num=1):
        if self.numberIndex()+num<0:
            self.number="1"
        else:
            self.number=rowAdd(self.number, num)
        return self

    ## Add a number of columns/letters to the Coord and return itself
    ## @param num Integer. Can be positive and negative. When num is negative, if Coord.letter is less than A, returns A.
    def addColumn(self, num=1):
        if self.letterIndex()+num<0:
            self.letter="A"
        else:
            self.letter=columnAdd(self.letter, num)
        return self

    ## Add a number of rows to the Coord and return a copy of the object
    ## @param num Integer Can be positive and negative. When num is negative, if Coord.number is less than 1, returns 1
    def addRowCopy(self, num=1):
        r=Coord(self.string())
        r.addRow(num)
        return r

    ## Add a number of columns/letters to the Coord and return a copy of the objject
    ## @param num Integer. Can be positive and negative. When num is negative, if Coord.letter is less than A, returns A.
    def addColumnCopy(self, num=1):
        r=Coord(self.string())
        r.addColumn(num)
        return r

    ## Reterns the letter/column index (letterPosition()-1)
    ## @return int
    def letterIndex(self):
        return column2index(self.letter)

    ## Returns the letter/column position
    ## @return int 
    def letterPosition(self):
        return column2number(self.letter)

    ## Returns the number/row index (numberPosition()-1)
    def numberIndex(self):
        return row2index(self.number)

    ## Returns the number/row position
    def numberPosition(self):
        return row2number(self.number)

    @staticmethod
    def assertCoord(o):
        if o.__class__==Coord:
            return o
        elif o.__class__==str:
            return Coord(o)
        else:
            error("{} is not a coord".format(o))
            



## Class that manages spreadsheet Ranges for ods and xlsx
class Range:
    def __init__(self,strrange):
        self.start, self.end=self.__extract(strrange)

    def __repr__(self):
        return f"Range <{self}>"
        
    def __str__(self):
        return self.string()

    ##Return the outcome of the test b in a. Note the reversed operands.
    def __contains__(self, b):
        if (    b.letterIndex()>=self.start.letterIndex() and 
                b.letterIndex()<=self.end.letterIndex() and 
                b.numberIndex()>=self.start.numberIndex() and 
                b.numberIndex()<=self.end.numberIndex())==True:
            return True
        return False

    ## Returns a list of rows of all Coord objects in the range
    def coords(self):
        r=[]
        for row in range(self.numRows()):
            tmprow=[]
            for column in range(self.numColumns()):
                tmprow.append(self.start.addRowCopy(row).addColumnCopy(column))
            r.append(tmprow)
        return r

    ## Converts a string to a Range. Returns None if conversion can't be done
    def __extract(self,range):
        if range.find(":")==-1:
            print("I can't manage this range")
            return
        a=range.split(":")
        return (Coord(a[0]), Coord(a[1]))

    ## String of a range in spreadsheets
    def string(self):
        return "{}:{}".format(self.start.string(), self.end.string())

    ## Number of range of the range
    def numRows(self):
        return row2number(self.end.number)-row2number(self.start.number) +1

    ## Number of columns of the range
    def numColumns(self):
        return column2number(self.end.letter)-column2number(self.start.letter) +1

    ## Returns a Range object even o is a str or a Range
    @staticmethod
    def assertRange(o):
        if o.__class__==Range:
            return o
        elif o.__class__==str:
            return Range(o)


    ## Adds a row to the end Coord, so it adds a row to the range
    def addRowAfter(self, num=1):
        self.end=self.end.addRow(num)
        return self

    def addRowAfterCopy(self, num=1):
        r=Range(self.string())
        r.addRowAfter(num)
        return r

    ## Adds a column to the end Coord, so it adds a column to the range
    def addColumnAfter(self, num=1):
        self.end=self.end.addColumn(num)
        return self

    def addColumnAfterCopy(self, num=1):
        r=Range(self.string())
        r.addColumnAfter(num)
        return r

    ## Adds a row to the top of the start Coord, so it adds a row to the range. If start Coord number is 1 returns the same Coord
    def addRowBefore(self, num=1):
        self.start=self.start.addRow(-num)
        return self

    def addRowBeforeCopy(self, num=1):
        r=Range(self.string())
        r.addRowBefore(num)
        return r

    ## Adds a column to the end Coord, so it adds a column to the range. If start Coord letter is A, returns the same Coord
    def addColumnBefore(self, num=1):
        self.start=self.start.addColumn(-num)
        return self

    def addColumnBeforeCopy(self, num=1):
        r=Range(self.string())
        r.addColumnBefore(num)
        return r

    ## Returns a list of tuples(column_index, row_index)
    def indexes_list(self, plain=False):
        r=[]
        if plain is True:
            for letter_index in range(self.start.letterIndex(), self.end.letterIndex()+1):
                for number_index in range(self.start.numberIndex(), self.end.numberIndex()+1):
                    r.append((letter_index, number_index))
        else:
            for number_index in range(self.start.numberIndex(), self.end.numberIndex()+1):
                row=[]
                for letter_index in range(self.start.letterIndex(), self.end.letterIndex()+1 ):
                    row.append((letter_index, number_index))
                r.append(row)
        return r

    def coords_list(self, plain=False):
        r=[]
        if plain is True:
            for letter_index, number_index in self.indexes_list():
                r.append(Coord_from_index(letter_index, number_index))
        else:
            for row  in self.indexes_list():
                r2=[]
                for  letter_index , number_index  in row:
                    r2.append(Coord_from_index(letter_index, number_index))
                r.append(r2)
        return r

## Sets debug sustem, needs
## @param args It's the result of a argparse     args=parser.parse_args()        
def addDebugSystem(level):
    logFormat = "%(asctime)s.%(msecs)03d %(levelname)s %(message)s [%(module)s:%(lineno)d]"
    dateFormat='%F %I:%M:%S'

    if level=="DEBUG":#Show detailed information that can help with program diagnosis and troubleshooting. CODE MARKS
        basicConfig(level=DEBUG, format=logFormat, datefmt=dateFormat)
    elif level=="INFO":#Everything is running as expected without any problem. TIME BENCHMARCKS
        basicConfig(level=INFO, format=logFormat, datefmt=dateFormat)
    elif level=="WARNING":#The program continues running, but something unexpected happened, which may lead to some problem down the road. THINGS TO DO
        basicConfig(level=WARNING, format=logFormat, datefmt=dateFormat)
    elif level=="ERROR":#The program fails to perform a certain function due to a bug.  SOMETHING BAD LOGIC
        basicConfig(level=ERROR, format=logFormat, datefmt=dateFormat)
    elif level=="CRITICAL":#The program encounters a serious error and may stop running. ERRORS
        basicConfig(level=CRITICAL, format=logFormat, datefmt=dateFormat)
    info("Debug level set to {}".format(level))



## Creates a Coord object from spreadsheet letters
def Coord_from_letters(column, letter):
    return Coord(column+letter)

## Creates a Coord object from spreadsheet index coords
def Coord_from_index(column_index, row_index):
    return Coord(index2column(column_index)+index2row(row_index))
    
## Creates a Range object from itself stard and end coords
def Range_from_coords(start, end):
    start=Coord.assertCoord(start)
    end=Coord.assertCoord(end)
    return Range(f"{start}:{end}")

def generate_formula_total_string(key, coord_from, coord_to):
    if key == "#SUM":
        s="=SUM({}:{})".format(coord_from.string(), coord_to.string())
    elif key == "#AVG":
        s="=AVERAGE({}:{})".format(coord_from.string(), coord_to.string())
    elif key == "#MEDIAN":
        s="=MEDIAN({}:{})".format(coord_from.string(), coord_to.string())
    else:
        s=key
    return s

def guess_object_style(o):
    if o is None:
        return "Normal"
    elif o.__class__.__name__=="int":
        return "Integer"
    elif o.__class__.__name__=="str":
        return "Normal"
    elif o.__class__.__name__ in ["Currency", "Money" ]:
        return o.currency
    elif o.__class__.__name__=="Percentage":
        return "Percentage"
    elif o.__class__.__name__ in ("Decimal", "float"):
        return  "Float2"
    elif o.__class__.__name__=="datetime":
        return "Datetime"
    elif o.__class__.__name__=="date":
        return "Date"
    elif o.__class__.__name__=="timedelta":
        return "Normal"
    elif o.__class__.__name__=="time":
        return "Time"
    elif o.__class__.__name__=="bool":
        return "Bool"
    else:
        info("guess_object_style not guessed {}".format( o.__class__))
        return "Bold"
        
def datetime2localc1989(o):
    delta = o -  datetime(1899, 12, 30)
    return float(delta.days) + float(delta.seconds) / 86400
    
def date2localc1989(o):
    delta = o -  date(1899, 12, 30)
    return float(delta.days) 

def time2localc1989(o):
    seconds=o.hour*3600+o.minute*60+o.second
    return float(seconds) / 86400
    
## Used to change port when there are multiple sockets of libreoffice accepting
def next_port(last,  first_port,  instances):
    if last==first_port+instances -1:
        return first_port
    else:
        return last+1

def get_range_from_iterable_object( coord_start, o):
    coord_start=Coord.assertCoord(coord_start)
    r=Range(f"{coord_start.string()}:{coord_start.string()}")

    if len(o)==0:
        return r

    if o[0].__class__.__name__ in ("list","dict", "OrderedDict"): #Iterables con len
        len_rows=len(o)-1
        len_columns=len(o[0])-1
    else:
        len_rows=0
        len_columns=len(o)-1
    return Range(f"{coord_start.string()}:{coord_start.addRowCopy(len_rows).addColumnCopy(len_columns).string()}")

## @param attempts (Integer). Sometimes when server is busy this method fails to detect info. So I make several attempts with a time interval
## @return listdict with process info

def get_from_process_info(cpu_percentage=False, attempts=10):
    for attempt in range(attempts):
        try:
            r=[]
            for p in process_iter(['name','cmdline', 'pid']): 
                d={}
                if p.info['name']=='soffice.bin':
                    if  'file:///tmp/unogenerator'  in ' '.join(p.info['cmdline']):
                        d["port"]=int(p.info['cmdline'][1][-4:])
                        d["pid"]=p.pid
                        d["mem"]=p.memory_info().rss
                        d["cpu_number"]=p.cpu_num()
                        if cpu_percentage is True:
                            d["cpu_percentage"]=p.cpu_percent(interval=0.01)
                            d["object"]=p
                        r.append(d)
            if len(r)==0:
                raise
            return r
        except:
            print(_(f"I couldn't detect unogenerator process info ({attempt}/{attempts} attempts)"))
            sleep(5)
            continue
    print(_("Have you launched unogenerator instances?. Please run unogenerator_start"))
    return []
    

def get_from_process_numinstances_and_firstport():        
    ld=get_from_process_info()
    if len(ld)==0:
        print(_("I couldn't detect unogenerator instances"))
    return len(ld), listdict_min(ld,"port")
   
def is_formula(s):
    if s is None:
        return False
    s=str(s)
    if s.startswith("=") or s.startswith("+"):
        return True
    return False
   
def red(s):
    return Style.BRIGHT + Fore.RED + str(s) + Style.RESET_ALL
def green(s):
    return Style.BRIGHT + Fore.GREEN + str(s) + Style.RESET_ALL
def magenta(s):
    return Style.BRIGHT + Fore.MAGENTA + str(s) + Style.RESET_ALL
