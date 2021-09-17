## THIS IS FILE IS FROM https://github.com/turulomio/reusingcode IF YOU NEED TO UPDATE IT PLEASE MAKE A PULL REQUEST IN THAT PROJECT
## DO NOT UPDATE IT IN YOUR CODE IT WILL BE REPLACED USING FUNCTION IN README

from decimal import Decimal
from json import dumps
from logging import warning
from currency import Currency
from percentage import Percentage

def valueORempty(v):
    return "" if v is None else v

def list2string(lista):
        """Covierte lista a string"""
        if  len(lista)==0:
            return ""
        if str(lista[0].__class__) in ["<class 'int'>", "<class 'float'>"]:
            resultado=""
            for l in lista:
                resultado=resultado+ str(l) + ", "
            return resultado[:-2]
        elif str(lista[0].__class__) in ["<class 'str'>",]:
            resultado=""
            for l in lista:
                resultado=resultado+ "'" + str(l) + "', "
            return resultado[:-2]

## Reverse function of list2string where class is a str
def string2list_of_strings(s):
    arr=[]
    if s!="":
        arrs=s.split(", ")
        for a in arrs:
            arr.append(a[1:-1])
    return arr

def string2list_of_integers(s, separator=", "):
    """Convers a string of integer separated by comma, into a list of integer"""
    arr=[]
    if s!="":
        arrs=s.split(separator)
        for a in arrs:
            arr.append(int(a))
    return arr
    
    
## Converts a string  to a decimal
def string2decimal(s, type=1):
    if type==1: #2.123,25
        try:
            return Decimal(s.replace(".","").replace(",", "."))
        except:
            return None

## Converts a decimal to a localized number string
def l10nDecimal(dec, digits=2):
    from PyQt5.QtCore import QLocale
    l=QLocale()
    return l.toCurrencyString(float(dec))

## Converts strings True or False to boolean
## @param s String
## @return Boolean
def str2bool(value):
    if value=="0" or value.lower()=="false":
        return False
    elif value=="1" or value.lower()=="true":
        return True

## Converts boolean to  True or False string
## @param s String
## @return Boolean
def bool2string(b):
    if b==True:
        return "VERDADERO"
    return "FALSO"
    
## Function that converts a None value into a Decimal('0')
## @param dec Should be a Decimal value or None
## @return Decimal
def none2decimal0(dec):
    return none2alt(dec,Decimal('0'))

## If a value is None, returns an alternative
def none2alt(value, alternative):
    if value==None:
        return alternative
    return value

## Bytes 2 string
def b2s(b, code='UTF-8'):
    return b.decode(code)
    
def s2b(s, code='UTF8'):
    """String 2 bytes"""
    if s==None:
        return "".encode(code)
    else:
        return s.encode(code)

def c2b(state):
    """QCheckstate to python bool"""
    from PyQt5.QtCore import Qt
    if state==Qt.Checked:
        return True
    else:
        return False

def b2c(booleano):
    """Bool to QCheckstate"""
    from PyQt5.QtCore import Qt
    if booleano==True:
        return Qt.Checked
    else:
        return Qt.Unchecked     

## Returns a list with object in positions removed
def list_remove_positions(l, listindex):
    if l is None:
        warning("I can't remove positions from a None list")
        return None
    r=[]
    for i, o in enumerate(l):
        if i not in listindex:
            r.append(o)
    return r

## LOR is a list of list. Naned List Of Rows, used in myqtablewidget for example
## @param rows LOR
## @param index int with the index of the position where we are going to insert row
## @param column List with the values to add. Must be of the same size of rows
def lor_add_column(rows, index, column):
    if len(rows)!=len(column):
        warning("I can't add a column with different size of LOR")
        return
    r_rows=[]
    for i, row in enumerate(rows):
        r_rows.append(row[0:index] + [column[i],] + row[index:len(row)])
    return r_rows

## LOR is a list of list. Naned List Of Rows, used in myqtablewidget
## @param listindex is a list of column indexes to remove
def lor_remove_columns(rows, listindex):
    r_rows=[]
    for i, row in enumerate(rows):
        r_rows.append(list_remove_positions(row,listindex))
    return r_rows

## LOR is a list of list. Naned List Of Rows, used in myqtablewidget
## @param listindex is a list of column indexes to remove
def lor_remove_rows(rows, listindex):
    return list_remove_positions(rows, listindex) #It's a list but of row

## Return a lor transposed. Changed rows by columns
def lor_transposed(lor):
    r=[]
    columns=len(lor[0])
    for column in range(columns):
        tran_row=[]
        for row in lor:
            tran_row.append(row[column])
        r.append(tran_row)
    return r

## Extract a column from the list of row
def lor_get_row(lor, row):
    return lor[row]

## Extract a column from the list of row
def lor_get_column(lor, column):
    r=[]
    for row in lor:
        r.append(row[column])
    return r

## Return sum of values of a column from and index to and other index position. This method ignores None values
## This method can sum several objects
## @param row
## @param from_index
## @param to_index
## @param zerovalue 0 or Money(self.mem, 0, self.mem.localcurrency)....
def lor_sum_row(row, from_index, to_index, zerovalue=0):
    s=zerovalue
    for i, column in enumerate(row):
        if i>=from_index and i<=to_index:
            if column is not None:
                s=s + column
    return s

## Return sum of values of a column from and index to and other index position. This method ignores None values
## This method can sum several objects
## @zerovalue 0 or Money(self.mem, 0, self.mem.localcurrency)....
def lor_sum_column(lor, column, from_index, to_index, zerovalue=0):
    s=zerovalue
    for i, row in enumerate(lor):
        if i>=from_index and i<=to_index:
            if row[column] is not None:
                s=s + row[column]
    return s

## String to linux shell
#def string2shell(cadena):
#    cadena=str(cadena)
#    cadena=cadena.replace("'","\\'")
#    return cadena

## strint to latex
def string2tex(cadena):
    cadena=str(cadena)
    cadena=cadena.replace('[','$ [ $')
    cadena=cadena.replace(']','$ ] $')
    cadena=cadena.replace('&','\&')
    cadena=cadena.replace('²','$ ^2 $')
    cadena=cadena.replace('#', '\#')
    return cadena

## Converts a string to set inside an XML to a valid XML string
def string2xml(s):
    s=s.replace('"','&apos;' )
    s=s.replace('<','&lt;' )
    s=s.replace('>','&gt;' )
    s=s.replace('&','&amp;' )
    s=s.replace("'",'&apos;' )
    return s

## Converts a string to set inside an XML to a valid XML string
def xml2string(s):
    s=s.replace('&apos;','"')
    s=s.replace('&lt;','<')
    s=s.replace('&gt;','>')
    s=s.replace('&amp;','&')
    s=s.replace('&apos;',"'")
    return s
    
## Converts my common objects to its numeric value
def object2value(o):
    if o.__class__.__name__ in ["int", "float", "Decimal"]:
        return o
    elif o.__class__.__name__ in ["Currency",  "Money", "USD", "EUR"]:
        return o.amount
    elif o.__class__.__name__ == "Percentage":
        return o.value
    return o
        
        
def value2object(value, stringtypes):
    if value is None:
        return None
    if stringtypes=="int":
        return int(value)
    elif stringtypes=="float":
        return float(value)
    elif stringtypes=="Decimal":
        return Decimal(value)
    elif stringtypes==["datetime", "date", "time"]:
        return value
    elif stringtypes in ["EUR", "€"]:
        return Currency(value, "EUR")
    elif stringtypes in ["USD", "$"]:
        return Currency(value, "EUR")
    elif stringtypes=="Percentage":
        return Percentage(value, 1)
    return value

    
def var2json(var):
    if var.__class__.__name__=="Decimal":
        return float(var)
    elif var.__class__.__name__=="datetime":
        return var.isoformat()[:-6]+"Z"
    elif var.__class__.__name__=="date":
        return str(var)
    elif var.__class__.__name__=="bool":
        return dumps(var)
    elif var.__class__.__name__=="Percentage":
        return dumps(var2json(var.value))
    elif var.__class__.__name__=="Currency":
        return dumps(var2json(var.amount))
    elif var is None:
        return dumps(var)
    return var



if __name__ == "__main__":
    def print_lor(lor):
        print("")
        for row in lor:
            print(row)

    lor=[]
    column_to_add=[]
    for i in range(10):
        lor.append([1*i,2*i,3*i])
        column_to_add.append(-i)
    print_lor(lor)

    lor=lor_add_column(lor, 0, column_to_add)
    lor=lor_add_column(lor, 2, column_to_add)
    lor=lor_add_column(lor, 5, column_to_add)
    print_lor(lor)

    a=lor_remove_columns(lor,[2,3])
    print_lor(a)
    b=lor_remove_rows(a,[8,9])
    print_lor(b)

    c=lor_transposed(b)
    print_lor(c)

    d=Decimal("12.3")
    json_d=var2json(d)
    print (d, json_d, json_d.__class__)
    d=None
    json_d=var2json(d)
    print (d, json_d, json_d.__class__)
    d=True
    json_d=var2json(d)
    print (d, json_d, json_d.__class__)