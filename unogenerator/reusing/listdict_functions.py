## THIS IS FILE IS FROM https://github.com/turulomio/reusingcode IF YOU NEED TO UPDATE IT PLEASE MAKE A PULL REQUEST IN THAT PROJECT
## DO NOT UPDATE IT IN YOUR CODE IT WILL BE REPLACED USING FUNCTION IN README

from .casts import var2json




## El objetivo es crear un objeto list_dict que se almacenera en self.ld con funciones set
## set_from_db #Todo se carga desde base de datos con el minimo parametro posible
## set_from_db_and_variables #Preguntara a base datos aquellas variables que falten. Aunque no estén en los parámetros p.e. money_convert
## set_from_variables #Solo con variables
## set #El listdict ya está hecho pero se necesita el objeto para operar con el
##class Do:
##    def __init__(self,d):
##        self.d=d
##        self.create_attributes()
##
##    def number_keys(self):
##        return len(self.d)
##
##    def has_key(self,key):
##        return key in self.d
##
##    def print(self):
##        listdict_print(self.d)
##
##    ## Creates an attibute from a key
##    def create_attributes(self):
##        for key, value in self.d.items():
##            setattr(self, key, value)




## Class that return a object to manage listdict
## El objetivo es crear un objeto list_dict que se almacenera en self.ld con funciones set
## set_from_db #Todo se carga desde base de datos con el minimo parametro posible
## set_from_db_and_variables #Preguntara a base datos aquellas variables que falten. Aunque no estén en los parámetros p.e. money_convert
## set_from_variables #Solo con variables
## set #El listdict ya está hecho pero se necesita el objeto para operar con el
class Ldo:
    def __init__(self, name=None):
        self.name=self.__class__.__name__ if name is None else name
        self.ld=[]

    def length(self):
        return len(self.ld)

    def has_key(self,key):
        return listdict_has_key(self.ld,key)

    def print(self):
        listdict_print(self.ld)

    def print_first(self):
        listdict_print_first(self.ld)

    def sum(self, key, ignore_nones=True):
        return listdict_sum(self.ld, key, ignore_nones)

    def list(self, key, sorted=True):
        return listdict2list(self.ld, key, sorted)

    def average_ponderated(self, key_numbers, key_values):
        return listdict_average_ponderated(self.ld, key_numbers, key_values)

    def set(self, ld):
        del self.ld
        self.ld=ld
        return self

    def is_set(self):
        if hasattr(self, "ld"):
            return True
        print(f"You must set your listdict in {self.name}")
        return False

    def append(self,o):
        self.ld.append(o)

    def first(self):
        return self.ld[0] if self.length()>0 else None

    ## Return list keys of the first element[21~
    def first_keys(self):
        if self.length()>0:
            return self.first().keys()
        else:
            return "I can't show keys"
    
    def order_by(self, key, reverse=False):
        self.ld=sorted(self.ld,  key=lambda item: item[key], reverse=reverse)
        
    def json(self):
        return listdict2json(self.ld)

def listdict_has_key(listdict, key):
    if len(listdict)==0:
        return False
    return key in listdict[0]


## Order data columns. None values are set at the beginning
def listdict_order_by(ld, key, reverse=False, none_at_top=True):
    nonull=[]
    null=[]
    for o in ld:
        com=o[key]
        if com is None:
            null.append(o)
        else:
            nonull.append(o)
    nonull=sorted(nonull, key=lambda item: item[key], reverse=reverse)
    if none_at_top==True:#Set None at top of the list
        return null+nonull
    else:
        return nonull+null



def listdict_print(listdict):
    for row in listdict:
        print(row)

def listdict_print_first(listdict):
    if len(listdict)==0:
        print("No rows in listdict")
        return
    print("Printing first dict in a listdict")
    keys=list(listdict[0].keys())
    keys.sort()
    for key in keys:
        print(f"    - {key}: {listdict[0][key]}")

def listdict_sum(listdict, key, ignore_nones=True):
    r=0
    for d in listdict:
        if ignore_nones is True and d[key] is None:
            continue
        r=r+d[key]
    return r




def listdict_sum_negatives(listdict, key):
    r=0
    for d in listdict:
        if d[key] is None or d[key]>0:
            continue
        r=r+d[key]
    return r

def listdict_sum_positives(listdict, key):
    r=0
    for d in listdict:
        if d[key] is None or d[key]<0:
            continue
        r=r+d[key]
    return r


def listdict_average(listdict, key):
    return listdict_sum(listdict,key)/len(listdict)

def listdict_average_ponderated(listdict, key_numbers, key_values):
    prods=0
    for d in listdict:
        prods=prods+d[key_numbers]*d[key_values]
    return prods/listdict_sum(listdict, key_numbers)


def listdict_median(listdict, key):
    from statistics import median
    return median(listdict2list(listdict, key, sorted=True))


## Converts a listdict to a dict using key as new dict key
def listdict2dict(listdict, key):
    d={}
    for ld in listdict:
        d[ld[key]]=ld
    return d

## Returns a list from a listdict key
## @param listdict
## @param key String with the key to extract
## @param sorted Boolean. If true sorts final result
## @param cast String. "str", "float", casts the content of the key
def listdict2list(listdict, key, sorted=False, cast=None):
    r=[]
    for ld in listdict:
        if cast is None:
            r.append(ld[key])
        elif cast == "str":
            r.append(str(ld[key]))
        elif cast == "float":
            r.append(float(ld[key]))
    if sorted is True:
        r.sort()
    return r

def listdict2json(listdict):
    if len(listdict)==0:
        return "[]"

    r="["
    for o in listdict:
        d={}
        for field in o.keys():
            d[field]=var2json(o[field])
        r=r+str(d).replace("': 'null'", "': null").replace("': 'true'", "': true").replace("': 'false'", "': false") +","
    r=r[:-1]+"]"
    return r

## Returns the max of a key in listdict
def listdict_max(listdict, key):
    return max(listdict2list(listdict,key))

## Returns the min of a key in listdict
def listdict_min(listdict, key):
    return min(listdict2list(listdict,key))


if __name__ == "__main__":
    from datetime import datetime, date
    from decimal import Decimal
    ld=[]
    ld.append({"a": datetime.now(), "b": date.today(), "c": Decimal(12.32), "d": None, "e": int(12), "f":None, "g":True, "h":False})
    print(listdict2json(ld))


    def print_lor(lor):
        print("")
        for row in lor:
            print(row)