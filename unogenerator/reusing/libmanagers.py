## @brief Module with objects managers as list or as dictionary.
## THIS IS FILE IS FROM https://github.com/turulomio/reusingcode IF YOU NEED TO UPDATE IT PLEASE MAKE A PULL REQUEST IN THAT PROJECT
## DO NOT UPDATE IT IN YOUR CODE IT WILL BE REPLACED USING FUNCTION IN README
##
## You have to use list objects if you are going to make selections and secuential access.

from datetime import datetime, timedelta, date
from logging import critical, debug
from .datetime_functions import dtaware_day_end_from_date, dtaware_day_start_from_date, dtnaive_day_end_from_date, dtnaive_day_start_from_date
from .call_by_name import call_by_name

## Defines who self.selected is managed
## If can take the following values
## - Object self.selected is a object
## - List self.selected is a list of objects
## - Manager. Selef selected is an object
class ManagerSelectionMode:
    Object=0
    List=1
    Manager=2

class ObjectManager(object):
    def __init__(self):
        self.arr=[]

    def __repr__(self):
        return "{} with {} objects".format(self.__class__.__name__, self.length())

    ## Method to iterate self.arr iterating object
    def __iter__(self):
        return iter(self.arr)

    ## Store constructor params to allow create new instances of this managers
    def setConstructorParameters(self, *params):
        self.initparams=params

    def append(self,  obj):
        self.arr.append(obj)

    ## Appends and object if it isn't in array. Array will act as a set()
    def append_distinct(self, obj):
        if obj not in self.arr:
            self.append(obj)

    def remove(self, obj):
        self.arr.remove(obj)

    ## Needs the setConstructorParameters before
    def emptyManager(self):
        return self.__class__(*self.initparams)

    def length(self):
        return len(self.arr)

    def clean(self):
        """Deletes all items"""
        self.arr=[]

    def clone(self):
        """Returns other Set object, with items referenced, ojo con las formas de las instancias
        initparams son los parametros de iniciación de la clase"""
        result=self.emptyManager()
        for a in self.arr:
            result.append(a)
        return result
   
    def first(self):
        return self.object(0)

    def index(self,o):
        return self.arr.index(o)

    def last(self):
        return self.object(self.length()-1)

    def print(self):
        print ("Objects in {}".format(self.__class__))
        for q in self.arr:
            print(" * {}".format(q))
            
    ## Return the object at the index position
    def object(self, index):
        try:
            return self.arr[index]
        except:
            critical("I couldn't retrive object from {} position".format(index))

    ## Order data columns. None values are set at the beginning
    def order_with_none(self, string_or_tuple, reverse=False, none_at_top=True):
        nonull=[]
        null=[]
        for o in self.arr:
            com=call_by_name(o, string_or_tuple)
            if com is None:
                null.append(o)
            else:
                nonull.append(o)
        nonull=sorted(nonull, key=lambda c: call_by_name(c,string_or_tuple), reverse=reverse)
        if none_at_top==True:#Set None at top of the list
            self.arr=null+nonull
        else:
            self.arr=nonull+null

    ## @param string_or_tuple String or tuple used with a call_by_name method
    ## @return List returned with a call_by_name string or tuplen array with all object ids
    def list_of(self, string_or_tuple):
        r=[]
        for o in self:
            r.append(call_by_name(o, string_or_tuple))
        return r

    ## Returns a new manager with the objects that have found a list of strings in several commands, passed as 
    ## This function doen't make exact matches. Only if strings in s_list are contained
    ## @param string_or_tuple_list List of _string_or_tuple_to_command parameters
    ## @param s_list List of string to search
    ## @param upper boolean
    def find_strings_contained_in_string_or_tuple_results(self, string_or_tuple_list, s_list, upper=False):
        r=self.emptyManager()
        for o in self.arr:
            string_=""
            for string_or_tuple in string_or_tuple_list:
                string_=string_+str(call_by_name(o, string_or_tuple))
                
            for s in s_list:
                if upper==True:
                    if s.upper() in string_.upper():
                        r.append(o)
                        break
                else:#upper False
                    if s in string_:
                        r.append(o)
                        break
        return r
        
    def find_string_exact_in_string_or_tuple_results(self, string_or_tuple, s, upper=False):        
        r=self.emptyManager()
        for o in self.arr:
            if upper==True:
                if s.upper() == call_by_name(o, string_or_tuple).upper():
                    r.append(o)
            else:#upper False
                if s == call_by_name(o, string_or_tuple):
                    r.append(o)
        return r


    ## Search by id iterating array
    def find_by_id_builtin(self, id_, logging=False):
        start=datetime.now()
        for o in self.arr:
            if id(o)==id_:
                if logging==True:
                    debug("{} took {} to find by id builtin {} with list".format(self.__class__.__name__, datetime.now()-start, id))
                return o
        return None


    ## It doesn't emit selected if selected is None a nd needtoselect is False, in the rest of the cases it emmit itemChanged
    ## @param combo
    ## @param selected it's an object
    ## @param needtoselect Adds a foo item with value==None with the text select one
    ## @param icons Boolean. If it's true uses o.qicon() method to add an icon to the item
    ## @param id_attr. Atribute to call_by_name for id. If None uses id(o). Normally will user "id"
    ## @param name_attr. Atribute to call_by_name for name. If None user str(o). Normally will use "name"
    
    def qcombobox(self, combo,  selected=None, needtoselect=False, icons=False, id_attr="id", name_attr="name"):
        combo.blockSignals(True)
        combo.clear()

        #Add items
        if needtoselect==True:
            if self.length()>0:
                combo.addItem(combo.tr("Select an option"), None)
            else:
                combo.addItem(combo.tr("No options to select"), None)
        for a in self.arr:
            print(a,id_attr,name_attr)
            id_  =call_by_name(a, id_attr) if id_attr is not None else id(a)
            name_=call_by_name(a, name_attr) if name_attr is not None else str(a)
            if icons==True:
                combo.addItem(a.qicon(), name_, id_)
            else:
                combo.addItem(name_, id_)

        #Force without signals to be in -1. There were problems when 0 is selected, becouse it didn't emit anything
        combo.setCurrentIndex(-1)

        #Set selection
        if needtoselect==True:
            combo.blockSignals(False)
            combo.setCurrentIndex(0)
        else:#need to select False
            if selected is None:
                combo.blockSignals(False)
            else:
                combo.blockSignals(False)
                if id_attr is None:
                    selected_id=id(id_attr)
                else:
                    selected_id=selected.id
                combo.setCurrentIndex(combo.findData(selected_id))


## Manager Selection class
## By default selectionmode is
class ManagerSelection(object):
    def __init__(self):
        self.__selected=None
        self.__selectionmode=ManagerSelectionMode.Object

    @property
    def selected(self):
        return self.__selected

    @selected.setter
    def selected(self, value):
        self.__selected=value
        
    def selectionMode(self):
        return self.__selectionmode

    ## @param value ManagerSelectionMode
    def setSelectionMode(self, value):
        self.__selectionmode=value
        if value==ManagerSelectionMode.Object:
            self.selected=None
        elif value==ManagerSelectionMode.List:
            self.selected=[]
        elif value==ManagerSelectionMode.Manager:
            self.selected=self.emptyManager()

    def cleanSelection(self):
        if self.selectionMode()==ManagerSelectionMode.Object:
            self.selected=None
        elif self.selectionMode()==ManagerSelectionMode.List:
            self.selected=[]
        elif self.selectionMode()==ManagerSelectionMode.Manager:
            self.selected.clean()

    ## Useful to setselection without interactivvite ui
    ## @param list Can be, list, manager or object
    def setSelected(self, list):
        self.cleanSelection()
        if self.selectionMode() in (ManagerSelectionMode.List, ManagerSelectionMode.Manager):
            for o in list:
                self.selected.append(o)
        else:#Object
            self.selected=list
    
    ## Converts selection to a manager. With this function I removed ManagerSelectionMode.Manager.
    def convertSelectionToManager(self):
        r=self.emptyManager()
        if self.selectionMode()==ManagerSelectionMode.List:
            r.arr=self.selected
        else:#Object
            r.append(self.selected)
        return r

## Objects in DictListObjectManager has and id. The Id can be a integer or a string or ...
class ObjectManager_With_Id(ObjectManager):
    def __init__(self):
        ObjectManager.__init__(self)
        self._find_dict={}
        self._use_dict_to_find=False

    ## If set to True enables the use of a dict to find by id
    def setUseDictToFind(self, value):
        self._use_dict_to_find=value

    def append(self,  obj):
        self.arr.append(obj)
        if self._use_dict_to_find==True:
            if obj.id is None:
                debug("You have added a key None to self._find_dict, perhaps you need to append it when object.id is set")
            self._find_dict[obj.id]=obj

    def remove(self, obj):
        self.arr.remove(obj)
        if self._use_dict_to_find==True:
            del self._find_dict[obj.id]


    def arr_position(self, id):
        """Returns arr position of the id, useful to select items with unittests"""
        for i, a in enumerate(self.arr):
            if a.id==id:
                return i
        return None

    ##Returns an array with all object ids
    def array_of_ids(self):
        r=[]
        for o in self.arr:
            r.append(o.id)
        return r

    ## Search by id iterating array
    def find_by_id(self, id, logging=False):
        if self._use_dict_to_find==True:
            try:
                return self._find_dict[id]
            except:
                return None
        else:
            start=datetime.now()
            for o in self.arr:
                if o.id==id:
                    if logging==True:
                        debug("{} took {} to find by id {} with list".format(self.__class__.__name__, datetime.now()-start, id))
                    return o
            return None

    def order_by_id(self, reverse=False, none_at_top=True):
        self.order_with_none("id", reverse=reverse, none_at_top=none_at_top)

    def union(self,  set,  *initparams):
        """Returns a new set, with the union comparing id
        initparams son los parametros de iniciación de la clse"""        
        resultado=self.__class__(*initparams)#Para que coja la clase del objeto que lo invoca SetProduct(self.mem), luego será self.mem
        for p in self.arr:
            resultado.append(p)
        for p in set.arr:
            if resultado.find_by_id(p.id, False)==None:
                resultado.append(p)
        return resultado


## Objects in DictListObjectManager has and id and a date attribute
class ObjectManager_With_IdDate(ObjectManager_With_Id):
    def __init__(self):
        ObjectManager_With_Id.__init__(self)

    def order_by_date(self, reverse=False, none_at_top=True):
        self.order_with_none("date", reverse=reverse, none_at_top=none_at_top)
        
## Objects in DictListObjectManager has and id and a datetime
class ObjectManager_With_IdDatetime(ObjectManager_With_Id):
    def __init__(self):
        ObjectManager_With_Id.__init__(self)

    def order_by_datetime(self, reverse=False, none_at_top=True):
        self.order_with_none("datetime", reverse=reverse, none_at_top=none_at_top)
                
    ## Function that returns the same object manager, with a pointer to the of the objects that contains from the datetime given in the parameter.
    ## For example the constuctor of InvemestOperationHomogeneous is InvesmentOperationHomogeneous(mem,investment). so to use this function you need ObjectManager_from_datetime(dt,mem,investment)
    ## @param dt datetime This function copies all object with datetime until this parameter
    def ObjectManager_from_datetime(self, dt):
        result=self.emptyManager()
        for a in self.arr:
            if a.datetime>=dt:
                result.append(a)
        return result         
        
    ## Function that returns the same object manager, with a pointer to the of the objects that contains from the datetime given in the parameter.
    ## For example the constuctor of InvemestOperationHomogeneous is InvesmentOperationHomogeneous(mem,investment). so to use this function you need ObjectManager_from_datetime(dt,mem,investment)
    ## @param datetime. This function copies all object with datetime until this parameter
    def ObjectManager_until_datetime(self, dt):
        result=self.emptyManager()
        for a in self.arr:
            if a.datetime<=dt:
                result.append(a)
        return result
        
    def ObjectManager_copy_from_datetime(self, dt):
        result=self.emptyManager()
        for a in self.arr:
            if a.datetime>=dt:
                result.append(a.copy())
        return result
        
    ## Function that returns the same object manager, but with a copy of the objects that contains until the datetime given in the parameter.
    ## For exemple the constuctor of InvemestOperationHomogeneous is InvemestOperationHomogeneous(mem,investment). so to use this function you need ObjectManager_copy_until_datetime(dt,mem,investment)
    ## @param dt datetime This function copies all object with datetime until this parameter
    def ObjectManager_copy_until_datetime(self, dt):
        result=self.emptyManager()
        for a in self.arr:
            if a.datetime<=dt:
                result.append(a.copy())
        return result

## Objects in DictListObjectManager has and id and a name
class ObjectManager_With_IdName(ObjectManager_With_Id):
    def __init__(self):
        ObjectManager_With_Id.__init__(self)
        
    ##Returns an array with all object name
    ## @param sort Boolean to sort or not the array
    ## @param nones Boolean. If True adds None and empty strings values to the list. If False it doesn't
    def array_of_names(self, sort=True, nones=False):
        r=[]
        for o in self.arr:
            if nones==False:
                if o.name==None or o.name=="":
                    continue
            r.append(o.name)
        if sort==True:
            r.sort()
        return r

    def find_by_name(self, s,  upper=False):
        return self.find_string_exact_in_string_or_tuple_results("name", s, upper)

    ## Returns another object manager of the same class with the elements that contains a string in the name
    ## @param s string to search
    ## @param casesensitive Boolean if it's a case sensitive search    
    def ObjectManager_which_name_contains(self, s, casesensitive):
        result=self.emptyManager()
        if s is None:
            critical("Search string can't be None")
            return result
        if casesensitive==True:
            for a in self.arr:
                if s in a.name:
                    result.append(a)
            return result
        else:
            for a in self.arr:
                if s.upper() in a.name.upper():
                    result.append(a)
            return result
        
    def order_by_name(self, reverse=False, none_at_top=True):
        self.order_with_none("name", reverse=reverse, none_at_top=none_at_top)

    def order_by_upper_name(self, reverse=False, none_at_top=True):
        self.order_with_none(("name.upper", []), reverse=reverse, none_at_top=none_at_top)


    ## Creates a libreoffice sheet from the ObjectManager
    ##
    ## This function needs the officegenerator package
    ## @param ods Officegenerator ODS_Write object
    ## @param sheetname String with the name of the libreoffice sheet
    ## @param titles List of strings with the titles of the columns
    ## @param order_by_name Boolean. True: orders by name. False: orders by id
    ## @returns Officegenerator OdfSheet
    def ods_sheet(self, ods, sheetname, titles=["Id", "Name"],  order_by_name=True):
        from officegenerator import Coord
        if order_by_name==True:
            self.order_by_name()
        else:
            self.order_by_id()
        s=ods.createSheet(sheetname)
        s.setColumnsWidth([80, 240])
        s.add("A1", [titles], "OrangeCenter")
        for number, o in enumerate(self.arr):
            s.add(Coord("A2").addRow(number), o.id, "WhiteRight")        
            s.add(Coord("B2").addRow(number), o.name, "WhiteLeft")
        s.freezeAndSelect("A1")
        return s

## Usefull when creating a class with two attributes self.id and self.name only
class Object_With_IdName:
    ## Constructor with the following attributes combination
    ## 1. Object_With_IdName(). Create an Object_With_IdName with all attributes to None
    ## 2. Object_With_IdName( id,  name). Create an Object_With_IdName settings all attributes.
    ## @param name String with the name of the Object_With_IdName
    ## @param id Integer that sets the id of the Object_With_IdName
    def __init__(self, id=None,  name=None):
            self.id=id
            self.name=name

    def __repr__(self):
        return "{}: {} (Id: {})".format(self.__class__.__name__, self.name, self.id)

    def upper_name(self, boolean):
        if self.name==None:
            return None
        if boolean==True:
            return self.name.upper()
        else:
            return self.name.lower()

class ObjectManager_Selectable(ObjectManager, ManagerSelection):
    def __init__(self):
        ObjectManager.__init__(self)
        ManagerSelection.__init__(self)

class ObjectManager_With_Id_Selectable(ObjectManager_With_Id, ManagerSelection):
    def __init__(self):
        ObjectManager_With_Id.__init__(self)
        ManagerSelection.__init__(self)

class ObjectManager_With_IdDate_Selectable(ObjectManager_With_IdDate, ManagerSelection):
    def __init__(self):
        ObjectManager_With_IdDate.__init__(self)
        ManagerSelection.__init__(self)

class ObjectManager_With_IdDatetime_Selectable(ObjectManager_With_IdDatetime, ManagerSelection):
    def __init__(self):
        ObjectManager_With_IdDatetime.__init__(self)
        ManagerSelection.__init__(self)

class ObjectManager_With_IdName_Selectable(ObjectManager_With_IdName, ManagerSelection):
    def __init__(self):
        ObjectManager_With_IdName.__init__(self)
        ManagerSelection.__init__(self)




## THIS IS A NEW SERIE OF MANAGERS DATETIME VALUE
class DatetimeValue:
    def __init__(self):
        self.datetime=None
        self.value=None

    def __repr__(self):
        return "DatetimeValue {} = {}".format(self.date,self.value)

class DatetimeValueManager(ObjectManager):
    def __init__(self):
        ObjectManager.__init__(self)
    def appendDV(self, datetime, value):
        o=DatetimeValue()
        o.datetime=datetime
        o.value=value
        self.append(o)    ## Returns a date value manager with the simple movil average 3 of weight
    
    ## Returns a DVManager with the simple movil average of the array
    def sma(self, period):
        r=DatetimeValueManager()
        for i in range(period, self.length()):
            sma=DatetimeValue()
            sma.value=0
            sma.datetime=self.arr[i].datetime
            for p in range(period):
                sma.value=sma.value+self.arr[i-p].value
            sma.value=sma.value/period
            r.append(sma)
        return r

    ## Return find less or equal dv of a given datetime
    ## @param dt datetime
    def find_le(self, dt):
        for o in reversed(self.arr):
            if o.datetime<=dt:
                return o
        return None

## THIS IS A NEW SERIE OF MANAGERS DATE VALUE
class DateValue:
    def __init__(self):
        self.date=None
        self.value=None

    def __repr__(self):
        return "DateValue {} = {}".format(self.date,self.value)

class DateValueManager(ObjectManager):
    def __init__(self):
        ObjectManager.__init__(self)

    def appendDV(self, date, value):
        o=DateValue()
        o.date=date
        o.value=value
        self.append(o)    ## Returns a date value manager with the simple movil average 3 of weight
        
    ## Fills days without data  with the data before or after
    def DateValueManager_filling_empty(self, with_data_before=True):
        r=DateValueManager()
        last=self.first()
        r.append(last)
        for o in self.arr[1:]:
            while o.date!=last.date+timedelta(days=1):
                missing=DateValue()
                missing.date=last.date+timedelta(days=1)
                missing.value=last.value
                r.append(missing)
                last=missing
            r.append(o)
            last=o
            
        return r
        
    ## Converts a DateValueManager to a DatetimeValueManager.
    ## @param start Boolean. If true date is converted to the start of the day. If false to the end of the day
    ## @param timezone String with pytz timexone. If None datetimes will be naive, else datetimes will be aware
    def DatetimeValueManager(self, start=True, timezone=None):
        r=DatetimeValueManager()
        for o in self.arr:
            if start==True:
                if timezone==None:
                    r.appendDV(dtnaive_day_start_from_date(o.date), o.value)
                else:
                    r.appendDV(dtaware_day_start_from_date(o.date, timezone), o.value)
            else:#end of day
                if timezone==None:
                    r.appendDV(dtnaive_day_end_from_date(o.date), o.value)
                else:
                    r.appendDV(dtaware_day_end_from_date(o.date, timezone), o.value)
        return r

    ## Returns a DVManager with the simple movil average of the array
    def sma(self, period):
        r=DateValueManager()
        for i in range(period, self.length()):
            sma=DateValue()
            sma.value=0
            sma.date=self.arr[i].date
            for p in range(period):
                sma.value=sma.value+self.arr[i-p].value
            sma.value=sma.value/period
            r.append(sma)
        return r



if __name__ == "__main__":
    print("TESTING FILLING DATEVALUEMANAGER")
    r=DateValueManager()
    r.appendDV(date(2019, 1, 10), 1)
    r.appendDV(date(2019, 1, 17), 2)
    r.print()
    filled=r.DateValueManager_filling_empty()
    filled.print()

    sizes=(1,10,100,1000,10000,100000)#,1000000,3000000)
    for size in sizes:
        l=ObjectManager_With_Id()
        for number in range(size):
            o=Object_With_IdName(number,"Name {}".format(number))
            l.append(o)
        middle=size*2//3
        start=datetime.now()
        l.find_by_id(middle)
        print("Benchmarking search_by_id in to element {} with {} objects took {}".format(middle,size,datetime.now()-start))
        
    print(l.find_by_id(500))


    print()
    print("Ordering with None")
    manager=ObjectManager_With_IdName()
    manager.setConstructorParameters()
    manager.append(Object_With_IdName(12, None))
    manager.append(Object_With_IdName(23,"AB"))
    manager.append(Object_With_IdName(34,"BC"))
    manager.append(Object_With_IdName(45,"CD"))
    manager.append(Object_With_IdName(None,"DE"))
    manager.print()
    manager.order_with_none(["upper_name", (True,)], False, False)
    manager.print()

    print()
    print("Finding objects")
    find_name=manager.find_strings_contained_in_string_or_tuple_results(
                            [
                                ("upper_name", (True, )),  #tuple
                                "id" #string
                            ],  
                            ["2", "C"] #Words to find
                        )
    find_name.print()
    find_name=manager.find_by_name("CD")
    find_name.print()

    print("Iterating an object")
    for o in manager:
        print(o)
