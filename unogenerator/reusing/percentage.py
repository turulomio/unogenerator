## @brief my Percentage object
## THIS IS FILE IS FROM https://github.com/turulomio/reusingcode IF YOU NEED TO UPDATE IT PLEASE MAKE A PULL REQUEST IN THAT PROJECT
## DO NOT UPDATE IT IN YOUR CODE IT WILL BE REPLACED USING FUNCTION IN README

from logging import warning
from decimal import Decimal

## Class to manage percentages in spreadsheets
class Percentage:
    def __init__(self, numerator=None, denominator=None):
        self.setValue(self.toDecimal(numerator),self.toDecimal(denominator))




    def toDecimal(self, o):
        if o==None:
            return o
        if o.__class__.__name__ in ["Currency", "Money"]:
            return o.amount
        elif o.__class__.__name__=="Decimal":
            return o
        elif o.__class__.__name__ in ["int", "float"]:
            return Decimal(o)
        elif o.__class__.__name__=="Percentage":
            return o.value
        else:
            warning (o.__class__.__name__)
            return None

    def __repr__(self):
        return self.string()

    def __neg__(self):
        if self.value==None:
            return self
        return Percentage(-self.value, 1)

    def __lt__(self, other):
        if self.value==None:
            value1=Decimal('-Infinity')
        else:
            value1=self.value
        if other.value==None:
            value2=Decimal('-Infinity')
        else:
            value2=other.value
        if value1<value2:
            return True
        return False



    def __add__(self,p):
        return self.__class__(self.value+p.value,1)

    def __sub__(self, p):
        return self.__class__(self.value-p.value,1)

    def __mul__(self, value):
        if self.value==None or value==None:
            r=None
        else:
            r=self.value*self.toDecimal(value)
        return Percentage(r, 1)

    def __truediv__(self, value):
        try:
            r=self.value/self.toDecimal(value)
        except:
            r=None
        return Percentage(r, 1)

    def setValue(self, numerator,  denominator):
        try:
            self.value=Decimal(numerator/denominator)
        except:
            self.value=None

    def value_100(self):
        if self.value==None:
            return None
        else:
            return self.value*Decimal(100)

    ## @return percentage float value
    def float_100(self):
        return float(self.value_100())

    def string(self, rnd=2):
        if self.value==None:
            return "None %"
        return "{} %".format(round(self.value_100(), rnd))

    ## Returns if the percentage is valid. I mean it's value different of None
    def isValid(self):
        if self.value!=None:
            return True
        return False

    def isGETZero(self):
        if self.isValid() and self.value>=0:
            return True
        return False

    def isGTZero(self):
        if self.isValid() and self.value>0:
            return True
        return False

    def isLTZero(self):
        if self.isValid() and self.value<0:
            return True
        return False

    def qtablewidgetitem(self, decimals=2):
        from .. ui.myqtablewidget  import qpercentage
        return qpercentage(self, decimals=2)

## Calculates porcentage to pass from a to b
## @param a. Can be an object divisible and that can be substracted
def percentage_between(a,b):
    try:
        return Percentage(b-a,a)
    except:
        return Percentage()
