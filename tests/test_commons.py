from unogenerator import Range, Coord,  commons

class TestCoord:
    def test_constructors(self):
        Coord(1)
        Coord.assertCoord(None)
        
class TestRange:
    def test_constructors(self):
        range_=Range("A1:C3")
        assert range_.c_start=="A1" 
        
        range_=Range.from_coords_indexes(0, 0, 3, 3)
        assert range_.string()=="A1:D4"
        
        range_=Range.from_coords("A1", "D4")
        assert range_.string()=="A1:D4"
        
        range_=Range.from_start_coord_change(range_, "B2")
        assert range_.string()=="B2:E5"
        
        range_=Range.from_iterable_object("A1", [[1, 2], [3, 4]])
        assert range_.string()=="A1:B2"

    def test_operations(self):
        range_= Range("A2:C5")
        
        range_=range_.addRowBefore(1)
        assert str(range_)=="A1:C5"
        
        range_=range_.addRowBefore(-1)
        assert str(range_)=="A2:C5"
        
        range_=range_.addRowAfter(-1)
        assert str(range_)=="A2:C4"
        
        range_=range_.addRowAfter(1)
        assert str(range_)=="A2:C5"
        
        range_=range_.addRowAfterCopy()
        assert str(range_)=="A2:C6"
        
        range_=range_.addRowBeforeCopy(1)
        assert str(range_)=="A1:C6"
        
        range_=range_.addColumnAfter()
        assert str(range_)=="A1:D6"
        
        range_=range_.addColumnAfterCopy()
        assert str(range_)=="A1:E6"
        
        range_=range_.addColumnBeforeCopy()
        assert str(range_)=="A1:E6"
        
    def test_contains(self):
        range_= Range("A2:C5")
        
        assert Coord("A2") in range_
        assert Coord("A1") not in range_
        
    def test_indexes_list(self):
        range_= Range("A1:B3")
        r=range_.indexes_list()
        assert len(r)==3
        assert r[2][1]==(1, 2)
        r=range_.indexes_list(plain=True)
        assert len(r)==6
        assert r[5]==(1, 2)
        
        
    def test_coords_list(self):
        range_= Range("A1:B3")
        r=range_.coords_list()
        assert len(r)==3
        assert r[2][1].string()=="B3"
        r=range_.coords_list(plain=True)
        assert len(r)==6
        assert r[5].string()=="B3"
        


def test_are_all_values_of_the_list_the_same():
    assert commons.are_all_values_of_the_list_the_same([])==True
    assert commons.are_all_values_of_the_list_the_same([1, 1, 1, 1])==True
    assert commons.are_all_values_of_the_list_the_same([1, 2])==False

def test_number2row():
    assert commons.number2row(5)=="5"


def test_get_from_process_numinstances_and_firstport():
    r=commons.get_from_process_numinstances_and_firstport()
    assert len(r)==2

def test_colorama():
    commons.red("s")
    commons.green("s")
    commons.magenta("s")
    
def test_is_formula():
    assert commons.is_formula(None) is False
    assert commons.is_formula("s") is False
    assert commons.is_formula("=SUM")
    assert commons.is_formula("+SUM")
