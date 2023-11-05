from unogenerator import Range
from unogenerator.commons import get_from_process_numinstances_and_firstport

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



def test_get_from_process_numinstances_and_firstport():
    r=get_from_process_numinstances_and_firstport()
    assert len(r)==2
