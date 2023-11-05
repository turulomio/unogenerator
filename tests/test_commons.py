from unogenerator import Range


class TestRange:
    def test_constructors(self):
        range_=Range("A1:C3")
        assert range_.c_start=="A1" 
        
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
