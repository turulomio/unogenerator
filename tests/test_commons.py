from datetime import date, datetime
from unogenerator import can_import_uno
if can_import_uno():
    from unogenerator import Range, Coord,  commons, exceptions
    from pytest import raises

    def test_coord():
        with raises(exceptions.CoordException):
            Coord(None)
        with raises(exceptions.CoordException):
            Coord(1)
        with raises(exceptions.CoordException):
            Coord("A1A")
        with raises(exceptions.CoordException):
            Coord("1")
        with raises(exceptions.CoordException):
            Coord("1A1")
        with raises(exceptions.CoordException):
            Coord("")
        with raises(exceptions.CoordException):
            Coord("A1:A2")
        with raises(exceptions.CoordException):
            Coord("Ç1:A2")
            
        Coord("A1")
        Coord("AAAAA99999")
        
        assert repr(Coord("A2"))=="Coord <A2>"
        with raises(exceptions.CoordException):
            repr(Coord(None))
        
        
        assert Coord("A1")!=Coord("A2")
        
        
        assert Coord("C3").letterPosition()==3
        assert Coord("C3").numberPosition()==3

    def test_coord_assertcoord():
        with raises(exceptions.CoordException):
            Coord.assertCoord(None)
        with raises(exceptions.CoordException):
            Coord.assertCoord(1)
        with raises(exceptions.CoordException):
            Coord.assertCoord("A1A")
        with raises(exceptions.CoordException):
            Coord.assertCoord("1")
        with raises(exceptions.CoordException):
            Coord.assertCoord("1A1")
        with raises(exceptions.CoordException):
            Coord.assertCoord("")
        with raises(exceptions.CoordException):
            Coord.assertCoord("A1:A2")
        Coord.assertCoord("A1")
        Coord.assertCoord("AAAAA99999")
            
    def test_range():
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
        
        range_=Range.from_iterable_object("A1", [])
        assert range_.string()=="A1:A1"

        range_=Range.from_iterable_object("A1", [1, 2, 3])
        assert range_.string()=="A1:C1"
        
        with raises(exceptions.RangeException):
            range_=Range("1A1:None")
            
        with raises(exceptions.RangeException):
            range_=Range("A1:A1A")
        
        with raises(exceptions.RangeException):
            range_=Range(None)
        
        with raises(exceptions.RangeException):
            range_=Range("A1")
            
        assert not Range("A1:A2")==Range("A2:A2")
        
        assert repr(Range("A1:A2"))=="Range <A1:A2>"
            
    def test_range_assertrange():
        assert Range.assertRange("A1:C3").string()=="A1:C3"
        
        with raises(exceptions.RangeException):
            Range.assertRange("1A1:None")
        
        with raises(exceptions.RangeException):
            Range.assertRange(None)
        
        with raises(exceptions.RangeException):
            Range.assertRange("A1")

    def test_range_operations():
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
        
    def test_range_contains():
        range_= Range("A2:C5")
        
        assert Coord("A2") in range_
        assert Coord("A1") not in range_
        
    def test_range_indexes_list():
        range_= Range("A1:B3")
        r=range_.indexes_list()
        assert len(r)==3
        assert r[2][1]==(1, 2)
        r=range_.indexes_list(plain=True)
        assert len(r)==6
        assert r[5]==(1, 2)
        
        
    def test_range_coords_list():
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


    def test_localc1989():
        date_=date(2023,12,2)
        datetime_=datetime(2023,12,2,12,0,0)
        date_localc1989=45262.0
        datetime_localc1989=45262.5

        assert commons.datetime2localc1989(datetime_)==datetime_localc1989
        assert commons.localc19892datetime(datetime_localc1989)==datetime_
        assert commons.date2localc1989(date_)==date_localc1989
        assert commons.localc19892date(date_localc1989)==date_
