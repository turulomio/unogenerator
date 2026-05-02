from unogenerator import ODS_Standard
from unogenerator.commons import Coord as C
from decimal import Decimal
from PIL import ImageColor
from ast import literal_eval
from colour import Color


def hexstring_to_rgb(color_in_hex):
    return ImageColor.getcolor(color_in_hex, "RGB")

def rgb_to_hexstring(rgb):
  r = rgb[0]
  g = rgb[1]
  b = rgb[2]
  return '#%02X%02X%02X' % (r,g,b)
  
def rgb_to_hex(rgb_tuple):
    r = rgb_tuple[0]
    g = rgb_tuple[1]
    b = rgb_tuple[2]
    s= '0x%02X%02X%02X' % (r,g,b)
#    print(literal_eval(s))
#    print("AJR", hex(literal_eval(s)))
    return hex(literal_eval(s))  
def rgb_to_int(rgb_tuple):
    r = rgb_tuple[0]
    g = rgb_tuple[1]
    b = rgb_tuple[2]
    s= '0x%02X%02X%02X' % (r,g,b)
#    print(s, r, g, b)
#    print("RGB_TO_INT", rgb_tuple,  s, literal_eval(s))
#    print("AJR", hex(literal_eval(s)))
    return literal_eval(s)


def setCellBackgroundColor(doc,  coord, color_int):
        coord=C.assertCoord(coord)
#        print("COLOR_INT", color_int)
                
        cell=doc.sheet.getCellByPosition(coord.letterIndex(), coord.numberIndex())
        cell.setPropertyValue("CellBackColor", color_int)

#def color(from_, to_, value):
#    """
#        Value entre 0 y 1
#    """
#    r1, g1, b1=from_
#    r2, g2, b2=to_
#    
##    if r1>r2:
##        r=r1-int((r1-r2)*value)
##    else:
#    r=r1+int(abs(r1-r2)*value)
##    if g1>g2:
##        g=g1-int((g1-g2)*value)
##    else:
#    g=g1+int(abs(g1-g2)*value)
##    if b1>b2:
##        b=b1-int((b1-b2)*value)
##    else:
#    b=b1+int(abs(b1-b2)*value)
#    
#    return r, g, b
#    


def list_ordered_without_none(list_):
    r=[]
    for o in list_:
        if o is None or o.__class__==str:
            continue
        r.append(o)
    r.sort()
    return r

def ordered_position_0_100(list_, item):
    """
    Function List must bue ordered
    """
    if item is None or item=="":
        return 0
    index=list_.index(item)
    return int(100*index/len(list_))
    


def percentil_color(doc, letter, skip_up=0, skip_down=0, cast=Decimal):
    """
        row_from and row_to are integers
    """
    values=doc.getValuesByColumn(letter, skip_up=skip_up, skip_down=skip_down)
    print(values, values.__class__)
        
        
#    from_=(255, 0, 0)
#    to_=(0, 255, 0)
        
    #Cast all values 
    for i in range(len(values)):
        if values[i] is None or values[i]=="":
            continue
        values[i]=cast(values[i])
    print(values)
    ordered_values=list_ordered_without_none(values)

    #Calculate percentil
    #colors=list(Color("green").range_to(Color("white"), 50))+ [Color("white")]+  list(Color("white").range_to(Color("red"), 50))
    colors=list(Color("green").range_to(Color("red"), 101))
    # Set colors
    for i in range(len(values)):
        coord=f"{letter}{i+skip_up+1}"
        i_color_hex=colors[ordered_position_0_100(ordered_values, values[i])].get_hex()#
        i_color_rgb=hexstring_to_rgb(i_color_hex)
        print(i_color_rgb)
        
        setCellBackgroundColor(doc, coord, rgb_to_int(i_color_rgb))


#################################
#
#rgb=hexstring_to_rgb("#23a9dd".upper())
#print(rgb)
#rrggbb=rgb_to_hexstring(rgb)
#print(rrggbb)
#print(rgb_to_hex(rgb))
#print("MIDDLE")
#middle=color((255, 0, 0), (0, 0, 0), 0.5)
#print(middle)
#print( rgb_to_hex(middle), rgb_to_int(middle))

Red=0xFF9999
Green=0xc0FFc0

with ODS_Standard() as doc:
    doc.createSheet("Get Values")
    doc.addColumnWithStyle("A1",  [1, 2, 3, 4, "=A4+1", 6, 7, None, 9, 10])
    percentil_color(doc, "A")
    doc.createSheet("Get Values2")
    doc.addColumnWithStyle("A1",  range(100))
    percentil_color(doc, "A")
    doc.export_pdf("percentil_color.pdf")
