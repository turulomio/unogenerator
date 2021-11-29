from datetime import datetime
from logging import debug
from unogenerator.commons import ColorsNamed, Coord as C
from unogenerator.helpers import helper_totals_row

from gettext import translation
from pkg_resources import resource_filename

try:
    t=translation('unogenerator', resource_filename("unogenerator","locale"))
    _=t.gettext
except:
    _=str

class Statistics:
    def __init__(self, doc):
        self.init=datetime.now()
        self.doc=doc
        
class StatisticsODS(Statistics):
    def __init__(self, doc):
        Statistics.__init__(self, doc)
        self.cells=[]
        self.cells_merged=[]
        self.cells_gets=[]
        self.sheet_freezes=[]
        self.sheet_creations=[]
        
    def __del__(self):
        self.debug()
        
    def debug(self):
        debug(f"- Number of cells {len(self.cells)}")
        debug(f"- Number of merged cells {len(self.cells_merged)}")
        debug(f"- Number of get values {len(self.cells_gets)}")
        debug(f"- Total time {datetime.now()-self.init}")
        
    def appendCellCreationStartMoment(self, start):
        self.cells.append(datetime.now()-start)
        if len(self.cells) % 500==0:
            debug(f"Wrote {len(self.cells)} cells in {datetime.now()-self.init}")

    def appendCellMergedCreationStartMoment(self, start):
        self.cells_merged.append(datetime.now()-start)

    def appendSheetFreezesCreationStartMoment(self, start):
        self.sheet_freezes.append(datetime.now()-start)

    def appendSheetCreationsCreationStartMoment(self, start):
        self.sheet_creations.append(datetime.now()-start)
        
    def appendCellGetValuesStartMoment(self, start):
        self.cells_gets.append(datetime.now()-start)

    ## Creates a new sheet called "Style names" with alll ods styles grouped by families
    def ods_sheet_statistics(self):
        self.doc.createSheet("Internal statistcs")
        self.doc.setColumnsWidth([5]+[3]*10)
        self.doc.addCellWithStyle("A1", "Document generation", ColorsNamed.Orange, "Bold")
        self.doc.addCellMergedWithStyle("B1:C1", datetime.now()-self.init, ColorsNamed.Yellow, "Right")
        
        
        self.doc.addRowWithStyle("A3", [_("Concept"), _("Number"), _("Total seconds"), _("Average")],  ColorsNamed.Orange, "BoldCenter")
        self.doc.addColumnWithStyle("A4", [_("Cells"),  _("Merged cells"), _("Cell gets"), _("Sheet creations"), _("Sheet freezes")], ColorsNamed.Green)
        for row,  l in enumerate([self.cells, self.cells_merged, self.cells_gets, self.sheet_creations, self.sheet_freezes]):
            self.doc.addCellWithStyle(C("B4").addRow(row), len(l))
            self.doc.addCellWithStyle(C("C4").addRow(row), self.timedelta_total_seconds(l))
            self.doc.addCellWithStyle(C("D4").addRow(row), self.timedelta_average(l), style="Float6")
            
        helper_totals_row(self.doc, C("A5").addRow(row), ["Total", "#SUM","#SUM",""], styles=["BoldCenter","Integer","Float2","Integer"], row_from="4")
        
    def timedelta_total_seconds(self, l):
        total=0
        for td in l:
            total=total+td.total_seconds()
        return total
        
    def timedelta_average(self, l):
        if len(l)==0:
            return 0
        return self.timedelta_total_seconds(l)/len(l)
            
        
    

class StatisticsODT(Statistics):
    def __init__(self, doc):
        Statistics.__init__(self, doc)
        
