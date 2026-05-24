
import pytest
from unogenerator import ODS_Standard, LibreofficeServer
from unogenerator.commands import copy_sheet
import os

def test_command_copy_sheet(libreoffice_server, tmp_path):
    src_file = str(tmp_path / "source.ods")
    dest_file = str(tmp_path / "destination.ods")
    
    # 1. Create source with data and style
    with ODS_Standard(server=libreoffice_server) as doc:
        doc.addCellWithStyle("A1", "Styled Data", color=0xFF0000) # Red background
        doc.save(src_file)
        
    # 2. Copy to a new destination
    copy_sheet(src_file, "Hoja1", dest_file, "CopiedSheet", server=libreoffice_server)
    
    # 3. Verify
    assert os.path.exists(dest_file)
    from unogenerator import ODS
    with ODS(dest_file, server=libreoffice_server) as doc_check:
        doc_check.setActiveSheet("CopiedSheet")
        val = doc_check.getValue("A1", detailed=True)
        assert val["value"] == "Styled Data"
        assert val["style"] == "Normal" # ODS_Standard uses Normal
        # Check color (converted to int)
        assert doc_check.sheet.getCellByPosition(0,0).CellBackColor == 0xFF0000

def test_copy_sheet_existing_dest(libreoffice_server, tmp_path):
    src_file = str(tmp_path / "source2.ods")
    dest_file = str(tmp_path / "destination2.ods")
    
    with ODS_Standard(server=libreoffice_server) as doc:
        doc.addCell("A1", "From Source")
        doc.save(src_file)
        
    with ODS_Standard(server=libreoffice_server) as doc:
        doc.addCell("A1", "Already in Dest")
        doc.save(dest_file)
        
    copy_sheet(src_file, "Hoja1", dest_file, "SheetFromSource", server=libreoffice_server)
    
    from unogenerator import ODS
    with ODS(dest_file, server=libreoffice_server) as doc_check:
        names = doc_check.getSheetNames()
        assert "SheetFromSource" in names
        assert "Hoja1" in names
        doc_check.setActiveSheet("SheetFromSource")
        assert doc_check.getValue("A1") == "From Source"
