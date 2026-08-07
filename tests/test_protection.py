from unogenerator import ODS_Standard

def test_sheet_lock_unlock(libreoffice_server):
    with ODS_Standard(server=libreoffice_server) as doc:
        assert not doc.isSheetProtected()

        # Protect / Lock sheet with password
        doc.protectSheet("pass123")
        assert doc.isSheetProtected()

        # Unprotect / Unlock sheet with password
        doc.unprotectSheet("pass123")
        assert not doc.isSheetProtected()

        # Using lockSheet / unlockSheet aliases
        doc.lockSheet()
        assert doc.isSheetProtected()
        doc.unlockSheet()
        assert not doc.isSheetProtected()

def test_cell_lock_unlock(libreoffice_server):
    with ODS_Standard(server=libreoffice_server) as doc:
        doc.addCell("A1", "Cell A1")
        doc.addCell("B1", "Cell B1")

        # Initial state in template is locked
        assert doc.isCellLocked("A1")
        assert doc.isCellLocked("B1")

        # Unlock cell B1
        doc.unlockCell("B1")
        assert not doc.isCellLocked("B1")
        assert doc.isCellLocked("A1")

        # Lock cell B1 again
        doc.lockCell("B1")
        assert doc.isCellLocked("B1")

        # Using protectCell / unprotectCell aliases
        doc.unprotectCell("A1")
        assert not doc.isCellLocked("A1")
        doc.protectCell("A1")
        assert doc.isCellLocked("A1")

def test_block_range_lock_unlock(libreoffice_server):
    with ODS_Standard(server=libreoffice_server) as doc:
        doc.addCell("A1", "Val 1")
        doc.addCell("B1", "Val 2")
        doc.addCell("A2", "Val 3")
        doc.addCell("B2", "Val 4")

        # Unlock range / block A1:B2
        doc.unlockRange("A1:B2")
        assert not doc.isRangeLocked("A1:B2")
        assert not doc.isBlockLocked("A1:B2")
        assert not doc.isCellLocked("A1")
        assert not doc.isCellLocked("B2")

        # Lock block / range A1:B2
        doc.lockBlock("A1:B2")
        assert doc.isRangeLocked("A1:B2")
        assert doc.isBlockLocked("A1:B2")
        assert doc.isCellLocked("A1")
        assert doc.isCellLocked("B2")

        # Using unlockBlock / lockRange
        doc.unlockBlock("A1:B2")
        assert not doc.isBlockLocked("A1:B2")
        doc.lockRange("A1:B2")
        assert doc.isBlockLocked("A1:B2")
