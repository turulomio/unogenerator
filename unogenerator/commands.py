from unogenerator import ODS, LibreofficeServer, exceptions
from com.sun.star.sheet.SheetLinkMode import NORMAL, NONE
import os
from uno import systemPathToFileUrl
import logging

logger = logging.getLogger(__name__)

def copy_sheet(src_path: str, src_sheet_name: str, dest_path: str, dest_sheet_name: str | None = None, server: LibreofficeServer | None = None, index: int | None = None) -> None:
    """
    Copies a sheet from one ODS file to another, including content and formatting.

    Args:
        src_path (str): Path to the source ODS file.
        src_sheet_name (str): Name of the sheet to copy from the source file.
        dest_path (str): Path to the destination ODS file.
        dest_sheet_name (str | None): Name for the copied sheet in the destination file.
                                     Defaults to src_sheet_name.
        server (LibreofficeServer | None): Optional existing LibreOffice server instance.
        index (int | None): Optional position index to insert the sheet. If None, it is appended at the end.
    
    Raises:
        UnogeneratorException: If the source file doesn't exist or destination sheet already exists.
    """
    if not os.path.exists(src_path):
        raise exceptions.UnogeneratorException(f"Source file '{src_path}' does not exist.")
        
    if dest_sheet_name is None:
        dest_sheet_name = src_sheet_name
        
    src_url = systemPathToFileUrl(os.path.abspath(src_path))
    
    # We open or create the destination file
    with ODS(dest_path if os.path.exists(dest_path) else None, server=server) as doc_dest:
        sheets = doc_dest.document.getSheets()
        
        # Check if sheet already exists
        if sheets.hasByName(dest_sheet_name):
            raise exceptions.UnogeneratorException(f"Sheet '{dest_sheet_name}' already exists in '{dest_path}'")
             
        if index is None:
            index = sheets.Count

        sheets.insertNewByName(dest_sheet_name, index)
        sheet = sheets.getByName(dest_sheet_name)
        
        # Perform the copy via linking
        try:
            sheet.link(src_url, src_sheet_name, "", "", NORMAL)
            sheet.link(src_url, src_sheet_name, "", "", NONE)
            logger.debug(f"Successfully copied sheet '{src_sheet_name}' from '{src_path}' to '{dest_path}' as '{dest_sheet_name}'")
        except Exception as e:
            raise exceptions.UnogeneratorException(f"Failed to copy sheet: {e}")
        
        doc_dest.save(dest_path, overwrite_template=True)
