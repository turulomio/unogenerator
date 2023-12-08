from unogenerator import exceptions

def can_import_uno():
    """
        Check if uno module (from LibreOffice package) can be imported
        
        This should be used by external programs to avoid crashes when uno.py modules is not found
    """
    try:
        from uno import Any
        Any    
    except ImportError:
        return False
    return True
    
def raise_unoexception():
        raise exceptions.UnoException(  """This exception is due to you haven't bee able to import uno.py module.
        
This module belongs to libreoffice package and you won't find it in pip,  poetry ....
        """)
