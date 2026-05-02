from colorama import Style, Fore

def can_import_uno(output=True):
    """
        Check if uno module (from LibreOffice package) can be imported
        
        This should be used by external programs to avoid crashes when uno.py modules is not found
        
        Parameters:
            - output -> bool: Show fail text
    """
    try:
        from uno import Any
        Any    
    except ImportError:
        if output is True:
            e="""This exception is due to you haven't been able to import uno.py module.
        
This module belongs to libreoffice package and you won't find it in systems like pip, poetry ....

To fix it, you must try:
    - Install libreoffice and its python modules
    - Check your python version is the same as the libreoffice compiled module of python
    - If you're using a  python virtual environment you must edit pyvenv.cfg and set:   include-system-site-packages = true
"""
            print(Style.BRIGHT + Fore.RED + str(e) + Style.RESET_ALL)
        return False
    return True
