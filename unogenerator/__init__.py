## @package unogenerator.objects
## @brief Objects for unogenerator package
from colorama import init
from datetime import datetime


__version__ = '1.0.0'
__versiondatetime__=datetime(2024, 10, 27, 16, 11)
__versiondate__=__versiondatetime__.date()


from unogenerator.can_import_uno import can_import_uno
init()
if can_import_uno():
    try:
        from .unogenerator import ODT, ODS, ODT_Standard, ODS_Standard, LibreofficeServer
        from .commons import Coord,  Range,  ColorsNamed
    except:
        pass
