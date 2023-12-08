## @package unogenerator.objects
## @brief Objects for unogenerator package
from colorama import init, Style, Fore
from unogenerator.can_import_uno import can_import_uno, raise_unoexception
from unogenerator import exceptions
init()
if can_import_uno():
    try:
        from .unogenerator import ODT, ODS, ODT_Standard, ODS_Standard
        from .commons import __version__, __versiondate__,  __versiondatetime__,  Coord,  Range,  ColorsNamed
    except(exceptions.UnoException) as e:
        print(Style.BRIGHT + Fore.RED + str(e) + Style.RESET_ALL)

else:
    try:
        raise_unoexception()
    except (exceptions.UnoException) as e:
        print(Style.BRIGHT + Fore.RED + str(e) + Style.RESET_ALL)
