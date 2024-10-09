from argparse import ArgumentParser, RawTextHelpFormatter
from colorama import init as colorama_init
from gettext import translation
from humanize import naturalsize
from multiprocessing import cpu_count
from pydicts import lod
from importlib.resources import files
from unogenerator import __version__
from unogenerator.commons import argparse_epilog, addDebugSystem, get_from_process_info, green, red
from pydicts.percentage import Percentage

try:
    t=translation('unogenerator', files("unogenerator") / 'locale')
    _=t.gettext
except:
    _=str


    
    
def monitor():    
    colorama_init()
    parser=ArgumentParser(
        description=_('UnoGenerator monitor'), 
        epilog=argparse_epilog(), 
        formatter_class=RawTextHelpFormatter
    )
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--debug', help=_("Debug program information"), choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")
    parser.add_argument('--restart', help=_("Restart server when idle"), action="store_true", default=False)
    parser.add_argument('--recommended', help=_("Recommended max memory (Mb) to restart server when idle. By default: 600Mb per instance"), action="store", type=int, default=cpu_count()*600)
    args=parser.parse_args()

    addDebugSystem(args.debug)
    command_monitor(args.restart, args.recommended)
    
## @param restart boolean. To restart unogenerator server when idle and used memory above recommended memory
## @param recommended. Integer. Recomended memory in Mb
def command_monitor(restart, recommended_memory):
    ld=get_from_process_info(cpu_percentage=True)
    ld=lod.lod_order_by(ld, "port")
    instances=len(ld)
    max_mem_recommended=recommended_memory*1024*1024
    list_ports=lod.lod2list(ld, 'port', True)
    cpu_nums=lod.lod2list(ld, 'cpu_number', True)
    cpu_percentage=Percentage(lod.lod_average(ld, "cpu_percentage"), 100)
    mem_total=lod.lod_sum(ld, "mem")
    str_mem_total=green(naturalsize(mem_total)) if mem_total<max_mem_recommended else red(naturalsize(mem_total))
    str_cpu_percentage= green(cpu_percentage) if cpu_percentage.value==0 else red(cpu_percentage)

    print(_("Instances: {0}").format(green(instances)))
    print(_("Ports used: {0}").format(green(str(list_ports)[1:-1])))
    print(_("CPU used: {0}").format(green(str(cpu_nums)[ 1:-1])))
    print(_("Memory used: {0}").format(str_mem_total))
    print(_("Max memory recommended: {0}".format(green(naturalsize(max_mem_recommended)))))    
    print(_("CPU percentage: {0}".format(str_cpu_percentage)))
    
    total_cons=0
    for d in ld:
        for con in  d["object"].net_connections():
            if con.status=="ESTABLISHED":
                total_cons=total_cons+1
    
    str_connections= green(total_cons) if total_cons==0 else red(total_cons)
    print(_("Connections: {0}").format(str_connections))
