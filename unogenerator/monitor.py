from argparse import ArgumentParser, RawTextHelpFormatter
from colorama import init as colorama_init
from datetime import datetime
from gettext import translation
from pydicts import lod
from importlib.resources import files
from humanize import naturalsize
from psutil import process_iter
from unogenerator import __version__
from unogenerator.commons import argparse_epilog, addDebugSystem, green, red
from os import system, path, scandir
from pydicts.percentage import Percentage
from time import sleep

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
    args=parser.parse_args()

    addDebugSystem(args.debug)
    command_monitor()
    
   
## @param restart boolean. To restart unogenerator server when idle and used memory above recommended memory
## @param recommended. Integer. Recomended memory in Mb
def command_monitor():
    while True:
        system("clear")
        
        directorios_deben=set()
        
        r=[]
        for p in process_iter(['name','cmdline', 'pid']): 
            d={}
            try:
                if p.info['name']=='soffice.bin':
                    if  p.info['cmdline'] is not None and 'file:///tmp/unogenerator'  in ' '.join(p.info['cmdline']):
                        d["pid"]=p.pid
                        d["port"]=int(p.info['cmdline'][1].split("unogenerator")[1])
                        directorios_deben.add(f"unogenerator{d['port']}")
                        d["mem"]=naturalsize(p.memory_info().rss)
                        d["cpu_percentage"]=Percentage(p.cpu_percent(interval=0.01), 100)
                        d["status"]=p.status()
                        d["duration"]=datetime.now()-datetime.fromtimestamp(p.create_time())
                        d["conexiones"]=len(p.connections())
                        r.append(d)
            except:
                pass
        
        r=lod.lod_order_by(r, "port")
        lod.lod_print(r)
        
        print("Número de procesos activos:",  len(r))
        
        print("Directorios que no tienen proceso")
        directorios_existen=set()
        for entry in scandir("/tmp"):
            if entry.is_dir() and entry.name.startswith("unogenerator"):
                directorios_existen.add(entry.name)
        print(directorios_existen-directorios_deben)
       
        sleep(1)
    #    instances=len(ld)
#    list_ports=lod.lod2list(ld, 'port', True)
#    cpu_nums=lod.lod2list(ld, 'cpu_number', True)
#    cpu_percentage=Percentage(lod.lod_average(ld, "cpu_percentage"), 100)
#    mem_total=lod.lod_sum(ld, "mem")
#    str_mem_total=green(naturalsize(mem_total)) if mem_total<max_mem_recommended else red(naturalsize(mem_total))
#    str_cpu_percentage= green(cpu_percentage) if cpu_percentage.value==0 else red(cpu_percentage)
#
#    print(_("Instances: {0}").format(green(instances)))
#    print(_("Ports used: {0}").format(green(str(list_ports)[1:-1])))
#    print(_("CPU used: {0}").format(green(str(cpu_nums)[ 1:-1])))
#    print(_("Memory used: {0}").format(str_mem_total))
#    print(_("Max memory recommended: {0}".format(green(naturalsize(max_mem_recommended)))))    
#    print(_("CPU percentage: {0}".format(str_cpu_percentage)))
#    
#    total_cons=0
#    for d in ld:
#        for con in  d["object"].net_connections():
#            if con.status=="ESTABLISHED":
#                total_cons=total_cons+1
#    
#    str_connections= green(total_cons) if total_cons==0 else red(total_cons)
#    print(_("Connections: {0}").format(str_connections))

    
    
def cleaner():    
    colorama_init()
    parser=ArgumentParser(
        description=_('UnoGenerator cleaner'), 
        epilog=argparse_epilog(), 
        formatter_class=RawTextHelpFormatter
    )
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--debug', help=_("Debug program information"), choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")
    args=parser.parse_args()

    addDebugSystem(args.debug)
    command_cleaner()
    
## @param restart boolean. To restart unogenerator server when idle and used memory above recommended memory
## @param recommended. Integer. Recomended memory in Mb
def command_cleaner():
    system("killall -9 soffice.bin")
    system("rm -Rf /tmp/unogenerator*")
