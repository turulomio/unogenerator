from argparse import ArgumentParser, RawTextHelpFormatter
from colorama import init as colorama_init
from gettext import translation
from humanize import naturalsize
from os import system, makedirs
from pkg_resources import resource_filename
from unogenerator.commons import __version__, argparse_epilog, addDebugSystem, get_from_process_info, green, red
from unogenerator.reusing.casts import list2string
from unogenerator.reusing.listdict_functions import listdict_sum, listdict_average, listdict2list
from unogenerator.reusing.percentage import Percentage
from subprocess import run

try:
    t=translation('unogenerator', resource_filename("unogenerator","locale"))
    _=t.gettext
except:
    _=str

def server_start():#!/bin/bash
    parser=ArgumentParser(
        description=_('Launches libreoffice server to run unogenerator code'), 
        epilog=argparse_epilog(), 
        formatter_class=RawTextHelpFormatter
    )
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--instances', help="Number of concurrent libreoffice instances", action="store", type=int,  default=8)
    parser.add_argument('--first_port', help="First port assigned to server", action="store", type=int, default=2002)
    parser.add_argument('--backtrace', help="Backtrace logs", action="store_true", default=False)
    args=parser.parse_args()

    print(_(f"Preparing {args.instances} libreoffice server instances from port {args.first_port}:"))
    
    if args.backtrace is True:
        makedirs("/var/log/unogenerator/", exist_ok=True)
        print(_("You can see instances output in /var/log/unogenerator"))
        backtrace="--backtrace"
    else:
        backtrace=""
    
    for i in range(args.instances):
        port=args.first_port+i
        if args.backtrace is True:
            command=f'loffice --accept="socket,host=localhost,port={port};urp;StarOffice.ServiceManager" -env:UserInstallation=file:///tmp/unogenerator{port} --headless {backtrace} > /var/log/unogenerator/unogenerator.{port}.log 2>&1 &'
        else:
            command=f'loffice --accept="socket,host=localhost,port={port};urp;StarOffice.ServiceManager" -env:UserInstallation=file:///tmp/unogenerator{port} --headless &'

        print(_(f"  - Launched instance in port {port}"))
        system(command)

def server_stop():
    parser=ArgumentParser(
        description=_('Stops libreoffice server to finish using unogenerator code'), 
        epilog=argparse_epilog(), 
        formatter_class=RawTextHelpFormatter
    )
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--debug', help="Debug program information", choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")
    args=parser.parse_args()

    addDebugSystem(args.debug)
    command="pkill -c -f ';urp;StarOffice.ServiceManager'"
    s=run(command ,shell=True, capture_output=True)
    print(_(f"  - Server was stopped killing {s.stdout.decode('UTF-8').strip()} processes"))
    
    
def monitor():    
    colorama_init()
    parser=ArgumentParser(
        description=_('Monitor unogenerator statistics'), 
        epilog=argparse_epilog(), 
        formatter_class=RawTextHelpFormatter
    )
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--debug', help="Debug program information", choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")

    args=parser.parse_args()

    addDebugSystem(args.debug)
    
    ld=get_from_process_info(cpu_percentage=True)
    instances=len(ld)
    list_ports=listdict2list(ld, 'port', True)
    cpu_percentage=Percentage(listdict_average(ld, "cpu_percentage"), 100)
    mem_total=listdict_sum(ld, "mem")
    str_mem_total=green(naturalsize(mem_total)) if mem_total>(instances*440010752)*1.50 else red(naturalsize(mem_total))
    str_cpu_percentage= green(cpu_percentage) if cpu_percentage==0 else red(cpu_percentage)

    print(_(f"Instances: {green(instances)}"))
    print(_(f"Ports used: {green(list2string(list_ports))}"))
    print(_(f"Memory used: {str_mem_total}"))    
    print(_(f"CPU percentage: {str_cpu_percentage}"))
    
    total_cons=0
    for d in ld:
        for con in  d["object"].connections():
#            print(con.laddr,  con.raddr, con.status)
            if con.status=="ESTABLISHED":
                total_cons=total_cons+1
    
    str_connections= green(total_cons) if total_cons==0 else red(total_cons)
    print(_(f"Connections: {str_connections}"))
