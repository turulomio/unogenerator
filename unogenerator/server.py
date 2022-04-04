from argparse import ArgumentParser, RawTextHelpFormatter
from colorama import init as colorama_init
from gettext import translation
from humanize import naturalsize
from os import system, makedirs
from pkg_resources import resource_filename
from unogenerator.commons import __version__, argparse_epilog, addDebugSystem, get_from_process_info, green, red, magenta
from unogenerator.reusing.casts import list2string
from unogenerator.reusing.listdict_functions import listdict_sum, listdict_average, listdict2list, listdict_order_by
from unogenerator.reusing.percentage import Percentage
from subprocess import run
from time import sleep

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

    command_start(args.instances, args.first_port, args.backtrace)
    
def command_start(instances, first_port=2002, backtrace=False):
    print(_(f"Preparing {instances} libreoffice server instances from port {first_port}:"))
    if backtrace is True:
        makedirs("/var/log/unogenerator/", exist_ok=True)
        print(_("You can see instances output in /var/log/unogenerator"))
        backtrace="--backtrace"
    else:
        backtrace=""
    
    for i in range(instances):
        port=first_port+i
        if backtrace is True:
            command=f'loffice --accept="socket,host=localhost,port={port};urp;StarOffice.ServiceManager" -env:UserInstallation=file:///tmp/unogenerator{port} --headless {backtrace} > /var/log/unogenerator/unogenerator.{port}.log 2>&1 &'
        else:
            command=f'loffice --accept="socket,host=localhost,port={port};urp;StarOffice.ServiceManager" -env:UserInstallation=file:///tmp/unogenerator{port} --headless &'
        sleep(0.4)

        print(_(f"  - Launched LibreOffice headless instance in port {port}"))
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
    command_stop()
def command_stop():
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
    parser.add_argument('--restart', help="Restart server when idle", action="store_true", default=False)
    parser.add_argument('--max_mem_multiplier', help="Restart server when idle", action="store", type=int, default=2)
    args=parser.parse_args()

    addDebugSystem(args.debug)
    command_monitor(args.restart, args.max_mem_multiplier)
    
def command_monitor(restart, max_mem_multiplier):
    
    ld=get_from_process_info(cpu_percentage=True)
    ld=listdict_order_by(ld, "port")
    instances=len(ld)
    max_mem_recommended=instances*440010752*max_mem_multiplier
    list_ports=listdict2list(ld, 'port', True)
    cpu_nums=listdict2list(ld, 'cpu_number', True)
    cpu_percentage=Percentage(listdict_average(ld, "cpu_percentage"), 100)
    mem_total=listdict_sum(ld, "mem")
    str_mem_total=green(naturalsize(mem_total)) if mem_total<max_mem_recommended else red(naturalsize(mem_total))
    str_cpu_percentage= green(cpu_percentage) if cpu_percentage.value==0 else red(cpu_percentage)

    print(_(f"Instances: {green(instances)}"))
    print(_(f"Ports used: {green(list2string(list_ports))}"))
    print(_(f"CPU used: {green(list2string(cpu_nums))}"))
    print(_(f"Memory used: {str_mem_total}"))
    print(_(f"Max memory recommended: {green(naturalsize(max_mem_recommended))}"))    
    print(_(f"CPU percentage: {str_cpu_percentage}"))
    
    total_cons=0
    for d in ld:
        for con in  d["object"].connections():
#            print(con.laddr,  con.raddr, con.status)
            if con.status=="ESTABLISHED":
                total_cons=total_cons+1
    
    str_connections= green(total_cons) if total_cons==0 else red(total_cons)
    print(_(f"Connections: {str_connections}"))
    if restart is True:
        if mem_total>max_mem_recommended:
            
            if total_cons==0 :
                print(magenta(_("Restarting server")))
                first_port=min(list_ports)
                command_stop()
                sleep(1)
                command_start(instances, first_port, False)
                sleep(10)
                print(magenta(_("Server restarted")))
                command_monitor(False, max_mem_multiplier)
            else:
                print(magenta(_("Server wasn't restarted because it's busy")))
                
        else:
            print(magenta(_("Server doesn't need to be restarted")))
