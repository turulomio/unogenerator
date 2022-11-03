from argparse import ArgumentParser, RawTextHelpFormatter
from colorama import init as colorama_init
from gettext import translation
from humanize import naturalsize
from multiprocessing import cpu_count
from os import system, makedirs
from importlib.resources import files
from unogenerator.commons import __version__, argparse_epilog, addDebugSystem, get_from_process_info, green, red, magenta
from unogenerator.reusing.casts import list2string
from unogenerator.reusing.listdict_functions import listdict_sum, listdict_average, listdict2list, listdict_order_by
from unogenerator.reusing.percentage import Percentage
from socket import socket, AF_INET, SOCK_STREAM
from subprocess import run
from time import sleep

try:
    t=translation('unogenerator', files("unogenerator") / 'locale')
    _=t.gettext
except:
    _=str

def is_port_opened(host, port):
    sock = socket(AF_INET, SOCK_STREAM)
    result = sock.connect_ex((host,port))
    sock.close()
    if result == 0:
       return False
    else:
       return True


def server_start():#!/bin/bash
    parser=ArgumentParser(
        description=_('Launches LibreOffice server to run unogenerator code'), 
        epilog=argparse_epilog(), 
        formatter_class=RawTextHelpFormatter
    )
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--instances', help=_("Number of concurrent LibreOffice instances"), action="store", type=int,  default=8)
    parser.add_argument('--first_port', help=_("First port assigned to server"), action="store", type=int, default=2002)
    parser.add_argument('--backtrace', help=_("Backtrace logs"), action="store_true", default=False)
    args=parser.parse_args()

    command_start(args.instances, args.first_port, args.backtrace)
    
def command_start(instances, first_port=2002, backtrace=False):
    if not is_port_opened("127.0.0.1", first_port):
        print(red(_("It seems that server is already launched in port {0}").format(first_port)))
        return 

    print(_("Preparing {0} LibreOffice server instances from port {1}:").format(instances, first_port))
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

        print(_("  - Launched LibreOffice headless instance in port {0}").format(port))
        system(command)

def server_stop():
    parser=ArgumentParser(
        description=_('Stops LibreOffice server to finish using unogenerator code'), 
        epilog=argparse_epilog(), 
        formatter_class=RawTextHelpFormatter
    )
    parser.add_argument('--version', action='version', version=__version__)
    parser.add_argument('--debug', help=_("Debug program information"), choices=["DEBUG","INFO","WARNING","ERROR","CRITICAL"], default="ERROR")

    args=parser.parse_args()

    addDebugSystem(args.debug)
    command_stop()
def command_stop():
    command="pkill -c -f ';urp;StarOffice.ServiceManager'"
    s=run(command ,shell=True, capture_output=True)
    print(_("  - Server was stopped killing {0} processes").format(s.stdout.decode('UTF-8').strip()))
    
    
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
    ld=listdict_order_by(ld, "port")
    instances=len(ld)
    max_mem_recommended=recommended_memory*1024*1024
    list_ports=listdict2list(ld, 'port', True)
    cpu_nums=listdict2list(ld, 'cpu_number', True)
    cpu_percentage=Percentage(listdict_average(ld, "cpu_percentage"), 100)
    mem_total=listdict_sum(ld, "mem")
    str_mem_total=green(naturalsize(mem_total)) if mem_total<max_mem_recommended else red(naturalsize(mem_total))
    str_cpu_percentage= green(cpu_percentage) if cpu_percentage.value==0 else red(cpu_percentage)

    print(_("Instances: {0}").format(green(instances)))
    print(_("Ports used: {0}").format(green(list2string(list_ports))))
    print(_("CPU used: {0}").format(green(list2string(cpu_nums))))
    print(_("Memory used: {0}").format(str_mem_total))
    print(_("Max memory recommended: {0}".format(green(naturalsize(max_mem_recommended)))))    
    print(_("CPU percentage: {0}".format(str_cpu_percentage)))
    
    total_cons=0
    for d in ld:
        for con in  d["object"].connections():
            if con.status=="ESTABLISHED":
                total_cons=total_cons+1
    
    str_connections= green(total_cons) if total_cons==0 else red(total_cons)
    print(_("Connections: {0}").format(str_connections))
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
                command_monitor(False, recommended_memory)
            else:
                print(magenta(_("Server wasn't restarted because it's busy")))
                
        else:
            print(magenta(_("Server doesn't need to be restarted")))
            
## @return Boolean. True if server is working
def is_server_working():
    ld=get_from_process_info(attempts=2)
    if ld==[]:
        return False
    return True
