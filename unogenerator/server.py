from argparse import ArgumentParser, RawTextHelpFormatter
from gettext import translation
from os import system
from pkg_resources import resource_filename
from unogenerator.commons import __version__, argparse_epilog, addDebugSystem
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
    
    for i in range(args.instances):
        port=args.first_port+i
        backtrace="" if args.backtrace is False else "--backtrace"
        command=f'loffice --accept="socket,host=localhost,port={port};urp;StarOffice.ServiceManager" -env:UserInstallation=file:///tmp/unogenerator{port} --headless {backtrace} > /var/log/unogenerator/unogenerator.{port}.log 2>&1 &'
        print(_(f"  - Launched instance in port {port}"))
        system(command)
#        Popen(command, shell=True)

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
    
