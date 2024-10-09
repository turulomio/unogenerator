from argparse import ArgumentParser, RawTextHelpFormatter
from colorama import init as colorama_init
from copy import deepcopy
from datetime import datetime, timedelta
from gettext import translation
from pydicts import lod
from importlib.resources import files
from humanize import naturalsize
from psutil import process_iter
from unogenerator import __version__
from unogenerator.commons import argparse_epilog, addDebugSystem, green, red
from os import system, scandir
from pydicts.percentage import Percentage
from time import sleep
from sys import stdout

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
    parser.add_argument('--seconds', help=_('Seconds to color Last CPU variation'), action="store",  type=int,  default=60)
    
    args=parser.parse_args()

    addDebugSystem(args.debug)
    command_monitor(args.seconds)
    
def clear_screen():
    stdout.write("\033[H\033[J")  # Move cursor to top-left and clear screen
    stdout.flush()
   
## @param restart boolean. To restart unogenerator server when idle and used memory above recommended memory
## @param recommended. Integer. Recomended memory in Mb
def command_monitor(seconds):
    last_dod_=None
    while True:
        clear_screen()
        directorios_deben=set()
        dod_={}
        for p in process_iter(['name','cmdline', 'pid']): 
            d={}
            try:
                if p.info['name']=='soffice.bin':
                    if  p.info['cmdline'] is not None and 'file:///tmp/unogenerator'  in ' '.join(p.info['cmdline']):
                        d["pid"]=p.pid
                        d["port"]=int(p.info['cmdline'][1].split("unogenerator")[1])
                        directorios_deben.add(f"unogenerator{d['port']}")
                        d["mem"]=naturalsize(p.memory_info().rss)
                        d["status"]=p.status()
                        d["duration"]=datetime.now()-datetime.fromtimestamp(p.create_time())
                        d["cpu_percentage"]=Percentage(p.cpu_percent(interval=0.01), 100)
                        connections=len(p.connections())
                        d["conexiones"]=connections
                        
                        if last_dod_ is None: #Primera ronda
                            d["last_cpu_percentage_datetime"]=datetime.now()
                            d["Last CPU"]=timedelta(seconds=0)
                        else: #Siguientes rondas
                            #Actualiza cpu_percentage
                            if d["cpu_percentage"].value>0:
                                d["last_cpu_percentage_datetime"]=datetime.now()
                                d["Last CPU"]=timedelta(seconds=0)
                            else: #No se ha utilizado cpu
                                if p.pid in last_dod_ and "last_cpu_percentage_datetime" in last_dod_[p.pid]:
                                    d["last_cpu_percentage_datetime"]=last_dod_[p.pid]["last_cpu_percentage_datetime"]
                                    d["Last CPU"]=datetime.now()-last_dod_[p.pid]["last_cpu_percentage_datetime"]
                                else:
                                    d["last_cpu_percentage_datetime"]=datetime.now()
                                    d["Last CPU"]=timedelta(seconds=0)

                        dod_[p.pid]=d
            except:
                pass
        last_dod_=deepcopy(dod_)
        
        result_dod=deepcopy(dod_)
        result_lod=lod.lod_order_by(result_dod.values(), "duration",  reverse=True)
        lod.lod_remove_key(result_lod,"last_cpu_percentage_datetime")
        
        #Change color on Last CPU
        for d in result_lod:
            d["duration"]=red(d["duration"] )if d["duration"].total_seconds()>seconds else green(d["duration"])
            d["Last CPU"]=red(d["Last CPU"] )if d["Last CPU"].total_seconds()>seconds else green(d["Last CPU"])
        
        lod.lod_print(result_lod)
        
        print("Número de procesos activos:",  len(result_lod))
        
        print("Directorios que no tienen proceso")
        directorios_existen=set()
        for entry in scandir("/tmp"):
            if entry.is_dir() and entry.name.startswith("unogenerator"):
                directorios_existen.add(entry.name)
        print(directorios_existen-directorios_deben)
       
        sleep(1)    
    
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
