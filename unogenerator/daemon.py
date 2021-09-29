
from subprocess import run

def daemon_start(port=2002):#!/bin/bash
    command=f"""loffice --accept="socket,host=localhost,port={port};urp;StarOffice.ServiceManager" --headless"""

    p=run(command ,shell=True, capture_output=True)
    print(p, dir(p))
    return p

def daemon_stop(port=2002):
    command=f"""pkill -f '{port};urp;StarOffice.ServiceManager'"""
    p=run(command ,shell=True, capture_output=True)
    pass