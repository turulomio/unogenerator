#!/bin/bash
INSTANCES=${1:-8}   # Defaults to /tmp dir.
FIRST_PORT=${2:-2002}           # Default value is 1.


pkill -f ";urp;StarOffice.ServiceManager"

echo "Preparing $INSTANCES libreoffice server instances from port $FIRST_PORT:"

for i in  $(seq 1 $INSTANCES)
do
    PORT=$((FIRST_PORT+i-1))
    echo "  - Launched instance in port $PORT"
    /usr/bin/loffice --accept="socket,host=localhost,port=$PORT;urp;StarOffice.ServiceManager"  -env:UserInstallation=file:///tmp/l$PORT --headless &
done

