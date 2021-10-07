# UnoGenerator problems

In some installations are some problems with concurrency

These are the use in gentoo where unogenerator works fine:

[ebuild   R    ] app-office/libreoffice-7.1.6.2::gentoo  USE="accessibility bluetooth branding cups dbus gstreamer gtk java kde mariadb odk pdfimport postgres vulkan -base -clang -coinmp -custom-cflags -debug -eds -firebird -googledrive -ldap -test" LIBREOFFICE_EXTENSIONS="-nlpsolver -scripting-beanshell -scripting-javascript -wiki-publisher" PYTHON_SINGLE_TARGET="python3_9 -python3_8 -python3_10" 0 KiB

## Debug
/opt/libreoffice4.1/program/soffice --accept=“socket,host=127.0.0.1,port=8100;urp;” --headless --norestore --invisible --nologo --nofirststartwizard --language=pt-BR >>/var/log/libreoffice 2>&1 &

Lately it simply stops working after some time(hours, days, not certain), queuing the new files, returning no error. I dunno what’s happening 'cause i can see no option to log some exception or error. There’s any way to enable some kind of log or debug of this software? I’m kinda lost :s

Use the --backtrace argument. To obtain even greater detail on what is occurring I would suggest downloading a daily build with debugging symbols.

## Couldn't connect to socket

Traceback (most recent call last):
File "/usr/lib/python3.9/concurrent/futures/process.py", line 243, in _process_worker
r = call_item.fn(*call_item.args, **call_item.kwargs)
File "/usr/lib/python3.9/site-packages/unogenerator-0.6.0-py3.9.egg/unogenerator/demo.py", line 238, in demo_odt_standard
doc=ODT_Standard(port)
File "/usr/lib/python3.9/site-packages/unogenerator-0.6.0-py3.9.egg/unogenerator/unogenerator.py", line 621, in init
ODT.init(self, resource_filename(name, 'templates/standard.odt'), loserver_port)
File "/usr/lib/python3.9/site-packages/unogenerator-0.6.0-py3.9.egg/unogenerator/unogenerator.py", line 88, in init
ODF.init(self, template, loserver_port)
File "/usr/lib/python3.9/site-packages/unogenerator-0.6.0-py3.9.egg/unogenerator/unogenerator.py", line 49, in init
ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={loserver_port};urp;StarOffice.ComponentContext')
unogenerator.unogenerator.com.sun.star.connection.NoConnectException: Connector : couldn't connect to socket (Connection refused) /var/tmp/portage/app-office/libreoffice-7.1.6.2/work/libreoffice-7.1.6.2/io/source/connector/connector.cxx:119