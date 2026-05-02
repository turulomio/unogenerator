# UnoGenerator installation methods

Only Linux is supported

You need LibreOffice installed in your system. Imagemagick is recommended for some image operations.

In some distros you need to install Python LibreOffice bindings too (python3-uno)

If you are using a python virtual environmet change `include-system-site-packages = false` in your pyvenv.conf file to allow to access libreoffice python bindings from your OS

Here you can find installation methods for some Linux distributions

## Gentoo Linux

I've developed this tool with Gentoo. 

If you want to install it in Gentoo you can use my ebuild at [myportage](https://github.com/turulomio/myportage/tree/master/dev-python/unogenerator).

## Debian (LiveCD)

I've downloaded debian-live-11.2.0-amd64-kde.iso. Run from a usb pendrive and opened Konsole

`sudo su -`

`apt update`

`apt install python3-pip`

`pip install unogenerator`

`unogenerator_demo --create`

If everything goes fine, you'll find demo files created in your current directory

## Kubuntu (LiveCD)

To instal python3 pip system:

`apt install python3-pip`

To install unogenerator:

`pip install unogenerator`

To generate unogenerator demo files:

`unogenerator_demo --create`

If everything goes fine, you'll find demo files created in your current directory

