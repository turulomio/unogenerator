# UnoGenerator
## Description
Python module to generate Libreoffice documents (ODT and ODS) programatically.

Morever, you can export them to (.xlsx, .docx, .pdf) easyly.

It uses Libreoffice uno module, so you need Libreoffice to be installed in your system.

## Installation
You can use pip to install this python package:

`pip install unogenerator`

## Hello World example

This is a Hello World example. You get the example in odt, docx and pdf formats:

First you have to launch libreoffice as a service. You can use `unogenerator_start` or copy `init.d/unogenerator` to your `/etc/init.d/` directory and run `/etc/init.d/unogenerator start`

```python
from unogenerator import ODT_Standard
doc=ODT_Standard()
doc.addParagraph("Hello World", "Heading 1")
doc.addParagraph("Easy, isn't it","Standard")
doc.save("hello_world.odt")
doc.export_docx("hello_world.docx")
doc.export_pdf("hello_world.pdf")
doc.close()
```

## Documentation
You can read [documentation](https://github.com/turulomio/unogenerator/blob/main/doc/unogenerator_documentation_en.odt?raw=true) in doc directory. It has been created with unogenerator.

## Development links

- [Libreoffice code](https://github.com/LibreOffice/core)
- [Libreoffice API](https://api.libreoffice.org/docs/idl/ref/index.html)


## Changelog

### 0.5.0 (2021-10-3)
- Added dispatcher and create service methods.
- Added posibility to link images to document.
- Image width and height are now in cm units.
- Added method addCellNames() for ranged names.
- Improving freezeAndSelect with None values.
- Added range column and number copy additions.

### 0.4.0 (2021-10-2)
- Replaced server.sh by `unogenerator_start` and `unogenerator_stop`
- Added helper_list_of_dict and helper_list_of_ordereddict methods
- Changed vertical alignment to center in all personal styles
- Removed Gentoo confd initd files. Moved to unogenerator_daemon in myportage

### 0.3.0 (2021-10-1)
- server.sh now launches n instances
- Code works with concurrency. Added method to change to next_port
- Added .conf and init.d for Gentoo
- Date and times formats are showed correctly in Calc
- Helpers added

### 0.2.0 (2021-09-29)
- Added init.d daemon to launch server
- Improving helpers and ods methods

### 0.1.0 (2021-09-18)
- Added basic features
