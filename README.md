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

First you have to launch libreoffice as a service. You can use `server.sh` or copy `init.d/unogenerator` to your `/etc/init.d/` directory and run `/etc/init.d/unogenerator start`

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

## Changelog

### 0.2.0 (2021-09-29)
- Added init.d daemon to launch server
- Improving helpers and ods methods

### 0.1.0 (2021-09-18)
- Added basic features
