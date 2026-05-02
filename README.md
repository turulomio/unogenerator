# UnoGenerator [![PyPI - Downloads](https://img.shields.io/pypi/dm/unogenerator?label=Pypi%20downloads)](https://pypi.org/project/unogenerator/) [![Github - Downloads](https://shields.io/github/downloads/turulomio/unogenerator/total?label=Github%20downloads )](https://github.com/turulomio/unogenerator/) [![Tests](https://github.com/turulomio/unogenerator/actions/workflows/python-app.yml/badge.svg)](https://github.com/turulomio/unogenerator/actions/workflows/python-app.yml)
<div align="center">
  <div style="display: flex; align-items: flex-start;">
    <img src="https://raw.githubusercontent.com/turulomio/unogenerator/main/unogenerator/images/unogenerator.png" width="150" title="Unogenerator logo">
  </div>
</div>

## Description

Python module to generate Libreoffice documents (ODT and ODS) programatically.

Morever, you can export them to (.xlsx, .docx, .pdf) easyly.

It uses Libreoffice uno module, so you need Libreoffice to be installed in your system.

## Installation

Only Linux is supported. I'm going to write unogenerator [installation methods](INSTALL.md) for some main Linux Distributions 


## ODT 'Hello World' example

This is a Hello World example. You get the example in odt, docx and pdf formats:

```python
from unogenerator import ODT_Standard
with ODT_Standard() as doc:
  doc.addParagraph("Hello World", "Heading 1")
  doc.addParagraph("Easy, isn't it","Standard")
  doc.save("hello_world.odt")
  doc.export_docx("hello_world.docx")
  doc.export_pdf("hello_world.pdf")
```

## ODS 'Hello World' example

This is a Hello World example. You'll get example files in ods, xlsx and pdf formats:

```python
from unogenerator import ODS_Standard
with ODS_Standard() as doc:
  doc.addCellMergedWithStyle("A1:E1", "Hello world", style="BoldCenter")
  doc.save("hello_world.ods")
  doc.export_xlsx("hello_world.xlsx")
  doc.export_pdf("hello_world.pdf")
```

## Unogenerator scripts

Python unogenerator package has the following scripts:

### unogenerator_monitor

Monitors your libreoffice server instances

### unogenerator_translation

With this tool you can translate several odt files with one command. It generates .pot and .po files, where you can set your translations. Then run your command again and you'll get your files translated

`unogenerator_translation --from_language es --to_language en --input original.odt --input original2.odt  --output_directory "translation_original"`

You can use --fake to see simulation of your translation

### unogenerator_demo

With this tool you can generate a demo, remove its result files and make benchmark comparations in your system

## Documentation
You can read [documentation](https://github.com/turulomio/unogenerator/blob/main/doc/unogenerator_documentation_en.odt?raw=true) in doc directory. It has been created with unogenerator.

## Development links

- [LibreOffice code](https://github.com/LibreOffice/core)
- [LibreOffice API](https://api.libreoffice.org/docs/idl/ref/index.html)
- [OpenOffice Forums](https://forum.openoffice.org/en/forum/viewforum.php?f=20)
- [LibreOffice Forums](https://ask.libreoffice.org/)
- [UnoGenerator API](https://coolnewton.mooo.com/doxygen/unogenerator/)
