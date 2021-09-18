# UnoGenerator
## Description
Python module to generate Libreoffice documents programatically

It uses uno module, so you need Libreoffice to be installed in your system

## Installation
You can use pip to install this python package:

`pip install unogenerator`

## Hello World example

This is a Hello World example. You get the example in odt, docx and pdf formats:

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
You can read [documentation](https://github.com/turulomio/unogenerator/blob/main/doc/unogenerator_documentation_en.odt) in doc directory. It has been created with unogenerator.

## Changelog

### 0.1.0 (2021-09-18)
- Added basic features
