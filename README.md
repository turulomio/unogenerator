# UnoGenerator [![PyPI - Downloads](https://img.shields.io/pypi/dm/unogenerator?label=Pypi%20downloads)](https://pypi.org/project/unogenerator/) [![Github - Downloads](https://shields.io/github/downloads/turulomio/unogenerator/total?label=Github%20downloads )](https://github.com/turulomio/unogenerator/) [![Tests](https://github.com/turulomio/unogenerator/actions/workflows/python-app.yml/badge.svg)](https://github.com/turulomio/unogenerator/actions/workflows/python-app.yml)
<div align="center">
  <div style="display: flex; align-items: flex-start;">
    <img src="https://raw.githubusercontent.com/turulomio/unogenerator/main/unogenerator/images/unogenerator.png" width="150" title="Unogenerator logo">
  </div>
</div>

## Description

**UnoGenerator** is a powerful Python module designed to generate LibreOffice documents (ODT and ODS) programmatically with high performance and professional aesthetics.

Key features include:
- **Professional Defaults**: Automatic text wrapping (`word_wrap=True`) and vertical centering for professional-looking reports out of the box.
- **High Performance**: Optimized for large datasets. Inserting 10,000 rows with default formatting takes less than 0.3 seconds.
- **Rich Exports**: Easy export to `.xlsx`, `.docx`, and `.pdf`.
- **Advanced Helpers**: Flexible totals generation, automatic column width calculation, and complex data block handling.
- **Multilingual Support**: Built-in support for translations and localized document generation.

It uses the LibreOffice UNO module, requiring a LibreOffice installation on your system.

## Installation

Only Linux is supported. See the [installation guide](INSTALL.md) for detailed instructions on main Linux distributions.

## Architecture

UnoGenerator follows a clean, template-based architecture:
- **ODS/ODT**: Generic base classes for document manipulation.
- **ODS_Standard / ODT_Standard**: Optimized subclasses that use the built-in professional templates, providing specific styles (like "Normal", "BoldCenter") and optimized row heights.

## ODT 'Hello World' example

Create a professional ODT document with just a few lines:

```python
from unogenerator import ODT_Standard

with ODT_Standard() as doc:
    doc.addParagraph("Hello World", "Heading 1")
    doc.addParagraph("Easy, isn't it", "Standard")
    doc.save("hello_world.odt")
    doc.export_docx("hello_world.docx")
    doc.export_pdf("hello_world.pdf")
```

## ODS 'Hello World' example

Generate a styled spreadsheet with automatic wrapping and alignment:

```python
from unogenerator import ODS_Standard

with ODS_Standard() as doc:
    # word_wrap and vertical alignment are enabled by default
    doc.addCellMergedWithStyle("A1:E1", "Sales Report 2026", style="BoldCenter")
    doc.addRowWithStyle("A2", ["Product", "Quantity", "Price", "Total"])
    doc.save("sales_report.ods")
    doc.export_xlsx("sales_report.xlsx")
```

## Advanced Features: Totals and Formulas

UnoGenerator provides advanced helpers to generate calculations quickly:

```python
from unogenerator import helpers

# Generates both row and column totals with a custom formula template
helpers.cross_totals_from_range(doc, "B2:D10", key="=SUM({}*1.21)")
```

## Unogenerator scripts

The package includes several useful CLI tools:

### unogenerator_monitor
Monitors your LibreOffice server instances and their status.

### unogenerator_translation
Translate multiple ODT files using standard `.pot` and `.po` files.
`unogenerator_translation --from_language es --to_language en --input original.odt --output_directory "translated"`

### unogenerator_demo
Generate comprehensive example files and perform performance benchmarks on your system.

## Documentation
Full technical [documentation](https://github.com/turulomio/unogenerator/blob/main/doc/unogenerator_documentation_en.odt?raw=true) is available in the `doc` directory, created using UnoGenerator itself.

## Development links

- [LibreOffice code](https://github.com/LibreOffice/core)
- [LibreOffice API](https://api.libreoffice.org/docs/idl/ref/index.html)
- [UnoGenerator API (Doxygen)](https://coolnewton.mooo.com/doxygen/unogenerator/)
