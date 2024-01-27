# UnoGenerator [![PyPI - Downloads](https://img.shields.io/pypi/dm/unogenerator?label=Pypi%20downloads)](https://pypi.org/project/unogenerator/) [![Github - Downloads](https://shields.io/github/downloads/turulomio/unogenerator/total?label=Github%20downloads )](https://github.com/turulomio/unogenerator/)
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

First you have to launch libreoffice as a service. Run `unogenerator_start`.

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

You don't have to relaunch `unogenerator_start` if you did before

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

### unogenerator_start

Launch libreoffice server, as many process as cpus in your computer.

### unogenerator_stop

Stops libreoffice server instances

### unogenerator_monitor

Monitors your libreoffice server instances

### unogenerator_translation

With this tool you can translate several odt files with one command. It generates .pot and .po files, where you can set your translations. Then run your command again and you'll get your files translated

`unogenerator_translation --from_language es --to_language en --input original.odt --input original2.odt  --output_directory "translation_original"`

You can use --fake to see simulation of your translation

### unogenerator_demo

With this tool you can generate a demo

### unogenerator_demo_concurrent

With this tool you can generate a demo with a lot of files to benchmark your unogenerator serve

## Documentation
You can read [documentation](https://github.com/turulomio/unogenerator/blob/main/doc/unogenerator_documentation_en.odt?raw=true) in doc directory. It has been created with unogenerator.

## Development links

- [LibreOffice code](https://github.com/LibreOffice/core)
- [LibreOffice API](https://api.libreoffice.org/docs/idl/ref/index.html)
- [OpenOffice Forums](https://forum.openoffice.org/en/forum/viewforum.php?f=20)
- [LibreOffice Forums](https://ask.libreoffice.org/)
- [UnoGenerator API](https://coolnewton.mooo.com/doxygen/unogenerator/)


## Changelog
### 0.40.0 (2024-01-27)
- Added to ODS.addRow, ODS.addListOfRows... a formulas parameter. By default is True and formulas are stored as formulas. If False formulas are written as text
- Coverage is 88%
- Unogenerator now needs pydicts-0.13.0

### 0.39.0 (2024-01-03)
- Currency and Percentage classes are now defined in pydicts project
- Tests coverage is now 87%
- Unogenerator now needs pydicts-0.11.0
- Ods.SetCellName now accepts a coord reference instead of a Coord Object
- Created ODS.toDictionaryOfDetailedValues method to get a python dictionary of the whole calc document

### 0.38.0 (2023-12-08)
- Removed statistics to improve speeds
- Added new exceptions: CoordException, RangeException
- Improve localc1989 methods and documentation
- Removed all reusing references
- Added method can_import_uno to allow programs to detect and ignore errors importing uno

### 0.37.0 (2023-11-27)
- Fixed bug with addListOfRows with empty lor. Removed innecesary reusing files
- Fixed bug with addRow with empty row
- Unogenerator now uses pydicts-0.8.0. Old casts and datetime_functions removed

### 0.36.0 (2023-11-24)
- Improved test coverage to 81%
- pydicts module is now used for list of dictionaries methods
- Added new methods to add values without styles. They are used to edit styled templates files. New methods: addCell, addCellMerged, addRow, addColumn, addListOfRows, addListOfColumns

### 0.35.0 (2023-11-05)
- Improved get_values speed

### 0.34.0 (2023-11-05)
- Optimize addRowWithStyle addColumnWithStyle and addListOfRowsWithStyle. 
- Removed cellbycell parameter from this methods.

### 0.33.0 (2023-08-20)
- Added ODT method to insert a HTML block

### 0.32.0 (2023-07-12)
- Fixed bug deleting default sheet
- Migrated to poetry installation method

### 0.31.0 (2022-11-03)
- Added method is_server_working()

### 0.30.0 (2022-10-06)
- [unogenerator_monitor] Added --recommended parameter to set max recmmended memory in Mb. 600Mb per cpu by default.
- Replaced pkg_resources module by importlib.resources as recommended in python.org

### 0.29.0 (2022-09-04)
- Added helper 'helper_split_big_listofrows' to automatically split a sheet in several ones, when is number or rows is bigger than a max_rows parameter

### 0.28.0 (2022-08-20)
- Fixed bug with image size detection.

### 0.27.0 (2022-08-19)
- Helper_totals_by_range now hides total of totals by default.
- ODS.export_pdf now uses one pdf page per sheet by default.
- Added ODS.setSheetStyle. If you use ODS_Standard you can use the portrait style.
- ODT images size can be managed automatically.
- Now you can trim images to remove innecesary space.

### 0.26.0 (2022-04-04)
- Range.start refactorized to Range.c_start. Range.end refactorized to Range.c_end. 
- Added getBlockValuesByRange and getBlockValuesByRangeWithCast, to get values from sheet very fast
- Added a sleep time launch libreoffice instances to try to avoid some issues.
- unogenerator_start detects if LibreOffice instances are already launched.

### 0.25.0 (2022-03-31)
- Added ODT.addStringHyperlinks method to create hyperlinks in ODT documents.
- unogenerator_translator. Removed --translate parameter. Now is implicit.
- unogenerator_translator. Hyperlink strings are now translated and url preserved.
- unogenerator_translator. ODT metadata are now translated.
- ODT.addListOfRows bug fixed when rows were 0.
- Standard style is now default paragraph in ODT.

### 0.24.0 (2022-03-24)
- unogenerator_translator. Entries in po files are now sorted by comment.

### 0.23.0 (2022-03-22)
- Now you can use with statement with ODF classes.
- Added Coord, Range, ColorsNamed to unogenerator import.
- Added skip_up and skip_down parameters to getvalues().
- addListOfRows and addListOfColumns now returns created range.
- Added VerticalBoldCenter style to standard.ods template

### 0.22.0 (2022-03-11)
- Improved installation methods documentation at [INSTALL.md](INSTALL.md)
- Fix bug with missing keys parameter

### 0.21.0 (2022-03-09)
- Improved log and debug system.
- ODS.addListOfRowsWithStyle is now almost instant with lots of cells.
- helper_list_of_ordereddicts now uses ODS.addListOfRowsWithStyle to improve doc generation.

### 0.20.0 (2022-03-06)
- Fixed error when percentage value is None.
- Fixed problem translating demo documents.

### 0.19.0 (2022-02-22)
- ODS.addCell is now used to change values or strings, but doesn't change defined property styles, if they aren't defined

### 0.18.0 (2022-02-10)
- Increased max attempts to reconnect to libreoffice instances.

### 0.17.0 (2022-01-04)
- Added feature to insert images from bytes sequence

### 0.16.0 (2021-11-29)
- Fixed bugs with old methods for numrows and numcolumns

### 0.15.0 (2021-11-24)
- Improved get methods. Added detailed output.

### 0.14.0 (2021-11-21)
- Added styles: Center, Right, VerticalCenter to standard.ods
- Added a helper to create a ods sheet with all internal style names.
- Added statistics module. Added a method to create an internal statistics ods sheet.

### 0.13.0 (2021-11-17)
- Added all needed dependencies (tqdm, colorama, polib, psutil) to setup.py
- Added --pdf argument to `unogenerator_translation` to return translation in a pdf file too.
- If there's a problem getting unogenerator instances process information, it retries 10 attempts by default.

### 0.12.0 (2021-11-14)
- Added `unogenerator_translation` script to translate odf files using .po files.
- Added sortRange method.

### 0.11.0 (2021-11-09)
- Added findall_and_replace method.

### 0.10.0 (2021-11-01)
- Changed vms to rss in psutil to get memory info.
- Auto remove default sheet feature enabled.
- Added find method.
- Added used cpu in monitor info.
- Added method find_and_delete_until_the_end_of_document.

### 0.9.0 (2021-10-28)
- Added monitor to restart server when needed.
- Find and Replace method improved.
- Automatically creates directories to save or export files.

### 0.8.0 (2021-10-18)
- Added statistics and timing withd --debug DEBUG.
- Process migrate to other port if there's a problem creating it.
- Page breaks improved.
- Tables improved.
- Added method LoadStylesFromFile.
- Added deleteAll method to delete all document content.


### 0.7.0 (2021-10-07)
- Added method helper_list_of_ordereddicts_with_totals.
- Concurrent demo has new parameter loops to increment script load.
- Added --backtrace in `unogenerator_start` to try to log something (Experimental)
- Save and export methods use now temporal directories to avoid locks

### 0.6.0 (2021-10-06)
- Added method helper_totals_from_range, 
- Added __str__ method to Coord and Range to write code faster without .string()

### 0.5.0 (2021-10-03)
- Added dispatcher and create service methods.
- Added posibility to link images to document.
- Image width and height are now in cm units.
- Added method addCellNames() for ranged names.
- Improving freezeAndSelect with None values.
- Added range column and number copy additions.

### 0.4.0 (2021-10-02)
- Replaced server.sh by `unogenerator_start` and `unogenerator_stop`
- Added helper_list_of_dict and helper_list_of_ordereddict methods
- Changed vertical alignment to center in all personal styles
- Removed Gentoo confd initd files. Moved to unogenerator_daemon in myportage

### 0.3.0 (2021-10-01)
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
