# UnoGenerator

## Description
Python module to generate Libreoffice documents (ODT and ODS) programatically.

Morever, you can export them to (.xlsx, .docx, .pdf) easyly.

It uses Libreoffice uno module, so you need Libreoffice to be installed in your system.

## Installation
You can use pip to install this python package:

`pip install unogenerator`

## ODT 'Hello World' example

This is a Hello World example. You get the example in odt, docx and pdf formats:

First you have to launch libreoffice as a service. Run `unogenerator_start`.

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

## ODS 'Hello World' example

This is a Hello World example. You'll get example files in ods, xlsx and pdf formats:

You don't have to relaunch `unogenerator_start` if you did before

```python
from unogenerator import ODS_Standard
doc=ODS_Standard()
doc.addCellMergedWithStyle("A1:E1", "Hello world", style="BoldCenter")
doc.save("hello_world.ods")
doc.export_xlsx("hello_world.xlsx")
doc.export_pdf("hello_world.pdf")
doc.close()
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

`unogenerator_translation --from_language es --to_language en --input original.odt --input original2.odt  --output_directory "translation_original" --translate`

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
- [LibreOffice Forums](https://forum.openoffice.org/en/forum/viewforum.php?f=20)
- [UnoGenerator API](http://turulomio.users.sourceforge.net/doxygen/unogenerator/)


## Changelog

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
