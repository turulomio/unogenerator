
import pytest
from unittest.mock import patch, MagicMock
from unogenerator.translation import command_translation, translation
from unogenerator import ODT_Standard, LibreofficeServer
import os
import shutil
import sys

def test_command_translation(libreoffice_server, tmp_path):
    # 1. Create a sample ODT file
    input_odt = str(tmp_path / "test_input.odt")
    with ODT_Standard(server=libreoffice_server) as doc:
        doc.addParagraph("Hello world", "Heading 1")
        doc.addParagraph("This is a test paragraph.", "Standard")
        doc.save(input_odt)
    
    assert os.path.exists(input_odt)
    
    output_dir = str(tmp_path / "translation_output")
    os.makedirs(os.path.join(output_dir, "es"), exist_ok=True)
    po_path = os.path.join(output_dir, "es", "es.po")

    # 2. Run command_translation
    # We mock run_check but we MUST provide a valid PO file for polib to read
    with patch("unogenerator.translation.run_check") as mock_run:
        def side_effect(command, shell=False):
            if "msginit" in command or "msgmerge" in command:
                from polib import POFile, POEntry
                po = POFile()
                po.append(POEntry(msgid="Hello world", msgstr="Hola mundo"))
                po.append(POEntry(msgid="This is a test paragraph.", msgstr="Este es un párrafo de prueba."))
                po.save(po_path)
            return MagicMock(returncode=0)
            
        mock_run.side_effect = side_effect
        
        command_translation(
            from_language="en",
            to_language="es",
            input=[input_odt],
            output_directory=output_dir,
            fake=False, 
            pdf=False
        )

    # 3. Verify output files
    assert os.path.exists(os.path.join(output_dir, "catalogue.pot"))
    assert os.path.exists(os.path.join(output_dir, "es", "test_input.odt"))

def test_translation_entry_point(tmp_path):
    # Test the translation() function which parses args
    input_odt = str(tmp_path / "test_input_entry.odt")
    # Just touch the file to make it exist
    with open(input_odt, "w") as f:
        f.write("dummy")
        
    with patch("unogenerator.translation.command_translation") as mock_command:
        with patch.object(sys, 'argv', ['unogenerator_translation', '--from_language', 'en', '--to_language', 'es', '--input', input_odt]):
            translation()
            mock_command.assert_called_once()

def test_translation_unsupported_file(capsys):
    # Test with a non-ODT file
    with patch("unogenerator.translation.getEntriesFromDocument"):
         command_translation(
            from_language="en",
            to_language="es",
            input=["test.txt"],
            output_directory="./test_unsupported"
        )
    captured = capsys.readouterr()
    # Locale agnostic check
    assert "ODT" in captured.out
    if os.path.exists("./test_unsupported"):
        shutil.rmtree("./test_unsupported")
