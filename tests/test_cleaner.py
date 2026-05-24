
import pytest
from unittest.mock import patch, MagicMock
from unogenerator.monitor import command_cleaner
import os
import shutil

def test_command_cleaner(tmp_path):
    # Mocking scandir to return our test entries
    # But wait, command_cleaner uses hardcoded "/tmp"
    # To test it without actually touching /tmp, we would need to mock scandir or change the function to accept a path.
    # However, since it's a "cleaner", we can test it by creating specific files in /tmp if we are careful.
    
    test_dir = "/tmp/unogenerator_pytest_dir"
    test_file = "/tmp/unogenerator_pytest_file"
    
    os.makedirs(test_dir, exist_ok=True)
    with open(test_file, "w") as f:
        f.write("test")
        
    assert os.path.exists(test_dir)
    assert os.path.exists(test_file)
    
    with patch("unogenerator.monitor.run") as mock_run:
        command_cleaner()
        
        # Verify killall was called
        mock_run.assert_called_with(['killall', '-9', 'soffice.bin'], check=False)
        
    # Verify files are gone
    assert not os.path.exists(test_dir)
    assert not os.path.exists(test_file)

import sys

def test_cleaner_entry_point():
    # Test the cleaner() function which parses args
    with patch("unogenerator.monitor.command_cleaner") as mock_command:
        with patch.object(sys, 'argv', ['unogenerator_cleaner']):
            from unogenerator.monitor import cleaner
            cleaner()
            mock_command.assert_called_once()

def test_cleaner_no_files():
    # Test that it doesn't crash if no files exist
    with patch("unogenerator.monitor.scandir", return_value=[]):
        with patch("unogenerator.monitor.run"):
            command_cleaner()
            # Should not raise exception
