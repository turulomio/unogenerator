
import pytest
from unittest.mock import patch, MagicMock
from unogenerator.monitor import command_monitor, monitor
import sys
from datetime import datetime, timedelta

class MockProcess:
    def __init__(self, pid, name, cmdline):
        self.pid = pid
        self.info = {'name': name, 'cmdline': cmdline, 'pid': pid}
    
    def memory_info(self):
        m = MagicMock()
        m.rss = 1024 * 1024
        return m
        
    def status(self):
        return "running"
        
    def create_time(self):
        return (datetime.now() - timedelta(minutes=5)).timestamp()
        
    def cpu_percent(self, interval=None):
        return 5.0
        
    def connections(self):
        return [1, 2]

def test_command_monitor_one_iteration():
    # Mock processes and scandir
    mock_p = MockProcess(1234, "soffice.bin", ["loffice", "--accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManager", "-env:UserInstallation=file:///tmp/unogenerator2002"])
    
    with patch("unogenerator.monitor.process_iter", return_value=[mock_p]):
        with patch("unogenerator.monitor.scandir", return_value=[]):
            with patch("unogenerator.monitor.sleep", side_effect=InterruptedError("Stop loop")):
                with pytest.raises(InterruptedError):
                    command_monitor(60, 1)

def test_monitor_entry_point():
    # Test the monitor() function which parses args
    with patch("unogenerator.monitor.command_monitor") as mock_command:
        with patch.object(sys, 'argv', ['unogenerator_monitor', '--refresh', '1']):
            monitor()
            mock_command.assert_called_once_with(60, 1)

def test_monitor_with_legacy_dirs():
    # Test identifying temporal directories without processes
    mock_p = MockProcess(1234, "soffice.bin", ["loffice", "--accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManager", "-env:UserInstallation=file:///tmp/unogenerator2002"])
    
    # Mock a directory entry
    mock_entry = MagicMock()
    mock_entry.is_dir.return_value = True
    mock_entry.name = "unogenerator2003" # Different port, no process
    
    with patch("unogenerator.monitor.process_iter", return_value=[mock_p]):
        with patch("unogenerator.monitor.scandir", return_value=[mock_entry]):
            with patch("unogenerator.monitor.sleep", side_effect=InterruptedError):
                with pytest.raises(InterruptedError):
                    command_monitor(60, 1)
                    # We could capture stdout here to verify it mentions unogenerator2003
