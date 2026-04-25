import pytest
from unittest.mock import MagicMock
import socket
from unogenerator import can_import_uno

if can_import_uno():
    from unogenerator import ODS_Standard, LibreofficeServer

    def test_ods_creation_connection_failure_cleans_server(monkeypatch):
        """
        Tests that if ODS creation fails to connect to the LibreOffice server,
        the server process is still properly terminated and cleaned up.
        """
        # 1. Mock subprocess.Popen to simulate LibreOffice process starting
        mock_process = MagicMock()
        mock_process.pid = 99999  # A dummy PID
        mock_process.poll.return_value = None  # Simulate process is still running
        mock_process.stdout.close = MagicMock()
        mock_process.stderr.close = MagicMock()
        mock_process.terminate = MagicMock()
        mock_process.wait = MagicMock()

        def mock_libreoffice_server_start(self):
            self.port = 2002  # Use a fixed port for predictability
            self.process = mock_process

        monkeypatch.setattr(LibreofficeServer, 'start', mock_libreoffice_server_start)

        # 2. Mock socket.socket to simulate connection refusal
        def mock_socket_init(*args, **kwargs):
            mock_sock = MagicMock()
            mock_sock.connect_ex.return_value = 111  # Simulate errno.ECONNREFUSED
            mock_sock.bind.return_value = None
            mock_sock.getsockname.return_value = ('127.0.0.1', 2002)
            mock_sock.close.return_value = None
            return mock_sock

        monkeypatch.setattr(socket, 'socket', mock_socket_init)

        # 3. Reduce maxtries in ODF.__init__ for faster test execution
        monkeypatch.setattr(ODS_Standard, 'maxtries', 2) # ODS_Standard inherits from ODF, so setting it here affects instances

        server_instance = None
        with ODS_Standard() as doc:
            server_instance = doc.server
            assert doc.ctx is None  # Connection should have failed

        assert server_instance is not None
        assert server_instance.process is None  # Assert server process reference is cleared
        mock_process.terminate.assert_called_once()  # Assert terminate was called
        mock_process.stdout.close.assert_called_once()  # Assert stdout was closed
        mock_process.stderr.close.assert_called_once()  # Assert stderr was closed