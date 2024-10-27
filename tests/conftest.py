from unogenerator import LibreofficeServer
# Launch server and stops when this file tests are finished


import pytest

@pytest.fixture(scope="session")
def libreoffice_server():
        # Code to run after all tests in the module

        server=LibreofficeServer()
        yield server
        server.stop()