from unogenerator import can_import_uno, commons
if can_import_uno():
    from unogenerator.server import command_monitor,  is_port_opened, is_server_working, command_stop, command_start


    def test_command_monitor():
        if commons.is_root():
            command_monitor(True,  1024)
        
    def test_is_server_working():
        is_server_working()
        
    def test_is_port_opened():
        is_port_opened("localhost",  2002)
        is_port_opened("localhost",  90)
        
    def test_command_start():
        if commons.is_root():
            command_start(1,  45000)
        
    def test_command_stop():
        if commons.is_root():
            command_stop()
        
