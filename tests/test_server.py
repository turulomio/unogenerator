from unogenerator import can_import_uno, commons
if can_import_uno():
    from unogenerator.monitor import command_monitor


    def test_command_monitor():
        if commons.is_root():
            command_monitor(True,  1024)


# THIS TEST SHOULD NOT BE USED. GITHUB ACTIONS STOPS THE SERVER AND THE REST OF TESTS FAIL        
#    def test_command_start():
#        if commons.is_root():
#            command_start(1,  2002)
#        
#    def test_command_stop():
#        if commons.is_root():
#            command_stop()
        
