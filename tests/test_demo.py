from unogenerator import can_import_uno
if can_import_uno():
    from unogenerator.demo import demo_ods_standard, demo_odt_standard

#    def test_demo_commonserver_sequential():
#        main(['--create', '--type', 'COMMONSERVER_SEQUENTIAL' ])
#        main(['--remove', ])
#        
#        
#    def test_demo_commonserver_concurrent_process():
#        main(['--create', '--type', 'COMMONSERVER_CONCURRENT_PROCESS' ])
#        main(['--remove', ])
#    def test_demo_commonserver_concurrent_threads():
#        main(['--create', '--type', 'COMMONSERVER_CONCURRENT_THREADS' ])
#        main(['--remove', ])
#    def test_demo_sequential():
#        main(['--create', '--type', 'SEQUENTIAL' ])
#        main(['--remove', ])
#    def test_demo_concurrent_process():
#        main(['--create', '--type', 'CONCURRENT_PROCESS' ])
#        main(['--remove', ])
#    def test_demo_concurrent_threads():
#        main(['--create', '--type', 'CONCURRENT_THREADS' ])
#        main(['--remove', ])
    # Launch server and stops when this file tests are finished
#    server=LibreofficeServer()
#    def teardown_module(module):
#          # Code to run after all tests in the module
#        server.stop()

    def test_demo(libreoffice_server):
        demo_ods_standard("es",  server=libreoffice_server)
        demo_odt_standard("es",  server=libreoffice_server)

