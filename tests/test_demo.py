
from unogenerator import can_import_uno
if can_import_uno():
    from unogenerator.demo import main, demo_ods_standard, demo_odt_standard

    def test_demo_commonserver_sequential():
        main(['--create', '--type', 'COMMONSERVER_SEQUENTIAL' ])
        main(['--remove', ])
        
        
    def test_demo_commonserver_concurrent_process():
        main(['--create', '--type', 'COMMONSERVER_CONCURRENT_PROCESS' ])
        main(['--remove', ])
    def test_demo_commonserver_concurrent_threads():
        main(['--create', '--type', 'COMMONSERVER_CONCURRENT_THREADS' ])
        main(['--remove', ])
    def test_demo_sequential():
        main(['--create', '--type', 'SEQUENTIAL' ])
        main(['--remove', ])
    def test_demo_concurrent_process():
        main(['--create', '--type', 'CONCURRENT_PROCESS' ])
        main(['--remove', ])
    def test_demo_concurrent_threads():
        main(['--create', '--type', 'CONCURRENT_THREADS' ])
        main(['--remove', ])

        demo_ods_standard("es")
        demo_odt_standard("es")

