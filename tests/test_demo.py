from unogenerator import can_import_uno
if can_import_uno():
    from unogenerator.demo import demo_ods_standard, demo_odt_standard

    def test_demo(libreoffice_server):
        demo_ods_standard("es",  server=libreoffice_server)
        demo_odt_standard("es",  server=libreoffice_server)

