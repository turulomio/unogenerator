
from unogenerator import can_import_uno
if can_import_uno():
    from unogenerator.demo import main, main_concurrent, demo_ods_standard, demo_odt_standard

    def test_demo():
        main(['--create', ])
        main_concurrent(['--create', '--loops', '1' ])
        demo_ods_standard("es")
        demo_odt_standard("es")

        main(['--remove', ])
        main_concurrent(['--remove', '--loops', '1' ])
