
from unogenerator import can_import_uno
if can_import_uno():

    from unogenerator.demo import main, main_concurrent, demo_ods_standard, demo_odt_standard
    from unogenerator import commons


    def test_demo():
        main(['--create', ])
        main_concurrent(['--create', '--loops', '1' ])
        num_instances, first_port=commons.get_from_process_numinstances_and_firstport()
        demo_ods_standard("es", first_port)
        num_instances, first_port=commons.get_from_process_numinstances_and_firstport()
        demo_odt_standard("es", first_port)

        main(['--remove', ])
        main_concurrent(['--remove', '--loops', '1' ])
