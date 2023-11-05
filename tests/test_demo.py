from unogenerator.demo import main, main_concurrent, demo_ods_standard, demo_odt_standard
from unogenerator import commons


def test_demo():
    main(['--create', ])
    main(['--remove', ])
    
def test_demo_concurrent():
    main_concurrent(['--create', '--loops', '1' ])
    main_concurrent(['--remove', '--loops', '1' ])
    assert True

def test_demo_ods_standard():
    num_instances, first_port=commons.get_from_process_numinstances_and_firstport()
    demo_ods_standard("es", first_port)
def test_demo_odt_standard():
    num_instances, first_port=commons.get_from_process_numinstances_and_firstport()
    demo_odt_standard("es", first_port)
