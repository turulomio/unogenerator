from unogenerator.demo import main, main_concurrent


def test_demo():
    main(['--create', ])
    main(['--remove', ])
    
def test_demo_concurrent():
    main_concurrent(['--create', '--loops', '1' ])
    main_concurrent(['--remove', '--loops', '1' ])
    assert True
