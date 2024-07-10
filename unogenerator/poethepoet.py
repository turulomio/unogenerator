from unogenerator import __version__
from os import system
from sys import argv


def test():
    system("pytest")
    
def coverage():
    system("coverage run --omit='*uno.py' -m pytest && coverage report && coverage html")

def translate():
        system("xgettext -L Python --no-wrap --no-location --from-code='UTF-8' -o unogenerator/locale/unogenerator.pot  unogenerator/*.py")
        system("msgmerge -N --no-wrap -U unogenerator/locale/es.po unogenerator/locale/unogenerator.pot")
        system("msgfmt -cv -o unogenerator/locale/es/LC_MESSAGES/unogenerator.mo unogenerator/locale/es.po")
        system("msgfmt -cv -o unogenerator/locale/en/LC_MESSAGES/unogenerator.mo unogenerator/locale/en.po")

    
def documentation():
        system("unogenerator_demo --create")
        system("cp -f unogenerator_documentation_en.odt doc/")
        system("cp -f unogenerator_documentation_en.pdf doc/")
        system("cp -f unogenerator_documentation_es.odt doc/")
        system("cp -f unogenerator_documentation_es.pdf doc/")
        system("cp -f unogenerator_example_en.ods doc/")
        system("cp -f unogenerator_example_en.pdf doc/")
        system("cp -f unogenerator_example_es.ods doc/")
        system("cp -f unogenerator_example_es.pdf doc/")
        system("unogenerator_demo --remove")

def release():
    print("""Nueva versión:
  * Cambiar la versión y la fecha en __init__.py
  * Cambiar la versión en pyproject.toml
  * Ejecutar otra vez poe release
  * git checkout -b unogenerator-{0}
  * Modificar el Changelog en README.md
  * poe coverage con pyvenv systempackages=false
  * poe coverage con pyvenv systempackages=true
  * poe translate
  * linguist
  * poe translate
  * poe documentation
  * git commit -a -m 'unogenerator-{0}'
  * git push
  * Hacer un pull request con los cambios a main
  * Hacer un nuevo tag en GitHub
  * git checkout main
  * git pull
  * poetry build
  * poetry publish --username --password  
  * Crea un nuevo ebuild de UNOGENERATOR Gentoo con la nueva versión
  * Subelo al repositorio del portage

""".format(__version__))

def docker_build():
    system("docker build --tag turulomio/unogenerator:latest .")

def docker():
    system(f"docker  run  -p 127.0.0.1:2002:2002 -it {argv[1]}")
