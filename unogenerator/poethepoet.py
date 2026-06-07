from unogenerator import __version__
from subprocess import run
from sys import argv


def test():
    run(["pytest", "-W", "error"], check=True)

def coverage():
    run(["coverage", "run", "--omit='*uno.py'", "-m", "pytest"], check=True)
    run(["coverage", "report"], check=True)
    run(["coverage", "html"], check=True)

def translate():
        run(["xgettext", "-L", "Python", "--no-wrap", "--no-location", "--from-code=UTF-8", "-o", "unogenerator/locale/unogenerator.pot", "unogenerator/can_import_uno.py", "unogenerator/columnswidth.py", "unogenerator/commands.py", "unogenerator/commons.py", "unogenerator/demo.py", "unogenerator/exceptions.py", "unogenerator/helpers.py", "unogenerator/monitor.py", "unogenerator/poethepoet.py", "unogenerator/translation.py", "unogenerator/types.py", "unogenerator/unogenerator.py"], check=True)
        run(["msgmerge", "-N", "--no-wrap", "-U", "unogenerator/locale/es.po", "unogenerator/locale/unogenerator.pot"], check=True)
        run(["msgmerge", "-N", "--no-wrap", "-U", "unogenerator/locale/fr.po", "unogenerator/locale/unogenerator.pot"], check=True)
        run(["msgmerge", "-N", "--no-wrap", "-U", "unogenerator/locale/ro.po", "unogenerator/locale/unogenerator.pot"], check=True)
        run(["msgfmt", "-cv", "-o", "unogenerator/locale/es/LC_MESSAGES/unogenerator.mo", "unogenerator/locale/es.po"], check=True)
        run(["msgfmt", "-cv", "-o", "unogenerator/locale/fr/LC_MESSAGES/unogenerator.mo", "unogenerator/locale/fr.po"], check=True)
        run(["msgfmt", "-cv", "-o", "unogenerator/locale/ro/LC_MESSAGES/unogenerator.mo", "unogenerator/locale/ro.po"], check=True)


def documentation():
        run(["unogenerator_demo", "--create"], check=True)
        run(["cp", "-f", "unogenerator_documentation_en.odt", "doc/"], check=True)
        run(["cp", "-f", "unogenerator_documentation_en.pdf", "doc/"], check=True)
        run(["cp", "-f", "unogenerator_documentation_es.odt", "doc/"], check=True)
        run(["cp", "-f", "unogenerator_documentation_es.pdf", "doc/"], check=True)
        run(["cp", "-f", "unogenerator_example_en.ods", "doc/"], check=True)
        run(["cp", "-f", "unogenerator_example_en.pdf", "doc/"], check=True)
        run(["cp", "-f", "unogenerator_example_es.ods", "doc/"], check=True)
        run(["cp", "-f", "unogenerator_example_es.pdf", "doc/"], check=True)
        run(["unogenerator_demo", "--remove"], check=True)
...
def docker_build():
    run(["docker", "build", "--tag", "turulomio/unogenerator:latest", "."], check=True)

def docker():
    run(["docker", "run", "-p", "127.0.0.1:2002:2002", "-it", argv[1]], check=True)

def release():
  print("""
Nueva versión:
  * Cambiar la versión y la fecha en __init__.py
  * Cambiar la versión en pyproject.toml
  * Ejecutar otra vez poe release
  * poe translate
  * linguist
  * poe translate
  * poe documentation
  * unogenerator_demo --benchmark
  * pytest
  * git commit -a -m 'unogenerator-{0}'
  * git push
  * Hacer un pull request con los cambios a main
  * Hacer un nuevo tag en GitHub
  * git checkout main
  * git pull
  * poetry build
  * poetry publish --username --password  

""".format(__version__))
