from unogenerator.reusing.github import download_from_github
from unogenerator.reusing.file_functions import replace_in_file
from unogenerator import __version__
from os import system
from sys import argv


def pytest():
    system("pytest")
    
def coverage():
    system("coverage run --omit='*/reusing/*,*uno.py' -m pytest && coverage report")


def reusing():
    """
        Actualiza directorio reusing
        poe reusing
        poe reusing --local
    """
    local=False
    if len(argv)==2 and argv[1]=="--local":
        local=True
        print("Update code in local without downloading was selected with --local")
    if local==False:
        download_from_github('turulomio','reusingcode','python/github.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/casts.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/datetime_functions.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/listdict_functions.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/file_functions.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/decorators.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/libmanagers.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/percentage.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/currency.py', 'unogenerator/reusing/')
    
    replace_in_file("unogenerator/reusing/casts.py","from currency","from unogenerator.reusing.currency")
    replace_in_file("unogenerator/reusing/casts.py","from percentage","from unogenerator.reusing.percentage")
    replace_in_file("unogenerator/reusing/listdict_functions.py","from casts","from unogenerator.reusing.casts")


def translate():
        system("xgettext -L Python --no-wrap --no-location --from-code='UTF-8' -o locale/unogenerator.pot  unogenerator/*.py unogenerator/reusing/*.py ")
        system("msgmerge -N --no-wrap -U locale/es.po locale/unogenerator.pot")
        system("msgfmt -cv -o unogenerator/locale/es/LC_MESSAGES/unogenerator.mo locale/es.po")
        system("msgfmt -cv -o unogenerator/locale/en/LC_MESSAGES/unogenerator.mo locale/en.po")

    
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
  * Cambiar la versión y la fecha en commons.py
  * Cambiar la versión en pyproject.toml
  * Modificar el Changelog en README
  * poe translate
  * linguist
  * poe translate
  * poe documentation
  * poe doxygen
  * git commit -a -m 'unogenerator-{0}'
  * git push
  * Hacer un nuevo tag en GitHub
  * poetry build
  * poetry publish --username --password  
  * Crea un nuevo ebuild de UNOGENERATOR Gentoo con la nueva versión
  * Subelo al repositorio del portage

""".format(__version__))

### Class to define doxygen command
#class Doxygen(Command):
#    description = "Create/update doxygen documentation in doc/html"
#
#    user_options = [
#      # The format is (long option, short option, description).
#      ( 'user=', None, 'Remote ssh user'),
#      ( 'directory=', None, 'Remote ssh path'),
#      ( 'port=', None, 'Remote ssh port'),
#      ( 'server=', None, 'Remote ssh server'),
#  ]
#
#    def initialize_options(self):
#        self.user="root"
#        self.directory="/var/www/html/doxygen/unogenerator/"
#        self.port=22
#        self.server="127.0.0.1"
#
#    def finalize_options(self):
#        pass
#
#    def run(self):
#        print("Creating Doxygen Documentation")
#        os.system("""sed -i -e "41d" doc/Doxyfile""")#Delete line 41
#        os.system("""sed -i -e "41iPROJECT_NUMBER         = {}" doc/Doxyfile""".format(__version__))#Insert line 41
#        os.system("rm -Rf build")
#        os.chdir("doc")
#        os.system("doxygen Doxyfile")      
#        command=f"""rsync -avzP -e 'ssh -l {self.user} -p {self.port} ' html/ {self.server}:{self.directory} --delete-after"""
#        print(command)
#        os.system(command)
#        os.chdir("..")
  
### Class to define uninstall command
#class Uninstall(Command):
#    description = "Uninstall installed files with install"
#    user_options = []
#
#    def initialize_options(self):
#        pass
#
#    def finalize_options(self):
#        pass
#
#    def run(self):
#        if platform.system()=="Linux":
#            os.system("rm -Rf {}/unogenerator*".format(site.getsitepackages()[0]))
#            os.system("rm /usr/bin/unogenerator*")
#        else:
#            os.system("pip uninstall unogenerator")



def docker_build():
    system("docker build --tag turulomio/unogenerator:latest .")

def docker():
    system(f"docker  run  -p 127.0.0.1:2002:2002 -it {argv[1]}")
