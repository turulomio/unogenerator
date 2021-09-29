from setuptools import setup, Command
import site
import os
import platform

class Reusing(Command):
    description = "Download modules from https://github.com/turulomio/reusingcode/"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        from sys import path
        path.append("unogenerator/reusing")
        from github import download_from_github
        download_from_github('turulomio','reusingcode','python/github.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/casts.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/datetime_functions.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/decorators.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/libmanagers.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/percentage.py', 'unogenerator/reusing/')
        download_from_github('turulomio','reusingcode','python/currency.py', 'unogenerator/reusing/')

## Class to define doc command
class Translate(Command):
    description = "Update translations"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        #es
        os.system("xgettext -L Python --no-wrap --no-location --from-code='UTF-8' -o locale/unogenerator.pot *.py unogenerator/*.py unogenerator/reusing/*.py setup.py")
        os.system("msgmerge -N --no-wrap -U locale/es.po locale/unogenerator.pot")
        os.system("msgfmt -cv -o unogenerator/locale/es/LC_MESSAGES/unogenerator.mo locale/es.po")

    
## Class to define doc command
class Documentation(Command):
    description = "Generate documentation for distribution"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        os.system("unogenerator_demo --create")
        os.system("cp -f unogenerator_documentation_en.odt doc/")
        os.system("cp -f unogenerator_documentation_en.pdf doc/")
        os.system("cp -f unogenerator_documentation_es.odt doc/")
        os.system("cp -f unogenerator_documentation_es.pdf doc/")
        os.system("cp -f unogenerator_example_en.ods doc/")
        os.system("cp -f unogenerator_example_en.pdf doc/")
        os.system("cp -f unogenerator_example_es.ods doc/")
        os.system("cp -f unogenerator_example_es.pdf doc/")
        os.system("unogenerator_demo --remove")

class Procedure(Command):
    description = "Show release procedure"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        print("""Nueva versión:
  * Cambiar la versión y la fecha en commons.py
  * Modificar el Changelog en README
  * python setup.py translate
  * linguist
  * python setup.py translate
  * python setup.py uninstall; python setup.py install
  * python setup.py documentation
  * python setup.py doxygen
  * git commit -a -m 'unogenerator-{}'
  * git push
  * Hacer un nuevo tag en GitHub
  * python setup.py sdist upload -r pypi
  * python setup.py uninstall
  * Crea un nuevo ebuild de UNOGENERATOR Gentoo con la nueva versión
  * Subelo al repositorio del portage
  * Crea un nuevo ebuild de UNOGENERATOR_DAEMON Gentoo con la nueva versión
  * Subelo al repositorio del portage

""".format(__version__))

## Class to define doxygen command
class Doxygen(Command):
    description = "Create/update doxygen documentation in doc/html"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        print("Creating Doxygen Documentation")
        os.system("""sed -i -e "41d" doc/Doxyfile""")#Delete line 41
        os.system("""sed -i -e "41iPROJECT_NUMBER         = {}" doc/Doxyfile""".format(__version__))#Insert line 41
        os.system("rm -Rf build")
        os.chdir("doc")
        os.system("doxygen Doxyfile")
        os.system("rsync -avzP -e 'ssh -l turulomio' html/ frs.sourceforge.net:/home/users/t/tu/turulomio/userweb/htdocs/doxygen/unogenerator/ --delete-after")
        os.chdir("..")

## Class to define uninstall command
class Uninstall(Command):
    description = "Uninstall installed files with install"
    user_options = []

    def initialize_options(self):
        pass

    def finalize_options(self):
        pass

    def run(self):
        if platform.system()=="Linux":
            os.system("rm -Rf {}/unogenerator*".format(site.getsitepackages()[0]))
            os.system("rm /usr/bin/unogenerator*")
        else:
            os.system("pip uninstall unogenerator")

########################################################################

## Version of unogenerator captured from commons to avoid problems with package dependencies
__version__= None
with open('unogenerator/commons.py', encoding='utf-8') as f:
    for line in f.readlines():
        if line.find("__version__ =")!=-1:
            __version__=line.split("'")[1]


setup(name='unogenerator',
     version=__version__,
     description='Python module to read and write LibreOffice and MS Office files',
     long_description='Project web page is in https://github.com/turulomio/unogenerator',
     long_description_content_type='text/markdown',
     classifiers=['Development Status :: 4 - Beta',
                  'Intended Audience :: Developers',
                  'Topic :: Software Development :: Build Tools',
                  'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
                  'Programming Language :: Python :: 3',
                 ], 
     keywords='office generator uno pyuno libreoffice',
     url='https://github.com/turulomio/unogenerator',
     author='Turulomio',
     author_email='turulomio@yahoo.es',
     license='GPL-3',
     packages=['unogenerator'],
     install_requires=[],
     entry_points = {'console_scripts': [
                            'unogenerator_demo=unogenerator.demo:main',
                            'unogenerator_daemon_start=unogenerator.daemon:daemon_start',
                            'unogenerator_daemon_stop=unogenerator.daemon:daemon_stop',
                        ],
                    },
     cmdclass={'doxygen': Doxygen,
               'uninstall':Uninstall, 
               'translate': Translate,
               'documentation': Documentation,
               'procedure': Procedure,
               'reusing': Reusing,
              },
     zip_safe=False,
     test_suite = 'unogenerator.tests',
     include_package_data=True
)
