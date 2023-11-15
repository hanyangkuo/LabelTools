-To build the python execute file using the command below:
>>> pyinstaller -w --hidden-import=xlrd  -n "LabelTools" app.py

x) Pyinstaller “Failed to execute script pyi_rth_pkgres” and missing packages
>>> pip uninstall pyinstaller
>>> pip install https://github.com/pyinstaller/pyinstaller/archive/develop.zip