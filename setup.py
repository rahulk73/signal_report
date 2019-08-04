import sys
import os.path
from cx_Freeze import setup,Executable
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

# options = {
#     'build_exe': {
#         'include_files':[
#             os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
#             os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll')
#          ],
#     },
# }

build_exe_options = {'packages':['os','tkinter','tkinter.ttk','xlsscript','sqlscript','threading','pickle'],'include_files':[
    os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
    os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
    './img/',
]}
shortcut_table = [
    (
        'DesktopShortcut',
        'DesktopFolder',
        'Signal Report',
        'TARGETDIR',
        '[TARGETDIR]main.exe',
        None,
        None,
        None,
        None,
        None,
        None,
        'TARGETDIR',
    )
]
base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

msi_data = {'Shortcut': shortcut_table}
bdist_msi_options = {'data':msi_data}

setup(
    name = 'SignalReport',
    version = '0.16.14',
    description = 'MVP',
    options = {'build_exe':build_exe_options,'bdist_msi':bdist_msi_options},
    executables = [Executable('main.pyw', base=base,icon="img/icon.ico")]
)