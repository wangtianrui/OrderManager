import os
from cx_Freeze import setup, Executable

PYTHON_INSTALL_DIR = r"E:\Anaconda3\envs\mypython"
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

options = {
    'build_exe':
        {
            'excludes': ['Tkinter'],
            'packages': ['pandas', 'os', 'pickle','numpy'],

        }
}

executables = [Executable("Main.py")]
setup(
    name='main',
    version='1.0',
    author="Adam",
    author_email="Omitted",
    options=options,
    executables=executables
)
