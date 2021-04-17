#executar pela prompt de comando do PyCharm
#python setup.py build

import sys
import os
import shutil
from cx_Freeze import setup, Executable

base = None
if sys.platform == "win32":
    base = "Win32GUI"

packages = ["os", "sys", "json", "PyQt5", "smtplib", "ssl", "email"]

includes = []

excludes =["sqlite3", "tensorboard", "tensorflow", "test", "tkinter", "graphviz",
           "keras", "matplotlib", "numba", "scipy", "sqlalchemy"]

include_files = ["e-mail-icon.jpg"]

options = {
    "build_exe": {"includes": includes,
                  "excludes": excludes,
                  "zip_include_packages": packages,
                  "include_files": include_files}
}

executables = [Executable("Send_CV_Out.py", base=base)]

setup(
    name="Send_CV_Out",
    version="1.1",
    description="Send CV Out by e-mail from a e-mail list.",
    options=options,
    executables=executables,
)

# Transfere a base de dados para área do executável.
if os.path.exists('build//exe.win-amd64-3.7'):
    os.makedirs('build//exe.win-amd64-3.7//DB')
    shutil.copyfile('DB//emails.csv', 'build//exe.win-amd64-3.7//DB//emails.csv')
    shutil.copyfile('DB//e-mail-struct.json', 'build//exe.win-amd64-3.7//DB//e-mail-struct.json')