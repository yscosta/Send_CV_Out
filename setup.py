#executar pela prompt de comando do PyCharm
#python setup.py build

# from cx_Freeze import setup, Executable
#
# packages = ["os", "sys", "json", "PyQt5.QtWidgets", "PyQt5.QtCore", "PyQt5.QtGui",
#              "smtplib", "ssl", "email"]
#
# # packages = ["os", "sys", "json", "PyQt5", "smtplib", "ssl", "email"]
#
# build_exe = {'packages': packages,
#              'include_msvcr': True,
#              'include_files': [("e-mail-icon.jpg")], }
#
# setup(name = "Send_CV_Out",
#       version = "1.0.0",
#       options = {"build_exe": build_exe},
#       executables = [Executable("Send_CV_Out.py",
#                                 base="Win32GUI"
#                                 #compress=True,
#                                 #copyDependentFiles=True,
#                                 #appendScriptToExe=True,
#                                 #appendScriptToLibrary=False,
#                               )
#                    ],
#       )

# import sys
# from cx_Freeze import setup, Executable
#
# base = None
# if sys.platform == "win32":
#     base = "Win32GUI"
#
# options = {"build_exe": {"includes": "atexit"}}
#
# executables = [Executable("Send_CV_Out.py", base=base)]
#
# setup(
#     name="Send_CV_Out",
#     version="1.0",
#     description="Send CV Out by e-mail from a e-mail list.",
#     options=options,
#     executables=executables,
# )


import sys
from cx_Freeze import setup, Executable

base = None
if sys.platform == "win32":
    base = "Win32GUI"

packages = ["os", "sys", "json", "PyQt5", "smtplib", "ssl", "email"]

# options = {
#     "build_exe": {"includes": "atexit", "zip_include_packages": ["PyQt5"]}
# }

options = {
    "build_exe": {"includes": "atexit", "zip_include_packages": packages}
}

executables = [Executable("Send_CV_Out.py", base=base)]

setup(
    name="Send_CV_Out",
    version="1.1",
    description="Send CV Out by e-mail from a e-mail list.",
    options=options,
    executables=executables,
)