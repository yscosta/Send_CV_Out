#executar pela prompt de comando do PyCharm
#python setup.py build

from cx_Freeze import setup, Executable
setup(
    name = "Send_CV_Out",
    version = "1.0.0",
    options = {"build_exe": {
        'packages': ["os","sys","json", "PyQt5.QtWidgets", "PyQt5.QtCore","PyQt5.QtGui",
                     "smtplib", "ssl", "email"],
        'include_msvcr': True,
        'include_files' : [("e-mail-icon.jpg")],
    }},
    executables = [Executable("Send_CV_Out.py",
                              base="Win32GUI"
                              #compress=True,
                              #copyDependentFiles=True,
                              #appendScriptToExe=True,
                              #appendScriptToLibrary=False,
                              )],
    )