import sys
from cx_Freeze import setup, Executable

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
"packages": ["xlrd","lxml"],
"excludes": [],
#"include_files": ["setting.ini", "res"]
}

#Executable("TestGUI.py", base=base, targetName="Joker3DAvatarMgr.exe", compress=True, icon="res/ico/icon.ico")

executables = [Executable("TestGUI.py", base=base, targetName="CaseConvertTool.exe", compress=True)]

setup( name = "setup", version = "0.1", description = "CaseConvertTool",
author = "zhangzheyu",
author_email = "zhangzheyuk@gmail.com",
options = {"build_exe": build_exe_options},
executables = executables,
)