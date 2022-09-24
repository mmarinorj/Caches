import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"], "includes": ["PyQt5", "pandas"]}

# GUI applications require a different base on Windows (the default is for
# a console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Cachês",
    version="0.1",
    description="Totalizador de cachês e transportes",
    options={"build_exe": build_exe_options},
    executables=[Executable("caches.py", base=base)]
)