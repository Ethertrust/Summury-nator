import sys
from cx_Freeze import setup, Executable

target = Executable(
    script = "sumhours.py",
    base = 'Win32GUI',
    targetName = "Sumhours.exe",
    icon = "C:\\Python34\\Scripts\\xmlparsesum\\plus.ico"
    )

setup(
    name="Sumhours",
    version="2.0.0",
    description="Simple Calculator",
    author="Shvedov Egor",
    executables=[target]
    )