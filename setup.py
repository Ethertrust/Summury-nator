import sys
from cx_Freeze import setup, Executable

includes = []
excludes = []
packages = []
path = sys.path

target = Executable(
    script = "sumhours.py",
    base = 'Win32GUI',
    targetName = "Sumhours.exe",
    icon = "C:\\plus_21081.ico"
    )

setup(
    name="Sumhours",
    version="2.8.0",
    description="Simple Calculator",
    author="Shvedov Egor",
    executables=[target]
    )