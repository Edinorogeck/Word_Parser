import sys
from cx_Freeze import setup, Executable

setup(
    name = "Configuration programm",
    version = "1.0",
    description = "Программа для парсинга .docx файла и создания папки с таблицами сообщений",
    executables = [Executable("main.py", base = "Win32GUI")])