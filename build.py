import PyInstaller.__main__
import os
import shutil

# Очистка предыдущих сборок
if os.path.exists('build'):
    shutil.rmtree('build')
if os.path.exists('dist'):
    shutil.rmtree('dist')

# Сборка EXE
PyInstaller.__main__.run([
    'main.py',
    '--onefile',
    '--windowed',
    '--name=PriceAnalyzer',
    '--icon=icon.ico'  # Убедитесь, что у вас есть файл icon.ico
])
