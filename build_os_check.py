import PyInstaller.__main__

PyInstaller.__main__.run([
    'os_check_gui.py',
    '--name=OSCheckProcessor',
    '--onefile',
    '--windowed',
    '--add-data=requirements.txt:.'
]) 