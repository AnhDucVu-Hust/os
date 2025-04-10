import PyInstaller.__main__

PyInstaller.__main__.run([
    'os_document_gui.py',
    '--name=OSDocument',
    '--onefile',
    '--windowed',
    '--add-data=requirements.txt:.'
]) 