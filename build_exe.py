import PyInstaller.__main__

PyInstaller.__main__.run([
    'excel_rounder_gui.py',
    '--name=ExcelRounder',
    '--onefile',
    '--windowed',
    '--add-data=requirements.txt:.',
]) 