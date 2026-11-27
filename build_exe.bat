@echo off
python -m pip install --upgrade pip
python -m pip install openpyxl pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "MDLSZ_Gyorspont" --add-data "assets;assets" mdlsz_gyorspont.py
echo.
echo Kesz! A futtathato fajl helye: dist\MDLSZ_Gyorspont.exe
pause
