@echo off
python -m pip install --upgrade pip
python -m pip install openpyxl pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name "MDLSZ_Gyorspont" mdlsz_gyorspont.py
echo.
echo Kesz! A futtathato fajl helye: dist\MDLSZ_Gyorspont.exe
pause
