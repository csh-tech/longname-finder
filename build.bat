@echo off
REM Builds dist\LongNameFinder.exe on Windows.
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --onefile --console --clean ^
  --name LongNameFinder ^
  --icon assets\icon.ico ^
  --add-data "locales;locales" ^
  longname_finder.py
echo.
echo Done: dist\LongNameFinder.exe
pause
