@echo off
REM Build a single executable for the DRDO internship dashboard
python -m pip install --no-index --find-links deps pyinstaller
pyinstaller --onefile --noconsole --name DRDO_Internship_Dashboard ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --add-data "sample_data.xlsx;." ^
  app.py
echo Build finished. Find the executable in %CD%\dist\DRDO_Internship_Dashboard.exe
pause
