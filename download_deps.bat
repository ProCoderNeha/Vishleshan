@echo off
REM Download Windows wheel dependencies for offline installation
python -m pip download -r requirements.txt -d deps
python -m pip download pyinstaller -d deps
echo Dependencies downloaded to %CD%\deps
pause
