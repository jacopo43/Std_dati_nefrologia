@echo off
setlocal
cd /d "%~dp0"
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m PyInstaller --noconfirm --clean PatientConverter.spec
if not exist "..\PatientConverter.exe" copy /Y "dist\PatientConverter.exe" "..\PatientConverter.exe"
echo.
echo Executable created at:
echo %~dp0..\PatientConverter.exe
pause
