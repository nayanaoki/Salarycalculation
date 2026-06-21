@echo off
rem ============================================
rem  Payroll app build (PyInstaller onefile)
rem  NOTE: keep this file ASCII-only. Japanese text
rem  in a .bat breaks cmd parsing on CP932 systems.
rem ============================================
cd /d "%~dp0"

echo [1/3] Installing dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 goto :error
python -m pip install pyinstaller
if errorlevel 1 goto :error

echo [2/3] Cleaning previous build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [3/3] Building exe...
python -m PyInstaller --noconfirm payroll.spec
if errorlevel 1 goto :error

echo.
echo Done: dist\payroll exe  (Japanese name set in payroll.spec)
echo Next: compile installer.iss with Inno Setup.
goto :end

:error
echo.
echo *** BUILD FAILED *** see messages above.

:end
pause
