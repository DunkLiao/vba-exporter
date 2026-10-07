@echo off
setlocal
title Build VBAExporter

REM ============================================================
REM  Build VBAExporter.exe  -  double-click this file to run it
REM
REM  What it does:
REM    1. Finds Python
REM    2. Installs pywin32 / pyinstaller / pillow if missing
REM    3. Regenerates the application icon
REM    4. Builds dist\VBAExporter.exe with PyInstaller
REM
REM  Optional argument:  build.bat /ci
REM    /ci = no pause and do not open the output folder (for automation)
REM ============================================================

cd /d "%~dp0"

set "NOPAUSE="
if /i "%~1"=="/ci" set "NOPAUSE=1"

echo ============================================================
echo    Build VBAExporter.exe
echo    Project: %cd%
echo ============================================================
echo.

REM --- 1/5  Locate Python -------------------------------------
set "PY="
where python >nul 2>nul && set "PY=python"
if defined PY (
    %PY% -c "import sys" >nul 2>nul || set "PY="
)
if not defined PY (
    where py >nul 2>nul && set "PY=py -3"
)
if defined PY (
    %PY% -c "import sys" >nul 2>nul || set "PY="
)
if not defined PY (
    echo [ERROR] A working Python was not found on PATH.
    echo         Install Python 3.10 or newer and enable "Add Python to PATH",
    echo         then run this file again.
    goto :fail
)
echo [1/5] Python found:
%PY% --version
echo.

REM --- 2/5  Required packages ---------------------------------
echo [2/5] Checking required packages (pywin32, pyinstaller, pillow)...
%PY% -c "import win32com.client, PyInstaller, PIL" >nul 2>nul
if errorlevel 1 (
    echo       Some packages are missing. Installing them now, please wait...
    if exist "requirements.txt" (
        %PY% -m pip install --disable-pip-version-check -r requirements.txt
    ) else (
        %PY% -m pip install --disable-pip-version-check pywin32 pyinstaller pillow
    )
    if errorlevel 1 (
        echo [ERROR] Failed to install the required packages.
        echo         Check your internet connection and try again.
        goto :fail
    )
) else (
    echo       All required packages are already installed.
)
echo.

REM --- 3/5  Application icon ----------------------------------
echo [3/5] Generating the application icon...
if exist "make_icon.py" (
    %PY% make_icon.py
    if errorlevel 1 (
        echo [ERROR] Failed to generate the application icon.
        goto :fail
    )
) else (
    echo       make_icon.py not found - skipping icon generation.
)
echo.

REM --- 4/5  Make sure the target executable is not in use -----
echo [4/5] Checking if the target file is in use...
tasklist /fi "imagename eq VBAExporter.exe" 2>nul | find /i "VBAExporter.exe" >nul
if not errorlevel 1 (
    echo [ERROR] VBAExporter.exe is currently running.
    echo         Please close it, then run this build again.
    goto :fail
)
echo       OK - the target file is not in use.
echo.

REM --- 5/5  Build ---------------------------------------------
echo [5/5] Building with PyInstaller (this may take a few minutes)...
if exist "VBAExporter.spec" (
    %PY% -m PyInstaller --noconfirm VBAExporter.spec
) else (
    %PY% -m PyInstaller --noconfirm --clean --onefile --windowed --name VBAExporter --icon=icon.ico --hidden-import win32timezone export_vba.py
)
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    goto :fail
)
echo.

if not exist "dist\VBAExporter.exe" (
    echo [ERROR] The build finished, but dist\VBAExporter.exe was not found.
    goto :fail
)

for %%F in ("dist\VBAExporter.exe") do set "EXESIZE=%%~zF"

echo ============================================================
echo    BUILD SUCCESSFUL
echo    Output : %cd%\dist\VBAExporter.exe
echo    Size   : %EXESIZE% bytes
echo ============================================================
echo.
if not defined NOPAUSE (
    start "" "%cd%\dist"
    echo Press any key to close this window...
    pause >nul
)
exit /b 0

:fail
echo.
echo ============================================================
echo    BUILD FAILED - see the messages above
echo ============================================================
echo.
if not defined NOPAUSE (
    echo Press any key to close this window...
    pause >nul
)
exit /b 1
