@echo off
chcp 65001 >nul
title Econnect Employee Manager - Setup

echo.
echo  ============================================
echo   Econnect Telecom - Employee Manager
echo   Creat by ENG Omar Gamal
echo   Setup v1.0
echo  ============================================
echo.

:: STEP 1: Check Python
echo [1/3] Checking Python...

py --version >nul 2>&1
if %errorlevel% == 0 (
    echo       OK - Python found via py launcher.
    set PYTHON=py
    goto install_packages
)

python --version >nul 2>&1
if %errorlevel% == 0 (
    echo       OK - Python found.
    set PYTHON=python
    goto install_packages
)

python3 --version >nul 2>&1
if %errorlevel% == 0 (
    echo       OK - Python3 found.
    set PYTHON=python3
    goto install_packages
)

echo.
echo  ============================================
echo   Python is NOT installed on this computer.
echo.
echo   1. Go to: https://www.python.org/downloads/
echo   2. Download and install Python
echo   3. IMPORTANT: check "Add Python to PATH"
echo   4. Run setup.bat again
echo  ============================================
echo.
pause
exit /b 1

:install_packages
echo.
echo [2/3] Installing required packages...
echo       This may take a minute, please wait...
echo.

%PYTHON% -m pip install --upgrade pip --quiet --no-warn-script-location
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: pip upgrade failed. Continuing anyway...
)

%PYTHON% -m pip install openpyxl requests --quiet --no-warn-script-location
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Could not install packages.
    echo  Try right-clicking setup.bat and "Run as Administrator"
    echo.
    pause
    exit /b 1
)

echo       OK - All packages installed successfully.

:: STEP 3: Desktop shortcut
echo.
echo [3/3] Creating desktop shortcut...

set "APP_PATH=%~dp0employee_manager.py"
set "SHORTCUT=%USERPROFILE%\Desktop\Econnect Employee Manager.lnk"

powershell -NoProfile -Command "$ws=New-Object -ComObject WScript.Shell; $sc=$ws.CreateShortcut('%SHORTCUT%'); $sc.TargetPath='pythonw'; $sc.Arguments='\"%APP_PATH%\"'; $sc.WorkingDirectory='%~dp0'; $sc.Description='Econnect Employee Manager'; $sc.Save()" >nul 2>&1

if exist "%SHORTCUT%" (
    echo       OK - Shortcut created on Desktop.
) else (
    echo       Note: Shortcut not created, but app will still work.
)

:: Launch
echo.
echo  ============================================
echo   Done! Starting Econnect Employee Manager...
echo  ============================================
echo.
timeout /t 2 /nobreak >nul

start "" pythonw "%~dp0employee_manager.py"
if %errorlevel% neq 0 (
    start "" python "%~dp0employee_manager.py"
)

exit
