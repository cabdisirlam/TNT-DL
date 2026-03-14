@echo off
setlocal enableextensions

set "QUIET=0"
if /I "%~1"=="/quiet" set "QUIET=1"

set "INSTALL_DIR=%LOCALAPPDATA%\Programs\NT_DL"
set "START_MENU_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\NT_DL"
set "START_MENU_LINK=%START_MENU_DIR%\NT_DL.lnk"
set "DESKTOP_LINK=%USERPROFILE%\Desktop\NT_DL.lnk"

taskkill /F /IM NT_DL.exe >nul 2>&1

del /F /Q "%START_MENU_LINK%" >nul 2>&1
rmdir "%START_MENU_DIR%" >nul 2>&1
del /F /Q "%DESKTOP_LINK%" >nul 2>&1

del /F /Q "%INSTALL_DIR%\NT_DL.exe" >nul 2>&1
del /F /Q "%INSTALL_DIR%\kdl_a.ico" >nul 2>&1

reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NT_DL" /f >nul 2>&1

if "%QUIET%"=="0" (
    echo.
    echo NT_DL uninstalled.
)

start "" cmd /c "timeout /t 2 /nobreak >nul & del /f /q \"%INSTALL_DIR%\uninstall.cmd\" >nul 2>&1 & rmdir \"%INSTALL_DIR%\" >nul 2>&1"
exit /b 0
