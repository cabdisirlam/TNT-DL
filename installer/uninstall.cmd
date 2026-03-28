@echo off
setlocal enableextensions

set "QUIET=0"
if /I "%~1"=="/quiet" set "QUIET=1"

set "INSTALL_DIR=%LOCALAPPDATA%\Programs\NT_DL"
set "START_MENU_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\NT_DL"
set "START_MENU_LINK=%START_MENU_DIR%\NT_DL.lnk"
set "DESKTOP_LINK=%USERPROFILE%\Desktop\NT_DL.lnk"
set "RUNTIME_TEMP_DIR=%LOCALAPPDATA%\N"
set "APPDATA_SETTINGS_DIR=%APPDATA%\KDL"
set "LOCALAPPDATA_SETTINGS_DIR=%LOCALAPPDATA%\KDL"
set "HOME_SETTINGS_DIR=%USERPROFILE%\.kdl"
set "INSTALLED_LAUNCHER=%INSTALL_DIR%\NT_DL.exe"
set "INSTALLED_APP=%INSTALL_DIR%\NT_DL_app.exe"

taskkill /F /T /IM NT_DL.exe /IM NT_DL_app.exe >nul 2>&1

for /L %%I in (1,1,12) do (
    tasklist /FI "IMAGENAME eq NT_DL.exe" 2>nul | find /I "NT_DL.exe" >nul
    if errorlevel 1 (
        tasklist /FI "IMAGENAME eq NT_DL_app.exe" 2>nul | find /I "NT_DL_app.exe" >nul
        if errorlevel 1 goto cleanup_done
    )
    timeout /t 1 /nobreak >nul
)

:cleanup_done

del /F /Q "%START_MENU_LINK%" >nul 2>&1
rmdir "%START_MENU_DIR%" >nul 2>&1
del /F /Q "%DESKTOP_LINK%" >nul 2>&1

del /F /Q "%INSTALLED_LAUNCHER%" >nul 2>&1
del /F /Q "%INSTALLED_APP%" >nul 2>&1
del /F /Q "%INSTALL_DIR%\kdl_a.ico" >nul 2>&1

reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NT_DL" /f >nul 2>&1

if "%QUIET%"=="0" (
    echo.
    echo NT_DL uninstalled.
)

start "" cmd /c "timeout /t 2 /nobreak >nul & del /f /q \"%INSTALL_DIR%\uninstall.cmd\" >nul 2>&1 & rmdir /s /q \"%INSTALL_DIR%\" >nul 2>&1 & rmdir /s /q \"%RUNTIME_TEMP_DIR%\" >nul 2>&1 & rmdir /s /q \"%APPDATA_SETTINGS_DIR%\" >nul 2>&1 & rmdir /s /q \"%LOCALAPPDATA_SETTINGS_DIR%\" >nul 2>&1 & rmdir /s /q \"%HOME_SETTINGS_DIR%\" >nul 2>&1"
exit /b 0
