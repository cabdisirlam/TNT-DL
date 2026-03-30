@echo off
setlocal enableextensions

set "QUIET=0"
if /I "%~1"=="/quiet" set "QUIET=1"

set "APP_NAME=NT_DL"
set "APP_DISPLAY=NT_DL"
set "APP_VERSION=1.1.73"
set "PUBLISHER=NT_DL"
set "LOG_FILE=%TEMP%\NT_DL_install.log"
set "EXIT_CODE=1"

>"%LOG_FILE%" echo [%date% %time%] NT_DL install start
>>"%LOG_FILE%" echo SourceDir=%~dp0

set "INSTALL_DIR=%LOCALAPPDATA%\Programs\NT_DL"
set "START_MENU_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\NT_DL"
set "START_MENU_LINK=%START_MENU_DIR%\NT_DL.lnk"
set "DESKTOP_LINK=%USERPROFILE%\Desktop\NT_DL.lnk"
set "SOURCE_ROOT=%~dp0app\NT_DL"
set "INSTALLED_EXE=%INSTALL_DIR%\NT_DL.exe"

rem Ensure old app instances do not lock binaries during upgrade.
>>"%LOG_FILE%" echo Closing running NT_DL processes
taskkill /F /T /IM NT_DL.exe /IM NT_DL_app.exe >nul 2>&1
>>"%LOG_FILE%" echo taskkill_exit=%errorlevel%

set "WAIT_OK=0"
for /L %%I in (1,1,12) do (
    tasklist /FI "IMAGENAME eq NT_DL.exe" 2>nul | find /I "NT_DL.exe" >nul
    if errorlevel 1 (
        set "WAIT_OK=1"
        goto wait_done
    )
    ping 127.0.0.1 -n 2 >nul
)

:wait_done
>>"%LOG_FILE%" echo wait_ok=%WAIT_OK%

if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%" >nul 2>&1
    if errorlevel 1 goto error
)
>>"%LOG_FILE%" echo InstallDir=%INSTALL_DIR%

if not exist "%SOURCE_ROOT%" (
    >>"%LOG_FILE%" echo Missing payload folder: %SOURCE_ROOT%
    set "EXIT_CODE=2"
    goto error
)
if not exist "%SOURCE_ROOT%\NT_DL.exe" (
    >>"%LOG_FILE%" echo Missing payload exe: %SOURCE_ROOT%\NT_DL.exe
    set "EXIT_CODE=3"
    goto error
)
if not exist "%~dp0uninstall.cmd" (
    >>"%LOG_FILE%" echo Missing payload: %~dp0uninstall.cmd
    set "EXIT_CODE=4"
    goto error
)

robocopy "%SOURCE_ROOT%" "%INSTALL_DIR%" /E /R:1 /W:1 /NFL /NDL /NJH /NJS /NC /NS >nul
set "ROBOCOPY_EXIT=%errorlevel%"
>>"%LOG_FILE%" echo robocopy_exit=%ROBOCOPY_EXIT%
if %ROBOCOPY_EXIT% GEQ 8 (
    set "EXIT_CODE=5"
    goto error
)

copy /Y "%~dp0uninstall.cmd" "%INSTALL_DIR%\uninstall.cmd" >nul
if errorlevel 1 (
    set "EXIT_CODE=6"
    goto error
)
>>"%LOG_FILE%" echo Copied app folder and uninstall.cmd

if not exist "%START_MENU_DIR%" mkdir "%START_MENU_DIR%" >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut('%START_MENU_LINK%'); $s.TargetPath='%INSTALLED_EXE%'; $s.WorkingDirectory='%INSTALL_DIR%'; if (Test-Path '%INSTALLED_EXE%') { $s.IconLocation='%INSTALLED_EXE%' }; $s.Save()" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut('%DESKTOP_LINK%'); $s.TargetPath='%INSTALLED_EXE%'; $s.WorkingDirectory='%INSTALL_DIR%'; if (Test-Path '%INSTALLED_EXE%') { $s.IconLocation='%INSTALLED_EXE%' }; $s.Save()" >nul 2>&1

reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NT_DL" /v DisplayName /t REG_SZ /d "%APP_DISPLAY%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NT_DL" /v DisplayVersion /t REG_SZ /d "%APP_VERSION%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NT_DL" /v Publisher /t REG_SZ /d "%PUBLISHER%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NT_DL" /v InstallLocation /t REG_SZ /d "%INSTALL_DIR%" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NT_DL" /v UninstallString /t REG_SZ /d "\"%INSTALL_DIR%\uninstall.cmd\"" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NT_DL" /v QuietUninstallString /t REG_SZ /d "\"%INSTALL_DIR%\uninstall.cmd\" /quiet" /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NT_DL" /v NoModify /t REG_DWORD /d 1 /f >nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\NT_DL" /v NoRepair /t REG_DWORD /d 1 /f >nul 2>&1
>>"%LOG_FILE%" echo Registry entries updated

if "%QUIET%"=="1" (
    >>"%LOG_FILE%" echo Install success ^(quiet^)
    exit /b 0
)

echo.
echo NT_DL installed successfully.
echo Location: %INSTALL_DIR%
>>"%LOG_FILE%" echo Install success (interactive)
start "" "%INSTALLED_EXE%"
exit /b 0

:error
>>"%LOG_FILE%" echo Install failed with exit_code %EXIT_CODE%
if "%QUIET%"=="1" exit /b %EXIT_CODE%
echo.
echo NT_DL installation failed.
echo Please run setup again. Log: %LOG_FILE%
exit /b %EXIT_CODE%
