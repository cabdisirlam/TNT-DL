@echo off
setlocal enableextensions

set "QUIET=0"
if /I "%~1"=="/quiet" set "QUIET=1"

set "APP_NAME=NT_DL"
set "APP_DISPLAY=NT_DL"
set "APP_VERSION=1.1.72"
set "PUBLISHER=NT_DL"
set "LOG_FILE=%TEMP%\NT_DL_install.log"

>"%LOG_FILE%" echo [%date% %time%] NT_DL install start
>>"%LOG_FILE%" echo SourceDir=%~dp0

set "INSTALL_DIR=%LOCALAPPDATA%\Programs\NT_DL"
set "START_MENU_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\NT_DL"
set "START_MENU_LINK=%START_MENU_DIR%\NT_DL.lnk"
set "DESKTOP_LINK=%USERPROFILE%\Desktop\NT_DL.lnk"

rem Ensure old app instance does not lock NT_DL.exe during upgrade.
taskkill /F /IM NT_DL.exe >nul 2>&1

if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%" >nul 2>&1
    if errorlevel 1 goto error
)
>>"%LOG_FILE%" echo InstallDir=%INSTALL_DIR%

set "COPY_OK=0"
for /L %%I in (1,1,3) do (
    del /F /Q "%INSTALL_DIR%\NT_DL.exe" >nul 2>&1
    copy /Y "%~dp0NT_DL.exe" "%INSTALL_DIR%\NT_DL.exe" >nul
    if not errorlevel 1 (
        >>"%LOG_FILE%" echo Copied NT_DL.exe on attempt %%I
        set "COPY_OK=1"
        goto copy_done
    )
    >>"%LOG_FILE%" echo Copy attempt %%I failed
    timeout /t 1 /nobreak >nul
)

:copy_done
if not "%COPY_OK%"=="1" goto error

if exist "%~dp0kdl_a.ico" copy /Y "%~dp0kdl_a.ico" "%INSTALL_DIR%\kdl_a.ico" >nul
copy /Y "%~dp0uninstall.cmd" "%INSTALL_DIR%\uninstall.cmd" >nul
if errorlevel 1 goto error
>>"%LOG_FILE%" echo Copied uninstall.cmd and icon

if not exist "%START_MENU_DIR%" mkdir "%START_MENU_DIR%" >nul 2>&1

powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut('%START_MENU_LINK%'); $s.TargetPath='%INSTALL_DIR%\NT_DL.exe'; $s.WorkingDirectory='%INSTALL_DIR%'; if (Test-Path '%INSTALL_DIR%\kdl_a.ico') { $s.IconLocation='%INSTALL_DIR%\kdl_a.ico' }; $s.Save()" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $s=$w.CreateShortcut('%DESKTOP_LINK%'); $s.TargetPath='%INSTALL_DIR%\NT_DL.exe'; $s.WorkingDirectory='%INSTALL_DIR%'; if (Test-Path '%INSTALL_DIR%\kdl_a.ico') { $s.IconLocation='%INSTALL_DIR%\kdl_a.ico' }; $s.Save()" >nul 2>&1

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
start "" "%INSTALL_DIR%\NT_DL.exe"
exit /b 0

:error
>>"%LOG_FILE%" echo Install failed with errorlevel %errorlevel%
if "%QUIET%"=="1" exit /b 1
echo.
echo NT_DL installation failed.
echo Please run setup again. Log: %LOG_FILE%
exit /b 1
