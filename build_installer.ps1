param(
    [string]$Version = ""
)

$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
if (-not $root) {
    $root = (Get-Location).Path
}

$launcherExePath = Join-Path $root "dist\NT_DL.exe"
$appExePath = Join-Path $root "dist\NT_DL_app.exe"
if (-not (Test-Path $launcherExePath)) {
    throw "Missing dist\NT_DL.exe. Build the launcher first."
}
if (-not (Test-Path $appExePath)) {
    throw "Missing dist\NT_DL_app.exe. Build the app first."
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    $initPath = Join-Path $root "kdl\__init__.py"
    $initText = Get-Content $initPath -Raw
    $m = [regex]::Match($initText, '__version__\s*=\s*"([^"]+)"')
    if (-not $m.Success) {
        throw "Could not read __version__ from kdl\__init__.py"
    }
    $Version = $m.Groups[1].Value
}

$payloadDir = Join-Path $env:TEMP "kdl_installer_payload"
if (Test-Path $payloadDir) {
    Remove-Item -Path $payloadDir -Recurse -Force
}
New-Item -ItemType Directory -Path $payloadDir | Out-Null

Copy-Item -Path $launcherExePath -Destination (Join-Path $payloadDir "NT_DL.exe") -Force
Copy-Item -Path $appExePath -Destination (Join-Path $payloadDir "NT_DL_app.exe") -Force
Copy-Item -Path (Join-Path $root "installer\install.cmd") -Destination (Join-Path $payloadDir "install.cmd") -Force
Copy-Item -Path (Join-Path $root "installer\uninstall.cmd") -Destination (Join-Path $payloadDir "uninstall.cmd") -Force
Copy-Item -Path (Join-Path $root "kdl\assets\kdl_a.ico") -Destination (Join-Path $payloadDir "kdl_a.ico") -Force

$installCmdPath = Join-Path $payloadDir "install.cmd"
$installCmd = Get-Content -Path $installCmdPath -Raw
$installCmd = [regex]::Replace($installCmd, 'set "APP_VERSION=.*"', ('set "APP_VERSION=' + $Version + '"'))
Set-Content -Path $installCmdPath -Value $installCmd -Encoding ASCII

$launcherScript = Join-Path $root "installer\iexpress_launcher.py"
if (-not (Test-Path $launcherScript)) {
    throw "Missing installer\iexpress_launcher.py"
}

$launcherBuild = Join-Path $env:TEMP "nt_dl_launcher_build"
if (Test-Path $launcherBuild) {
    Remove-Item -Path $launcherBuild -Recurse -Force
}
New-Item -ItemType Directory -Path $launcherBuild | Out-Null

$launcherDist = Join-Path $launcherBuild "dist"
$launcherWork = Join-Path $launcherBuild "build"
$launcherSpec = Join-Path $launcherBuild "spec"
$launcherName = "NT_DL_InstallerLauncher"

$pyArgs = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--onefile",
    "--console",
    "--name", $launcherName,
    "--distpath", $launcherDist,
    "--workpath", $launcherWork,
    "--specpath", $launcherSpec,
    $launcherScript
)

& python @pyArgs
if ($LASTEXITCODE -ne 0) {
    throw "Launcher build failed with exit code $LASTEXITCODE"
}

$launcherExe = Join-Path $launcherDist ($launcherName + ".exe")
if (-not (Test-Path $launcherExe)) {
    throw "Launcher build failed: executable not produced."
}
Copy-Item -Path $launcherExe -Destination (Join-Path $payloadDir "launcher.exe") -Force

$tempTarget = Join-Path $env:TEMP ("NT_DL-Setup-" + $Version + ".exe")
if (Test-Path $tempTarget) {
    Remove-Item -Path $tempTarget -Force
}

$sedPath = Join-Path $env:TEMP "NT_DL_Setup.sed"
$sed = @"
[Version]
Class=IEXPRESS
SEDVersion=3

[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=0
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=1
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=NT_DL installation completed.
TargetName=$tempTarget
FriendlyName=NT_DL Setup
AppLaunched=launcher.exe
PostInstallCmd=<None>
AdminQuietInstCmd=launcher.exe
UserQuietInstCmd=launcher.exe
SourceFiles=SourceFiles

[SourceFiles]
SourceFiles0=$payloadDir

[SourceFiles0]
%FILE0%=
%FILE1%=
%FILE2%=
%FILE3%=
%FILE4%=
%FILE5%=

[Strings]
FILE0=launcher.exe
FILE1=install.cmd
FILE2=uninstall.cmd
FILE3=NT_DL.exe
FILE4=NT_DL_app.exe
FILE5=kdl_a.ico
"@

Set-Content -Path $sedPath -Value $sed -Encoding ASCII

$iexpress = Join-Path $env:WINDIR "System32\iexpress.exe"
if (-not (Test-Path $iexpress)) {
    throw "IExpress not found at $iexpress"
}

& $iexpress /N /Q $sedPath | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "IExpress failed with exit code $LASTEXITCODE"
}

if (-not (Test-Path $tempTarget)) {
    throw "Installer build failed: setup executable not produced."
}

$finalTarget = Join-Path $root ("dist\NT_DL-Setup-" + $Version + ".exe")
Copy-Item -Path $tempTarget -Destination $finalTarget -Force

Write-Output "Installer created: $finalTarget"
