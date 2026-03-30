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

Copy-Item -Path $launcherExePath -Destination (Join-Path $payloadDir "NT_DL_payload.dat") -Force
Copy-Item -Path $appExePath -Destination (Join-Path $payloadDir "NT_DL_app.exe") -Force
Copy-Item -Path (Join-Path $root "installer\install.cmd") -Destination (Join-Path $payloadDir "install.cmd") -Force
Copy-Item -Path (Join-Path $root "installer\uninstall.cmd") -Destination (Join-Path $payloadDir "uninstall.cmd") -Force
Copy-Item -Path (Join-Path $root "kdl\assets\kdl_a.ico") -Destination (Join-Path $payloadDir "kdl_a.ico") -Force

$installCmdPath = Join-Path $payloadDir "install.cmd"
$installCmd = Get-Content -Path $installCmdPath -Raw
$installCmd = [regex]::Replace($installCmd, 'set "APP_VERSION=.*"', ('set "APP_VERSION=' + $Version + '"'))
Set-Content -Path $installCmdPath -Value $installCmd -Encoding ASCII

$bootstrapScript = Join-Path $root "installer\setup_bootstrap.py"
if (-not (Test-Path $bootstrapScript)) {
    throw "Missing installer\setup_bootstrap.py"
}

$installerBuild = Join-Path $env:TEMP "nt_dl_setup_build"
if (Test-Path $installerBuild) {
    Remove-Item -Path $installerBuild -Recurse -Force
}
New-Item -ItemType Directory -Path $installerBuild | Out-Null

$installerDist = Join-Path $installerBuild "dist"
$installerWork = Join-Path $installerBuild "build"
$installerSpec = Join-Path $installerBuild "spec"
$installerName = "NT_DL-Setup-$Version"

$pyArgs = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--onefile",
    "--windowed",
    "--name", $installerName,
    "--icon", (Join-Path $root "kdl\assets\kdl_a.ico"),
    "--distpath", $installerDist,
    "--workpath", $installerWork,
    "--specpath", $installerSpec,
    "--add-data", ((Join-Path $payloadDir "NT_DL_payload.dat") + ";."),
    "--add-data", ((Join-Path $payloadDir "NT_DL_app.exe") + ";."),
    "--add-data", ((Join-Path $payloadDir "install.cmd") + ";."),
    "--add-data", ((Join-Path $payloadDir "uninstall.cmd") + ";."),
    "--add-data", ((Join-Path $payloadDir "kdl_a.ico") + ";."),
    $bootstrapScript
)

& python @pyArgs
if ($LASTEXITCODE -ne 0) {
    throw "Installer build failed with exit code $LASTEXITCODE"
}

$installerExe = Join-Path $installerDist ($installerName + ".exe")
if (-not (Test-Path $installerExe)) {
    throw "Installer build failed: executable not produced."
}
$finalTarget = Join-Path $root ("dist\NT_DL-Setup-" + $Version + ".exe")
Copy-Item -Path $installerExe -Destination $finalTarget -Force

Write-Output "Installer created: $finalTarget"
