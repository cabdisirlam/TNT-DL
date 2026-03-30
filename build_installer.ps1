param(
    [string]$Version = "",
    [string]$AppDir = ""
)

$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
if (-not $root) {
    $root = (Get-Location).Path
}

if ([string]::IsNullOrWhiteSpace($AppDir)) {
    $appDirPath = Join-Path $root "dist\NT_DL"
} else {
    $appDirPath = $AppDir
}
if (-not (Test-Path $appDirPath)) {
    throw "Missing app directory: $appDirPath"
}
$appExePath = Join-Path $appDirPath "NT_DL.exe"
if (-not (Test-Path $appExePath)) {
    throw "Missing NT_DL.exe in app directory: $appDirPath"
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

$packageZipPath = Join-Path $payloadDir "NT_DL_package.zip"
Compress-Archive -Path (Join-Path $appDirPath "*") -DestinationPath $packageZipPath -Force
Copy-Item -Path (Join-Path $root "installer\install.cmd") -Destination (Join-Path $payloadDir "install.cmd") -Force
Copy-Item -Path (Join-Path $root "installer\uninstall.cmd") -Destination (Join-Path $payloadDir "uninstall.cmd") -Force

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
    "--onedir",
    "--windowed",
    "--name", $installerName,
    "--icon", (Join-Path $root "kdl\assets\kdl_a.ico"),
    "--distpath", $installerDist,
    "--workpath", $installerWork,
    "--specpath", $installerSpec,
    $bootstrapScript
)

& python @pyArgs
if ($LASTEXITCODE -ne 0) {
    throw "Installer build failed with exit code $LASTEXITCODE"
}

$installerDir = Join-Path $installerDist $installerName
$installerExe = Join-Path $installerDir ($installerName + ".exe")
if (-not (Test-Path $installerExe)) {
    throw "Installer build failed: executable not produced."
}

Copy-Item -Path $packageZipPath -Destination (Join-Path $installerDir "NT_DL_package.zip") -Force
Copy-Item -Path (Join-Path $payloadDir "install.cmd") -Destination (Join-Path $installerDir "install.cmd") -Force
Copy-Item -Path (Join-Path $payloadDir "uninstall.cmd") -Destination (Join-Path $installerDir "uninstall.cmd") -Force

$finalDir = Join-Path $root ("dist\" + $installerName)
if (Test-Path $finalDir) {
    Remove-Item -LiteralPath $finalDir -Recurse -Force
}
Copy-Item -Path $installerDir -Destination $finalDir -Recurse -Force

$finalZip = Join-Path $root ("dist\" + $installerName + ".zip")
if (Test-Path $finalZip) {
    Remove-Item -LiteralPath $finalZip -Force
}
Compress-Archive -LiteralPath $finalDir -DestinationPath $finalZip -Force

$staleSetupArtifacts = Get-ChildItem -Path (Join-Path $root "dist") -Filter "NT_DL-Setup-*" -ErrorAction SilentlyContinue
foreach ($staleArtifact in $staleSetupArtifacts) {
    if ($staleArtifact.FullName -ne $finalDir -and $staleArtifact.FullName -ne $finalZip) {
        if ($staleArtifact.PSIsContainer) {
            Remove-Item -LiteralPath $staleArtifact.FullName -Recurse -Force
        } else {
            Remove-Item -LiteralPath $staleArtifact.FullName -Force
        }
    }
}

$legacyReleaseFiles = @(
    (Join-Path $root "dist\NT_DL.exe"),
    (Join-Path $root "dist\NT_DL_app.exe")
)
foreach ($legacyPath in $legacyReleaseFiles) {
    if (Test-Path $legacyPath) {
        Remove-Item -LiteralPath $legacyPath -Force
    }
}

Write-Output "Installer directory created: $finalDir"
Write-Output "Installer zip created: $finalZip"
