param(
    [string]$Version = "",
    [string]$AppDir = "",
    [ValidateSet("Full", "Standard")]
    [string]$Edition = "Full"
)

$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
if (-not $root) {
    $root = (Get-Location).Path
}

# ── Resolve app directory ──────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($AppDir)) {
    $appDirPath = Join-Path $root "dist\NT_DL_$Edition"
} else {
    $appDirPath = $AppDir
}
if (-not (Test-Path $appDirPath)) {
    throw "Missing app directory: $appDirPath"
}
$appExeName = "NT_DL_$Edition.exe"
$appExePath = Join-Path $appDirPath $appExeName
if (-not (Test-Path $appExePath)) {
    throw "Missing $appExeName in app directory: $appDirPath"
}

if ($Edition -eq "Full") {
    $appName = "NT DL Full"
    $installDir = "NT_DL_Full"
    $appId = "{{A3F2C1D4-8B6E-4F9A-BC12-5E7D0A3F9C21}"
} else {
    $appName = "NT DL Standard"
    $installDir = "NT_DL_Standard"
    $appId = "{{4F54126B-643A-48D1-A281-7C3749DA229B}"
}

# ── Read version from kdl\__init__.py ─────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($Version)) {
    $initPath = Join-Path $root "kdl\__init__.py"
    $initText = Get-Content $initPath -Raw
    $m = [regex]::Match($initText, '__version__\s*=\s*"([^"]+)"')
    if (-not $m.Success) {
        throw "Could not read __version__ from kdl\__init__.py"
    }
    $Version = $m.Groups[1].Value
}
$outputBase = "NT_DL-$Edition-Setup-$Version"

# ── Locate Inno Setup compiler (ISCC.exe) ─────────────────────────────────────
$isccPaths = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 5\ISCC.exe",
    "C:\Program Files\Inno Setup 5\ISCC.exe"
)
$iscc = $null
foreach ($p in $isccPaths) {
    if (Test-Path $p) { $iscc = $p; break }
}

# Auto-download and silently install Inno Setup if not found
if (-not $iscc) {
    Write-Output "Inno Setup not found - downloading installer..."
    $isSetupExe = Join-Path $env:TEMP "innosetup_install.exe"
    $isUrl = "https://jrsoftware.org/download.php/is.exe"
    try {
        Invoke-WebRequest -Uri $isUrl -OutFile $isSetupExe -UseBasicParsing
    } catch {
        throw "Could not download Inno Setup from $isUrl : $_"
    }
    Write-Output "Installing Inno Setup silently..."
    Start-Process -FilePath $isSetupExe -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" -Wait
    Remove-Item $isSetupExe -Force -ErrorAction SilentlyContinue

    foreach ($p in $isccPaths) {
        if (Test-Path $p) { $iscc = $p; break }
    }
    if (-not $iscc) {
        throw "Inno Setup installation completed but ISCC.exe was not found. Please install Inno Setup manually from https://jrsoftware.org/isinfo.php"
    }
    Write-Output "Inno Setup installed: $iscc"
}

# ── Build the .iss script from template ───────────────────────────────────────
$templatePath = Join-Path $root "installer\NT_DL.iss.template"
if (-not (Test-Path $templatePath)) {
    throw "Missing installer template: $templatePath"
}

$iconFile = Join-Path $root "kdl\assets\kdl_a.ico"
$outputDir = Join-Path $root "dist"
$issContent = Get-Content $templatePath -Raw
$issContent = $issContent -replace "@@VERSION@@",  $Version
$issContent = $issContent -replace "@@APPDIR@@",   $appDirPath.TrimEnd('\')
$issContent = $issContent -replace "@@ICONFILE@@",  $iconFile
$issContent = $issContent -replace "@@OUTPUTDIR@@", $outputDir
$issContent = $issContent -replace "@@APPNAME@@", $appName
$issContent = $issContent -replace "@@EXENAME@@", $appExeName
$issContent = $issContent -replace "@@INSTALLDIR@@", $installDir
$issContent = $issContent -replace "@@APPID@@", $appId
$issContent = $issContent -replace "@@OUTPUTBASE@@", $outputBase

$issFile = Join-Path $env:TEMP "NT_DL_${Edition}_$Version.iss"
Set-Content -Path $issFile -Value $issContent -Encoding UTF8

# ── Compile with Inno Setup ────────────────────────────────────────────────────
Write-Output "Compiling installer with Inno Setup..."
& $iscc $issFile
if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup compilation failed (exit $LASTEXITCODE)"
}
Remove-Item $issFile -Force -ErrorAction SilentlyContinue

# ── Verify output ──────────────────────────────────────────────────────────────
$finalExe = Join-Path $outputDir "$outputBase.exe"
if (-not (Test-Path $finalExe)) {
    throw "Installer build completed but output EXE not found: $finalExe"
}

# ── Clean up old setup artifacts (folders and zips from old bootstrap approach) ─
$staleItems = Get-ChildItem -Path $outputDir -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "NT_DL-$Edition-Setup-*" -and $_.FullName -ne $finalExe }
foreach ($item in $staleItems) {
    if ($item.PSIsContainer) {
        Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        Remove-Item -LiteralPath $item.FullName -Force -ErrorAction SilentlyContinue
    }
}

# ── Clean up legacy release files ─────────────────────────────────────────────
Write-Output "Installer created: $finalExe"
