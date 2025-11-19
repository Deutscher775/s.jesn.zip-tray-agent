<#
build_installer.ps1
PowerShell helper to create a reproducible build for the JesnZIP tray agent.

This script will:
- create (or reuse) a virtual environment in `.venv_build`
- install/upgrade pip, setuptools, wheel inside the venv
- install project requirements into the venv
- run PyInstaller from the venv to produce a single-file, windowed exe
- zip the produced executable into a timestamped archive

Using a venv avoids global site-package conflicts (e.g. obsolete stdlib backports) and
makes builds reproducible.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Output "Running build_installer.ps1 in: $PSScriptRoot"
Set-Location $PSScriptRoot

$venvDir = Join-Path $PSScriptRoot '.venv_build'
$script = 'JesnZIP-tray.py'
$icon = 'ICON.ico'

if (-not (Test-Path $venvDir)) {
    Write-Output "Creating virtual environment at $venvDir"
    & python -m venv $venvDir
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to create virtual environment (exit $LASTEXITCODE)"
        exit $LASTEXITCODE
    }
} else {
    Write-Output "Using existing virtual environment at $venvDir"
}

$venvPython = Join-Path $venvDir 'Scripts\python.exe'
if (-not (Test-Path $venvPython)) {
    Write-Error "Virtual environment python not found at $venvPython"
    exit 1
}

Write-Output "Upgrading pip/setuptools/wheel in venv"
& $venvPython -m pip install --upgrade pip setuptools wheel
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to upgrade pip in venv"; exit $LASTEXITCODE }

Write-Output "Installing requirements into venv"
& $venvPython -m pip install -r .\requirements.txt
if ($LASTEXITCODE -ne 0) { Write-Error "Failed to install requirements in venv"; exit $LASTEXITCODE }

# Known problematic backports — attempt to remove inside venv (shouldn't be present by default)
$blacklist = @('typing','pathlib')
foreach ($pkg in $blacklist) {
    try {
        & $venvPython -m pip show $pkg > $null 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Output "Found incompatible package '$pkg' in venv; uninstalling..."
            & $venvPython -m pip uninstall -y $pkg > $null 2>&1
            Write-Output "Uninstalled $pkg (if present)"
        } else {
            Write-Output "Package $pkg not found in venv"
        }
    } catch {
        Write-Warning "Check/uninstall for package $pkg raised an exception: $_"
    }
}

Write-Output "Running PyInstaller from venv"
if (Test-Path (Join-Path $venvDir 'Scripts\pyinstaller.exe')) {
    $pyinstallerCmd = Join-Path $venvDir 'Scripts\pyinstaller.exe'
    if (Test-Path $icon) {
        & $pyinstallerCmd --noconfirm --onefile --windowed --icon=$icon $script
    } else {
        & $pyinstallerCmd --noconfirm --onefile --windowed $script
    }
} else {
    if (Test-Path $icon) {
        & $venvPython -m PyInstaller --noconfirm --onefile --windowed --icon=$icon $script
    } else {
        & $venvPython -m PyInstaller --noconfirm --onefile --windowed $script
    }
}

if ($LASTEXITCODE -ne 0) {
    Write-Error "PyInstaller failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}

# Zip the resulting exe
$distDir = Join-Path $PSScriptRoot 'dist'
if (-not (Test-Path $distDir)) {
    Write-Error "dist directory not found; PyInstaller may have failed"
    exit 1
}

$exe = Get-ChildItem $distDir -Filter *.exe -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $exe) {
    Write-Error "No executable found in dist; build may have failed"
    exit 1
}

$timestamp = Get-Date -Format yyyyMMddHHmm
$zipName = Join-Path $PSScriptRoot ("JesnZIP-tray-$timestamp.zip")
if (Test-Path $zipName) { Remove-Item $zipName -Force }
Compress-Archive -Path $exe.FullName -DestinationPath $zipName -Force
Write-Output "Built and zipped: $zipName"
exit 0
