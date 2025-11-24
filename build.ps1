Write-Host ================================
Write-Host  Plategen Build Script 
Write-Host ================================

# -------------------------------------------------------------
# Switch to script directory
# -------------------------------------------------------------
Set-Location -Path $PSScriptRoot
Write-Host Working Directory $PWD

# -------------------------------------------------------------
# Check  Create virtual environment
# -------------------------------------------------------------
if (!(Test-Path ".\.venv")) {
    Write-Host Creating virtual environment...
    python -m venv .venv
} else {
    Write-Host Virtual environment already exists.
}

# -------------------------------------------------------------
# Activate virtual environment
# -------------------------------------------------------------
Write-Host Activating virtual environment...
$activateScript = ".\.venv\Scripts\Activate.ps1"
if (!(Test-Path $activateScript)) {
    Write-Host ERROR Activation script not found!
    exit 1
}
. $activateScript

# -------------------------------------------------------------
# Install dependencies
# -------------------------------------------------------------
Write-Host Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt

# -------------------------------------------------------------
# Build EXE
# -------------------------------------------------------------
Write-Host Building EXE with PyInstaller...
pyinstaller --noconfirm --onefile --windowed --icon=installer/icons/plategen_icon.ico --name=plategen app.py --add-data "installer/icons/plategen_icon.ico;installer/icons" --collect-all requests
PyInstaller --noconfirm --onefile --windowed --icon=installer/icons/plategen_icon.ico --name=app_db app_db.py --add-data "installer/icons/plategen_icon.ico;installer/icons" --collect-all requests
pyinstaller --noconfirm --onefile --windowed --icon=installer/icons/plategen_icon.ico --name=app_ups app_ups.py --add-data "installer/icons/plategen_icon.ico;installer/icons" --collect-all requests
pyinstaller --noconfirm --onefile --windowed --icon=installer/icons/plategen_icon.ico --name=app_bch app_bch.py --add-data "installer/icons/plategen_icon.ico;installer/icons" --collect-all requests
pyinstaller --noconfirm --onefile --windowed --icon=installer/icons/plategen_icon.ico --name=app_np app_np.py --add-data "installer/icons/plategen_icon.ico;installer/icons" --collect-all requests
pyinstaller --noconfirm --onefile --windowed --icon=installer/icons/plategen_icon.ico --name=app_np_db_schema app_np_db_schema.py --add-data "installer/icons/plategen_icon.ico;installer/icons" --collect-all requests
pyinstaller --noconfirm --onefile --windowed --icon=installer/icons/sticker_icon.ico --name=app_sticker app_sticker.py --add-data "installer/icons/sticker_icon.ico;installer/icons" --collect-all requests
pyinstaller --noconfirm --onefile --windowed --icon=installer/icons/manual_icon.ico --name=app_mgen_ups app_mgen_ups.py --add-data "installer/icons/manual_icon.ico;installer/icons" --collect-all requests
pyinstaller --noconfirm --onefile --windowed --icon=installer/icons/manual_icon.ico --name=app_mgen_bch app_mgen_bch.py --add-data "installer/icons/manual_icon.ico;installer/icons" --collect-all requests

if ($LASTEXITCODE -ne 0) {
    Write-Host PyInstaller build FAILED!
    exit 1
}

Write-Host Executable built distapp.exe

# -------------------------------------------------------------
# Build installer with Inno Setup
# -------------------------------------------------------------
Write-Host Building installer with Inno Setup...

$ISS_PATH = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

if (!(Test-Path $ISS_PATH)) {
    Write-Host ERROR ISCC.exe NOT FOUND at
    Write-Host   $ISS_PATH
    Write-Host Install from httpsjrsoftware.orgisdl.php
    exit 1
}

Write-Host Using Inno Setup at
Write-Host   $ISS_PATH

& $ISS_PATH installeriscript.iss

if ($LASTEXITCODE -ne 0) {
    Write-Host Installer build FAILED!
    exit 1
}

Write-Host 
Write-Host =======================================
Write-Host           BUILD SUCCESSFUL!
Write-Host Installer located at installeroutput
Write-Host =======================================