# Gera SanitizadorRTF.exe (ficheiro único, sem consola)
# Requer: pip install -r requirements.txt
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$venvPy = Join-Path $PSScriptRoot ".gui-venv\Scripts\python.exe"
if (-not (Test-Path $venvPy)) {
    python -m venv .gui-venv
    $venvPy = Join-Path $PSScriptRoot ".gui-venv\Scripts\python.exe"
}
& $venvPy -m pip install -q -r requirements.txt
& $venvPy -m PyInstaller --noconfirm --clean `
    --onefile `
    --windowed `
    --name "SanitizadorRTF" `
    rtf_sanitize_gui.py

Write-Host "Executável: $PSScriptRoot\dist\SanitizadorRTF.exe"
