# Nixxis Maintenance Tool — Remote Launcher
# Usage (PowerShell 5.1 / Windows Server — forces TLS 1.2):
#   [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; irm "https://raw.githubusercontent.com/repnixxis/NixxisUI/main/launch.ps1" | iex

# Force TLS 1.2 — PowerShell 5.1 defaults to TLS 1.0, which GitHub rejects
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$repoRaw = "https://raw.githubusercontent.com/repnixxis/NixxisUI/main/NixxisUI.ps1"

Write-Host ""
Write-Host "  Nixxis Maintenance Tool" -ForegroundColor Cyan
Write-Host "  Loading from: $repoRaw" -ForegroundColor Gray
Write-Host ""

try {
    $script = Invoke-RestMethod -Uri $repoRaw -UseBasicParsing
    Invoke-Expression $script
} catch {
    Write-Host "ERROR: Failed to load NixxisUI.ps1" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "If you see an SSL/TLS error, run this first:" -ForegroundColor Yellow
    Write-Host "  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12" -ForegroundColor Cyan
}
