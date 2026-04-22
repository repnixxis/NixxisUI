# Nixxis Maintenance Tool — Remote Launcher
# Usage: irm "https://raw.githubusercontent.com/YOUR_ORG/NixxisUI/main/launch.ps1" | iex
#
# Replace YOUR_ORG with the actual GitHub username/org before publishing.

$repoRaw = "https://raw.githubusercontent.com/YOUR_ORG/NixxisUI/main/NixxisUI.ps1"

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
    Write-Host "Check that the GitHub repo is public and the URL is correct." -ForegroundColor Yellow
}
