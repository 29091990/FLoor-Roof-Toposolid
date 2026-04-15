# Deploy.ps1 - Deploy FloorRoofTopo to Revit Addins folder
# Copies DLL + .addin to each version's addins directory

param(
    [string[]]$Versions = @("2024", "2025", "2026")
)

$ErrorActionPreference = "Continue"
$projectDir = $PSScriptRoot
$addinFile = Join-Path $projectDir "FloorRoofTopo.addin"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  FloorRoofTopo - Deploy"                     -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

foreach ($version in $Versions) {
    $dllSource = Join-Path $projectDir "bin\Revit$version\FloorRoofTopo.dll"
    $addinDest = "$env:APPDATA\Autodesk\Revit\Addins\$version"

    Write-Host "[$version] " -NoNewline

    if (-not (Test-Path $dllSource)) {
        Write-Host "SKIP - DLL not found (bin\Revit$version\)" -ForegroundColor Yellow
        continue
    }

    if (-not (Test-Path $addinDest)) {
        New-Item -ItemType Directory -Path $addinDest -Force | Out-Null
    }

    try {
        Copy-Item $dllSource -Destination $addinDest -Force -ErrorAction Stop
        Copy-Item $addinFile -Destination $addinDest -Force -ErrorAction Stop
        $dllInfo = Get-Item (Join-Path $addinDest "FloorRoofTopo.dll")
        $size = [math]::Round($dllInfo.Length / 1KB, 1)
        Write-Host "OK - Deployed $size KB to $addinDest" -ForegroundColor Green
    }
    catch {
        Write-Host "FAILED - $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  (Close Revit if file is locked)" -ForegroundColor DarkYellow
    }
}

Write-Host ""
Write-Host "Done! Restart Revit to load the updated add-in." -ForegroundColor Cyan
