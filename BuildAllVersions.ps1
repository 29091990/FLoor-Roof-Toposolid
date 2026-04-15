# BuildAllVersions.ps1 - Multi-version build for FloorRoofTopo
# Builds for Revit 2020 through 2026 using dotnet build

$ErrorActionPreference = "Stop"
$projectDir = $PSScriptRoot
$csproj = Join-Path $projectDir "FloorRoofTopo.csproj"

# Target versions
$versions = @("2020", "2021", "2022", "2023", "2024", "2025", "2026")

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  FloorRoofTopo - Multi-Version Build"        -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$results = @()

foreach ($version in $versions) {
    $intVer = [int]$version
    $revitApiPath = "C:\Program Files\Autodesk\Revit $version\RevitAPI.dll"

    # Check if Revit API exists for this version
    if (-not (Test-Path $revitApiPath)) {
        Write-Host "[$version] SKIP - RevitAPI.dll not found" -ForegroundColor Yellow
        $results += [PSCustomObject]@{ Version = $version; Status = "SKIPPED"; Reason = "API not found" }
        continue
    }

    Write-Host "[$version] Building..." -ForegroundColor White -NoNewline

    # Clean obj folder to avoid cross-version conflicts
    $objDir = Join-Path $projectDir "obj"
    if (Test-Path $objDir) {
        Remove-Item $objDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    $success = $false
    $outputDir = Join-Path $projectDir "bin\Revit$version"

    try {
        $buildOutput = dotnet build $csproj -p:RevitVersion=$version -c Release --nologo -v q 2>&1
        $success = ($LASTEXITCODE -eq 0)
    }
    catch {
        $success = $false
    }

    $dllPath = Join-Path $outputDir "FloorRoofTopo.dll"
    if ($success -and (Test-Path $dllPath)) {
        $size = [math]::Round((Get-Item $dllPath).Length / 1KB, 1)
        Write-Host " OK ($size KB)" -ForegroundColor Green
        $results += [PSCustomObject]@{ Version = $version; Status = "OK"; Reason = "$size KB" }
    }
    else {
        Write-Host " FAILED" -ForegroundColor Red
        if ($buildOutput) {
            $errorLines = $buildOutput | Where-Object { $_ -match "error " } | Select-Object -First 3
            foreach ($line in $errorLines) {
                Write-Host "  $line" -ForegroundColor DarkRed
            }
        }
        $results += [PSCustomObject]@{ Version = $version; Status = "FAILED"; Reason = "Build error" }
    }
}

# Summary
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Build Summary" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
$results | Format-Table -AutoSize
