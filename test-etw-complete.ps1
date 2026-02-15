# Complete ETW Test Script
# This script performs a full end-to-end test of the ETW infrastructure

param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",
    [switch]$NoPause
)

Write-Host "================================" -ForegroundColor Cyan
Write-Host "Red Salamander ETW Complete Test" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Check privileges
Write-Host "[1/6] Checking privileges..." -ForegroundColor Yellow
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    Write-Host "  OK: Running as Administrator" -ForegroundColor Green
} else {
    Write-Host "  WARNING: NOT running as Administrator" -ForegroundColor Yellow
    Write-Host "     ETW trace sessions require admin privileges" -ForegroundColor Gray
    Write-Host "     Re-run this script as Administrator for best results" -ForegroundColor Gray
    if (-not $NoPause) {
        Write-Host ""
        Write-Host "Press any key to continue anyway or Ctrl+C to exit..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}
Write-Host ""

# Step 2: Clean old files
Write-Host "[2/6] Cleaning old trace files..." -ForegroundColor Yellow
.\clean-etw-trace.ps1
Write-Host ""

# Step 3: Start optimized trace
Write-Host "[3/6] Starting optimized ETW trace session..." -ForegroundColor Yellow
.\start-etw-trace.ps1 -NoPause:$NoPause
Write-Host ""

# Step 4: Verify MonitorTest exists
Write-Host "[4/6] Checking for MonitorTest.exe..." -ForegroundColor Yellow
$monitorTestExe = Join-Path "x64" (Join-Path $Configuration "MonitorTest.exe")
if (Test-Path $monitorTestExe) {
    $fileInfo = Get-Item $monitorTestExe
    Write-Host "  OK: Found: $($fileInfo.FullName)" -ForegroundColor Green
    Write-Host "    Size: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Gray
    Write-Host "    Modified: $($fileInfo.LastWriteTime)" -ForegroundColor Gray
} else {
    Write-Host "  ERROR: MonitorTest.exe not found!" -ForegroundColor Red
    Write-Host "    Please build the solution first" -ForegroundColor Yellow
    if (-not $NoPause) {
        Write-Host ""
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    exit 1
}
Write-Host ""

# Step 5: Run MonitorTest
Write-Host "[5/6] Running MonitorTest to generate ETW events..." -ForegroundColor Yellow
Write-Host "  This will generate ~150,000 events over ~10 seconds..." -ForegroundColor Gray
Write-Host ""
& $monitorTestExe
Write-Host ""
Write-Host "  OK: MonitorTest completed" -ForegroundColor Green
Write-Host ""

# Step 6: Stop trace and analyze
Write-Host "[6/6] Stopping trace and analyzing results..." -ForegroundColor Yellow
.\stop-etw-trace.ps1 -NoPause:$NoPause
Write-Host ""

# Summary
Write-Host "================================" -ForegroundColor Cyan
Write-Host "Test Complete!" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Check RedSalamanderTrace.csv for statistics" -ForegroundColor Gray
Write-Host "  2. Open RedSalamanderTrace.etl in Windows Performance Analyzer" -ForegroundColor Gray
Write-Host "  3. Review RedSalamanderTrace.txt for human-readable events" -ForegroundColor Gray
Write-Host ""
Write-Host "Expected result: 0 events lost with optimized buffer configuration" -ForegroundColor Green
Write-Host ""
