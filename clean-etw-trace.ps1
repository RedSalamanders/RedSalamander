# Clean ETW Trace Files
# This script removes old trace files and ensures no stale sessions exist

$SessionName = "RedSalamanderTrace"
$FilesToClean = @(
    "RedSalamanderTrace.etl",
    "RedSalamanderTrace.txt",
    "RedSalamanderTrace.csv"
)

Write-Host "=== Cleaning ETW Trace Files ===" -ForegroundColor Cyan
Write-Host ""

# Stop any existing session
Write-Host "Stopping any existing ETW session..." -ForegroundColor Gray
$result = logman stop $SessionName -ets 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "OK: Stopped existing session" -ForegroundColor Green
} else {
    Write-Host "  No active session found" -ForegroundColor DarkGray
}
Write-Host ""

# Delete old trace files
$filesDeleted = 0
foreach ($file in $FilesToClean) {
    if (Test-Path $file) {
        Write-Host "Deleting $file..." -ForegroundColor Gray
        Remove-Item $file -Force
        $filesDeleted++
    }
}

if ($filesDeleted -gt 0) {
    Write-Host ""
    Write-Host "OK: Deleted $filesDeleted old trace file(s)" -ForegroundColor Green
} else {
    Write-Host "  No old trace files found" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "OK: Cleanup complete! Ready for a fresh trace." -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Run: .\start-etw-trace.ps1" -ForegroundColor White
Write-Host "  2. Run: .\MonitorTest.exe" -ForegroundColor White
Write-Host "  3. Run: .\stop-etw-trace.ps1" -ForegroundColor White
Write-Host ""
