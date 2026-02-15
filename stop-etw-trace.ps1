# Stop ETW Trace Session and View Results
# This script stops the ETW trace session and converts the .etl file to readable format

param(
    [switch]$NoPause
)

$SessionName = "RedSalamanderTrace"
$EtlFile = "RedSalamanderTrace.etl"
$TxtFile = "RedSalamanderTrace.txt"
$CsvFile = "RedSalamanderTrace.csv"

Write-Host "=== Stopping ETW Trace Session ===" -ForegroundColor Cyan
Write-Host "Session Name: $SessionName" -ForegroundColor Gray
Write-Host ""

# Stop the trace session
Write-Host "Stopping trace session..." -ForegroundColor Gray
$result = logman stop $SessionName -ets 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host "OK: ETW trace session stopped successfully!" -ForegroundColor Green
    Write-Host ""
    
    # Check if ETL file was created
    if (Test-Path $EtlFile) {
        $fileSize = (Get-Item $EtlFile).Length / 1KB
        Write-Host "ETL file created: $EtlFile ($([math]::Round($fileSize, 2)) KB)" -ForegroundColor White
        Write-Host ""
        
        # Convert to text format using tracerpt
        Write-Host "Converting trace to readable format..." -ForegroundColor Gray
        tracerpt $EtlFile -o $TxtFile -of CSV -summary $CsvFile 2>&1 | Out-Null
        
        if (Test-Path $TxtFile) {
            Write-Host "OK: Trace converted successfully!" -ForegroundColor Green
            Write-Host ""
            Write-Host "Output files:" -ForegroundColor White
            Write-Host "  - $EtlFile (binary trace)" -ForegroundColor Gray
            Write-Host "  - $TxtFile (human-readable events)" -ForegroundColor Gray
            Write-Host "  - $CsvFile (summary statistics)" -ForegroundColor Gray
            Write-Host ""
            
            # Display buffer statistics if available in CSV
            if (Test-Path $CsvFile) {
                Write-Host "=== Buffer Statistics ===" -ForegroundColor Cyan
                $csvContent = Get-Content $CsvFile
                $buffersLine = $csvContent | Where-Object { $_ -match "Total Buffers Processed" }
                $eventsLine = $csvContent | Where-Object { $_ -match "Total Events.*Processed" }
                $lostLine = $csvContent | Where-Object { $_ -match "Total Events.*Lost" }
                
                if ($buffersLine) { Write-Host $buffersLine -ForegroundColor Gray }
                if ($eventsLine) { Write-Host $eventsLine -ForegroundColor Gray }
                if ($lostLine) { 
                    Write-Host $lostLine -ForegroundColor Yellow
                    # Extract lost event count and warn if significant
                    if ($lostLine -match "(\d+)") {
                        $lostCount = [int]$Matches[1]
                        if ($lostCount -gt 0) {
                            # Calculate loss ratio to determine if it's buffer exhaustion or provider failure
                            $processedCount = 0
                            if ($eventsLine -match "(\d+)") {
                                $processedCount = [int]$Matches[1]
                            }
                            $totalEvents = $processedCount + $lostCount
                            $lossRate = if ($totalEvents -gt 0) { ($lostCount / $totalEvents) * 100 } else { 0 }
                            
                            Write-Host ""
                            if ($processedCount -lt 10 -and $lostCount -gt 1000) {
                                # Very few events captured suggests provider registration failure
                                Write-Host "WARNING: $lostCount events reported as 'lost'" -ForegroundColor Red
                                Write-Host "   However, only $processedCount events were captured." -ForegroundColor Yellow
                                Write-Host "   This suggests the application failed to emit ETW events," -ForegroundColor Yellow
                                Write-Host "   not actual buffer exhaustion." -ForegroundColor Yellow
                                Write-Host ""
                                Write-Host "   Possible causes:" -ForegroundColor White
                                Write-Host "   - ETW provider registration failed in the application" -ForegroundColor Gray
                                Write-Host "   - Application not running with correct privileges" -ForegroundColor Gray
                                Write-Host "   - Application using wrong provider GUID" -ForegroundColor Gray
                                Write-Host ""
                                Write-Host "   To verify: Check application console output for" -ForegroundColor Cyan
                                Write-Host "   'ETW Status: Registration failed' message" -ForegroundColor Cyan
                            }
                            else {
                                # Significant events captured suggests real buffer exhaustion
                                Write-Host "WARNING: $lostCount events were lost! (Loss rate: $([math]::Round($lossRate, 1))%)" -ForegroundColor Red
                                Write-Host "   This indicates genuine ETW buffer exhaustion." -ForegroundColor Yellow
                                Write-Host "   Consider increasing buffer settings in start-etw-trace.ps1:" -ForegroundColor Yellow
                                Write-Host "     -nb 64 -bs 512 -max 256  (for 32-128 MB capacity)" -ForegroundColor Gray
                                Write-Host "   Current settings: -nb 32 -bs 256 -max 512 -ft 1 (8-128 MB)" -ForegroundColor Gray
                            }
                        }
                    }
                }
                Write-Host ""
            }
            
            # Show first few events
            Write-Host "=== First 10 Events ===" -ForegroundColor Cyan
            Get-Content $TxtFile -TotalCount 10 | ForEach-Object {
                Write-Host $_ -ForegroundColor Gray
            }
            Write-Host ""
            Write-Host "See $TxtFile for complete trace" -ForegroundColor White
        } else {
            Write-Host "WARNING: Failed to convert trace to text format" -ForegroundColor Yellow
            Write-Host "   You can view the .etl file using:" -ForegroundColor Gray
            Write-Host "   - Windows Performance Analyzer (WPA)" -ForegroundColor Gray
            Write-Host "   - PerfView" -ForegroundColor Gray
            Write-Host "   - TraceView (from Windows SDK)" -ForegroundColor Gray
        }
    } else {
        Write-Host "WARNING: No ETL file found" -ForegroundColor Yellow
        Write-Host "   The trace session may have been empty or failed to write" -ForegroundColor Gray
    }
} else {
    Write-Host "ERROR: Failed to stop trace session or session not found" -ForegroundColor Red
    Write-Host ""
    Write-Host "Session may have already been stopped or never started." -ForegroundColor Gray
    Write-Host "Use start-etw-trace.ps1 to start a new session." -ForegroundColor Gray
}

Write-Host ""
Write-Host "Note: RedSalamanderMonitor.exe displays ETW events in real-time" -ForegroundColor Cyan
Write-Host "      without needing external trace sessions." -ForegroundColor Cyan
Write-Host ""
if (-not $NoPause) {
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
