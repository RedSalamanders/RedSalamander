# Start ETW Trace Session for RedSalamanderMonitor
# This script starts an ETW trace session to capture events from the RedSalamander provider
# Use this BEFORE running MonitorTest if you want to capture ETW events externally

param(
    [switch]$NoPause
)

$ProviderGuid = "{440c70f6-6c6b-4ff7-9a3f-0b7db411b31a}"
$SessionName = "RedSalamanderTrace"
$OutputFile = "RedSalamanderTrace.etl"

Write-Host "=== Starting ETW Trace Session ===" -ForegroundColor Cyan
Write-Host "Provider GUID: $ProviderGuid" -ForegroundColor Gray
Write-Host "Session Name:  $SessionName" -ForegroundColor Gray
Write-Host "Output File:   $OutputFile" -ForegroundColor Gray
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "WARNING: Not running as Administrator" -ForegroundColor Yellow
    Write-Host "   ETW trace sessions typically require elevated privileges" -ForegroundColor Yellow
    Write-Host ""
}

# Stop any existing session with the same name
Write-Host "Stopping existing session (if any)..." -ForegroundColor Gray
logman stop $SessionName -ets 2>$null | Out-Null

# Start new trace session with optimized buffer settings
# Buffer configuration tuned for high-frequency event generation:
#   -nb 32    = MinimumBuffers (32 x 256 KB = 8 MB minimum capacity)
#   -bs 256   = BufferSize in KB (larger buffers reduce overhead)
#   -max 128  = MaximumBuffers (128 x 256 KB = 32 MB maximum capacity)
#   -ft 1     = FlushTimer in seconds (balance latency vs. throughput)
Write-Host "Starting ETW trace session with optimized buffers..." -ForegroundColor Gray
Write-Host "  MinBuffers: 32 x 256KB = 8 MB" -ForegroundColor DarkGray
Write-Host "  MaxBuffers: 512 x 256KB = 128 MB max capacity" -ForegroundColor DarkGray
Write-Host "  FlushTimer: 1 second" -ForegroundColor DarkGray
$result = logman create trace $SessionName -p $ProviderGuid 0xFFFFFFFFFFFFFFFF 5 -o $OutputFile -nb 32 -bs 256 -max 512 -ft 1 -ets

if ($LASTEXITCODE -eq 0) {
    Write-Host "OK: ETW trace session started successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "The trace is now capturing events in real-time." -ForegroundColor White
    Write-Host "Run your tests, then use stop-etw-trace.ps1 to stop and view results." -ForegroundColor White
    Write-Host ""
    Write-Host "Note: RedSalamanderMonitor.exe has its own built-in ETW listener" -ForegroundColor Cyan
    Write-Host "      and does not require this external trace session." -ForegroundColor Cyan
} else {
    Write-Host "ERROR: Failed to start ETW trace session" -ForegroundColor Red
    Write-Host ""
    Write-Host "Common issues:" -ForegroundColor Yellow
    Write-Host "  - Session already exists (use stop-etw-trace.ps1 first)" -ForegroundColor Gray
    Write-Host "  - Insufficient privileges (run PowerShell as Administrator)" -ForegroundColor Gray
    Write-Host "  - Provider not registered (ensure Common.dll is loaded)" -ForegroundColor Gray
}

Write-Host ""
if (-not $NoPause) {
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
# --- IGNORE ---

