Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Tools\ProcessStreaming.ps1'

Describe 'ProcessStreaming helper' {
    BeforeAll {
        . $helperScript
    }

    It 'streams process output before the process exits and mirrors it to the log file' {
        $scriptPath = Join-Path $TestDrive 'EmitLines.ps1'
        @'
Write-Output "alpha"
Start-Sleep -Milliseconds 900
Write-Output "omega"
'@ | Set-Content -Path $scriptPath -Encoding ASCII

        $logPath = Join-Path $TestDrive 'stream.log'
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $receivedLines = [System.Collections.Generic.List[string]]::new()
        $receivedAtMs = [System.Collections.Generic.List[int64]]::new()

        $exitCode = Invoke-RSStreamingProcess `
            -FilePath 'powershell.exe' `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath) `
            -WorkingDirectory $TestDrive `
            -LogPath $logPath `
            -OutputLineCallback {
            param(
                [string]$Line,
                [bool]$IsError
            )

                [void]$receivedLines.Add($Line)
                [void]$receivedAtMs.Add($stopwatch.ElapsedMilliseconds)
            }

        $stopwatch.Stop()

        $exitCode | Should Be 0
        $receivedLines | Should Be @('alpha', 'omega')
        $receivedAtMs.Count | Should Be 2
        $receivedAtMs[0] | Should BeLessThan 800
        (Get-Content -Path $logPath) | Should Be @('alpha', 'omega')
    }

    It 'reports stderr lines, creates missing log directories, and returns the child exit code' {
        $scriptPath = Join-Path $TestDrive 'EmitStdErr.ps1'
        @'
Write-Output "stdout-line"
[Console]::Error.WriteLine("stderr-line")
exit 17
'@ | Set-Content -Path $scriptPath -Encoding ASCII

        $logPath = Join-Path $TestDrive 'logs\nested\stream.log'
        $received = [System.Collections.Generic.List[psobject]]::new()

        $exitCode = Invoke-RSStreamingProcess `
            -FilePath 'powershell.exe' `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath) `
            -WorkingDirectory $TestDrive `
            -LogPath $logPath `
            -OutputLineCallback {
            param(
                [string]$Line,
                [bool]$IsError
            )

                [void]$received.Add([pscustomobject]@{
                        Line = $Line
                        IsError = $IsError
                    })
            }

        $exitCode | Should Be 17
        $global:LASTEXITCODE | Should Be 17
        (Test-Path $logPath) | Should Be $true
        $received.Count | Should Be 2
        $received[0].Line | Should Be 'stdout-line'
        $received[0].IsError | Should Be $false
        $received[1].Line | Should Be 'stderr-line'
        $received[1].IsError | Should Be $true
        (Get-Content -Path $logPath) | Should Be @('stdout-line', 'stderr-line')
    }
}
