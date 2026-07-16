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

    It 'streams child processes with a sanitized Path key' {
        $scriptPath = Join-Path $TestDrive 'PrintEnvironment.ps1'
        @'
$keys = @([System.Environment]::GetEnvironmentVariables('Process').Keys | Where-Object { $_ -ieq 'Path' })
Write-Output (($keys | Sort-Object) -join ',')
Write-Output ([System.Environment]::GetEnvironmentVariable('Path', 'Process'))
'@ | Set-Content -Path $scriptPath -Encoding ASCII

        $receivedLines = [System.Collections.Generic.List[string]]::new()

        $exitCode = Invoke-RSStreamingProcess `
            -FilePath 'powershell.exe' `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath) `
            -WorkingDirectory $TestDrive `
            -AdditionalEnvironment @{ 'PATH' = 'C:\StreamingChild' } `
            -OutputLineCallback {
            param(
                [string]$Line,
                [bool]$IsError
            )

                [void]$receivedLines.Add($Line)
            }

        $exitCode | Should Be 0
        $receivedLines | Should Be @('Path', 'C:\StreamingChild')
    }

    It 'kills the contained process tree when an output callback aborts streaming' {
        $grandchildScript = Join-Path $TestDrive 'GrandchildSleep.ps1'
        $parentScript = Join-Path $TestDrive 'ParentWithGrandchild.ps1'
        $grandchildPidPath = Join-Path $TestDrive 'grandchild.pid'
        @'
Start-Sleep -Seconds 60
'@ | Set-Content -LiteralPath $grandchildScript -Encoding ASCII

        @'
param(
    [Parameter(Mandatory = $true)]
    [string]$GrandchildScript,

    [Parameter(Mandatory = $true)]
    [string]$GrandchildPidPath
)

$pwsh = (Get-Process -Id $PID).Path
$startParameters = @{
    FilePath = $pwsh
    ArgumentList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $GrandchildScript))
    PassThru = $true
    WindowStyle = 'Hidden'
}
$grandchild = Start-Process @startParameters
[System.IO.File]::WriteAllText($GrandchildPidPath, $grandchild.Id.ToString())
Write-Output "grandchild=$($grandchild.Id)"
Start-Sleep -Seconds 60
'@ | Set-Content -LiteralPath $parentScript -Encoding UTF8

        $grandchildPid = 0
        $message = $null
        try {
            try {
                Invoke-RSStreamingProcess `
                    -FilePath 'powershell.exe' `
                    -Arguments @(
                        '-NoProfile',
                        '-ExecutionPolicy',
                        'Bypass',
                        '-File',
                        $parentScript,
                        '-GrandchildScript',
                        $grandchildScript,
                        '-GrandchildPidPath',
                        $grandchildPidPath
                    ) `
                    -WorkingDirectory $TestDrive `
                    -OutputLineCallback {
                    param(
                        [string]$Line,
                        [bool]$IsError
                    )

                        if ($Line -match '^grandchild=') {
                            throw 'intentional callback stop'
                        }
                    }
            }
            catch {
                $message = $_.Exception.Message
            }

            $message | Should Match 'intentional callback stop'
            (Test-Path -LiteralPath $grandchildPidPath) | Should Be $true
            $grandchildPid = [int](Get-Content -LiteralPath $grandchildPidPath -Raw)

            $exited = $false
            for ($attempt = 0; $attempt -lt 100; $attempt++) {
                if ($null -eq (Get-Process -Id $grandchildPid -ErrorAction SilentlyContinue)) {
                    $exited = $true
                    break
                }
                Start-Sleep -Milliseconds 25
            }
            $exited | Should Be $true
        }
        finally {
            if ($grandchildPid -gt 0) {
                Stop-Process -Id $grandchildPid -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
