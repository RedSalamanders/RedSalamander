Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot '../Build/ProcessStreaming.psm1') -Scope Local -ErrorAction Stop

function Get-RSHostedCaptureProcessSession {
    $process = [Diagnostics.Process]::GetCurrentProcess()
    try { return $process.SessionId }
    finally { $process.Dispose() }
}

function Get-RSHostedCaptureConsoleSession {
    if (-not ('RedSalamander.Testing.HostedConsole' -as [type])) {
        Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
namespace RedSalamander.Testing {
    public static class HostedConsole {
        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern uint WTSGetActiveConsoleSessionId();
    }
}
'@
    }
    return [RedSalamander.Testing.HostedConsole]::WTSGetActiveConsoleSessionId()
}

function Connect-RSHostedCaptureSession {
    param(
        [Parameter(Mandatory)][ValidateRange(1, [int]::MaxValue)][int]$SessionId,
        [Parameter(Mandatory)][ValidateRange(1, 30000)][int]$TimeoutMilliseconds
    )

    $tscon = Join-Path $env:SystemRoot 'System32/tscon.exe'
    $exitCode = Invoke-RSStreamingProcess -FilePath $tscon -Arguments @([string]$SessionId, '/dest:console') `
        -TimeoutMilliseconds $TimeoutMilliseconds -OutputDrainTimeoutMilliseconds 250
    if ($exitCode -ne 0) {
        throw "Connecting the owned hosted session $SessionId to its console failed: $exitCode."
    }
}

function Initialize-RSHostedCaptureDesktop {
    <#
    .SYNOPSIS
    Connects the GitHub-hosted Windows x64 job's own session to its console before HWND capture.
    .DESCRIPTION
    Refuses local, self-hosted, non-Windows, ARM64 and service-session callers before any
    session mutation. It never selects another user's session, launches tests, changes
    capture assertions or substitutes desktop capture. A failed connection is fatal.
    .PARAMETER TimeoutSeconds
    One deadline for the contained connection command and console-identity verification.
    .OUTPUTS
    Session identities, whether a connection was needed, and elapsed milliseconds.
    #>
    [CmdletBinding()]
    param([ValidateRange(1, 30)][int]$TimeoutSeconds = 30)

    if ($env:GITHUB_ACTIONS -cne 'true' -or $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or
        $env:RUNNER_OS -cne 'Windows' -or $env:RUNNER_ARCH -cne 'X64') {
        throw 'Hosted capture desktop preparation requires a GitHub-hosted Windows x64 job.'
    }

    $clock = [Diagnostics.Stopwatch]::StartNew()
    $budgetMilliseconds = $TimeoutSeconds * 1000
    $sessionId = Get-RSHostedCaptureProcessSession
    if ($sessionId -le 0) {
        throw 'Hosted capture cannot run in service session zero.'
    }
    $before = Get-RSHostedCaptureConsoleSession
    $console = $before
    Write-Host "Hosted capture desktop: process session=$sessionId, console session=$before."
    $connected = $before -ne $sessionId
    if ($connected) {
        $remaining = $budgetMilliseconds - [int]$clock.ElapsedMilliseconds
        if ($remaining -le 0) { throw 'Hosted capture console deadline expired before connection.' }
        Connect-RSHostedCaptureSession -SessionId $sessionId -TimeoutMilliseconds $remaining
    }
    do {
        if ($clock.ElapsedMilliseconds -ge $budgetMilliseconds) { break }
        $console = Get-RSHostedCaptureConsoleSession
        if ($clock.ElapsedMilliseconds -ge $budgetMilliseconds) { break }
        if ($console -eq $sessionId) {
            return [pscustomobject]@{
                ProcessSession = $sessionId
                PreviousConsoleSession = $before
                ConsoleSession = $console
                Connected = $connected
                ElapsedMilliseconds = $clock.ElapsedMilliseconds
            }
        }
        Start-Sleep -Milliseconds 100
    } while ($true)
    throw "Hosted process session $sessionId did not reach the console within $TimeoutSeconds seconds (console=$console)."
}

Export-ModuleMember -Function Initialize-RSHostedCaptureDesktop
