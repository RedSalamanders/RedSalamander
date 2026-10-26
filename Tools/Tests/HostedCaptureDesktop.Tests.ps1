Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '../Modules/Testing/HostedCaptureDesktop.psm1') -Force

InModuleScope HostedCaptureDesktop {
    function Assert-RSHostedCaptureFailure {
        param([scriptblock]$Action, [string]$Pattern)
        $message = $null
        try { $null = & $Action }
        catch { $message = $_.Exception.Message }
        ($null -ne $message) | Should Be $true
        $message | Should Match $Pattern
    }

    Describe 'Hosted capture desktop ownership' {
        BeforeEach {
            $script:savedEnvironment = @{}
            foreach ($name in @('GITHUB_ACTIONS','RUNNER_ENVIRONMENT','RUNNER_OS','RUNNER_ARCH')) {
                $script:savedEnvironment[$name] = [Environment]::GetEnvironmentVariable($name)
            }
            $env:GITHUB_ACTIONS = 'true'
            $env:RUNNER_ENVIRONMENT = 'github-hosted'
            $env:RUNNER_OS = 'Windows'
            $env:RUNNER_ARCH = 'X64'
            $script:connected = $false
            Mock Get-RSHostedCaptureProcessSession { 7 }
            Mock Get-RSHostedCaptureConsoleSession { if ($script:connected) { 7 } else { 1 } }
            Mock Connect-RSHostedCaptureSession { $script:connected = $true }
        }
        AfterEach {
            foreach ($name in $script:savedEnvironment.Keys) {
                [Environment]::SetEnvironmentVariable($name, $script:savedEnvironment[$name])
            }
        }

        It 'refuses a local launch without querying or changing sessions' {
            $env:GITHUB_ACTIONS = 'false'
            Assert-RSHostedCaptureFailure { Initialize-RSHostedCaptureDesktop } 'GitHub-hosted Windows x64'
            Assert-MockCalled Get-RSHostedCaptureProcessSession -Times 0 -Exactly -Scope It
            Assert-MockCalled Connect-RSHostedCaptureSession -Times 0 -Exactly -Scope It
        }
        It 'refuses a self-hosted runner before changing sessions' {
            $env:RUNNER_ENVIRONMENT = 'self-hosted'
            Assert-RSHostedCaptureFailure { Initialize-RSHostedCaptureDesktop } 'GitHub-hosted Windows x64'
            Assert-MockCalled Get-RSHostedCaptureProcessSession -Times 0 -Exactly -Scope It
            Assert-MockCalled Connect-RSHostedCaptureSession -Times 0 -Exactly -Scope It
        }
        It 'refuses a non-Windows runner before changing sessions' {
            $env:RUNNER_OS = 'Linux'
            Assert-RSHostedCaptureFailure { Initialize-RSHostedCaptureDesktop } 'GitHub-hosted Windows x64'
            Assert-MockCalled Connect-RSHostedCaptureSession -Times 0 -Exactly -Scope It
        }
        It 'leaves the already qualified ARM64 desktop untouched' {
            $env:RUNNER_ARCH = 'ARM64'
            Assert-RSHostedCaptureFailure { Initialize-RSHostedCaptureDesktop } 'GitHub-hosted Windows x64'
            Assert-MockCalled Connect-RSHostedCaptureSession -Times 0 -Exactly -Scope It
        }
        It 'refuses session zero rather than attaching another user session' {
            Mock Get-RSHostedCaptureProcessSession { 0 }
            Assert-RSHostedCaptureFailure { Initialize-RSHostedCaptureDesktop } 'service session zero'
            Assert-MockCalled Connect-RSHostedCaptureSession -Times 0 -Exactly -Scope It
        }
        It 'does not reconnect an already active console session' {
            Mock Get-RSHostedCaptureConsoleSession { 7 }
            $result = Initialize-RSHostedCaptureDesktop
            $result.Connected | Should Be $false
            $result.ProcessSession | Should Be 7
            Assert-MockCalled Connect-RSHostedCaptureSession -Times 0 -Exactly -Scope It
        }
        It 'connects only its own process session and verifies the resulting console' {
            $result = Initialize-RSHostedCaptureDesktop
            $result.Connected | Should Be $true
            $result.PreviousConsoleSession | Should Be 1
            $result.ConsoleSession | Should Be 7
            Assert-MockCalled Connect-RSHostedCaptureSession -Times 1 -Exactly -Scope It -ParameterFilter { $SessionId -eq 7 }
        }
        It 'propagates connection failure rather than running capture with a broken desktop' {
            Mock Connect-RSHostedCaptureSession { throw 'Connection refused.' }
            Assert-RSHostedCaptureFailure { Initialize-RSHostedCaptureDesktop } 'Connection refused'
        }
        It 'rejects a successful command that did not connect the owned session' {
            Mock Get-RSHostedCaptureConsoleSession { 1 }
            Assert-RSHostedCaptureFailure { Initialize-RSHostedCaptureDesktop -TimeoutSeconds 1 } 'did not reach the console'
            Assert-MockCalled Connect-RSHostedCaptureSession -Times 1 -Exactly -Scope It -ParameterFilter { $SessionId -eq 7 }
        }
        It 'rejects a connection that succeeds after the shared deadline' {
            Mock Connect-RSHostedCaptureSession {
                Start-Sleep -Milliseconds 1100
                $script:connected = $true
            }
            Assert-RSHostedCaptureFailure { Initialize-RSHostedCaptureDesktop -TimeoutSeconds 1 } 'did not reach the console'
        }
    }

    Describe 'Contained hosted console command' {
        It 'bounds the owned session command and its redirected output drain' {
            Mock Invoke-RSStreamingProcess { 0 }
            Connect-RSHostedCaptureSession -SessionId 7 -TimeoutMilliseconds 1200
            Assert-MockCalled Invoke-RSStreamingProcess -Times 1 -Exactly -Scope It -ParameterFilter {
                $FilePath -eq (Join-Path $env:SystemRoot 'System32/tscon.exe') -and
                $Arguments.Count -eq 2 -and $Arguments[0] -eq '7' -and $Arguments[1] -eq '/dest:console' -and
                $TimeoutMilliseconds -eq 1200 -and $OutputDrainTimeoutMilliseconds -eq 250
            }
        }
        It 'fails on a nonzero connection exit code' {
            Mock Invoke-RSStreamingProcess { 5 }
            Assert-RSHostedCaptureFailure { Connect-RSHostedCaptureSession -SessionId 7 -TimeoutMilliseconds 1200 } 'failed: 5'
        }
        It 'propagates a contained-process timeout instead of accepting late console state' {
            Mock Invoke-RSStreamingProcess { throw 'Timed out running owned tscon.' }
            Assert-RSHostedCaptureFailure { Connect-RSHostedCaptureSession -SessionId 7 -TimeoutMilliseconds 1200 } 'Timed out'
        }
    }
}
