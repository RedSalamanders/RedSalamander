Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperModule = Join-Path $repoRoot 'Tools\Modules\Build\ProcessStreaming.psm1'
$artifactOperationLockModule = Join-Path $repoRoot 'Tools\Modules\Build\ArtifactOperationLock.psm1'

Describe 'ProcessStreaming helper' {
    BeforeAll {
        Import-Module $artifactOperationLockModule -Force -Global -ErrorAction Stop
        Import-Module $helperModule -Force -ErrorAction Stop
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
        # Prove that callbacks are delivered while the child is running without
        # making the contract depend on host-specific PowerShell startup time.
        ($receivedAtMs[1] - $receivedAtMs[0]) | Should BeGreaterThan 500
        (Get-Content -Path $logPath) | Should Be @('alpha', 'omega')
    }

    It 'reports stderr lines, creates missing log directories, and returns the child exit code' {
        $scriptPath = Join-Path $TestDrive 'EmitStdErr.ps1'
        @'
Write-Output "stdout-line"
[Console]::Error.WriteLine("stderr-line")
exit 17
'@ | Set-Content -Path $scriptPath -Encoding ASCII

        $logPath = Join-Path $TestDrive 'logs\nested\stdout.log'
        $errorLogPath = Join-Path $TestDrive 'logs\nested\stderr.log'
        $received = [System.Collections.Generic.List[psobject]]::new()

        $exitCode = Invoke-RSStreamingProcess `
            -FilePath 'powershell.exe' `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath) `
            -WorkingDirectory $TestDrive `
            -LogPath $logPath `
            -ErrorLogPath $errorLogPath `
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
        (Get-Content -Path $logPath) | Should Be @('stdout-line')
        (Get-Content -Path $errorLogPath) | Should Be @('stderr-line')
    }

    It 'validates split log paths before creation and disposes stdout if stderr open fails' {
        $samePath = Join-Path $TestDrive 'same.log'
        $samePathMessage = $null
        try {
            Invoke-RSStreamingProcess `
                -FilePath 'powershell.exe' `
                -Arguments @('-NoProfile', '-Command', 'exit 0') `
                -WorkingDirectory $TestDrive `
                -LogPath $samePath `
                -ErrorLogPath $samePath
        }
        catch {
            $samePathMessage = $_.Exception.Message
        }

        $samePathMessage | Should Match 'LogPath and ErrorLogPath must name different files'
        (Test-Path -LiteralPath $samePath) | Should Be $false

        $stdoutPath = Join-Path $TestDrive 'partial-open-stdout.log'
        $errorDirectory = Join-Path $TestDrive 'error-log-is-a-directory'
        [void](New-Item -ItemType Directory -Path $errorDirectory -Force)
        $openFailureMessage = $null
        try {
            Invoke-RSStreamingProcess `
                -FilePath 'powershell.exe' `
                -Arguments @('-NoProfile', '-Command', 'exit 0') `
                -WorkingDirectory $TestDrive `
                -LogPath $stdoutPath `
                -ErrorLogPath $errorDirectory
        }
        catch {
            $openFailureMessage = $_.Exception.Message
        }

        [string]::IsNullOrWhiteSpace($openFailureMessage) | Should Be $false
        $exclusive = $null
        try {
            $exclusive = [System.IO.File]::Open(
                $stdoutPath,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None)
            $exclusive.Length | Should Be 0
        }
        finally {
            if ($null -ne $exclusive) {
                $exclusive.Dispose()
            }
        }
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

    It 'bounds output drain when an exited parent leaves inherited pipe handles open' {
        $grandchildScript = Join-Path $TestDrive 'HoldInheritedPipe.ps1'
        $parentScript = Join-Path $TestDrive 'ExitWithPipeHoldingChild.ps1'
        $grandchildPidPath = Join-Path $TestDrive 'pipe-holder.pid'
        $logPath = Join-Path $TestDrive 'pipe-drain.log'
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
$grandchild = Start-Process `
    -FilePath $pwsh `
    -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $GrandchildScript)) `
    -PassThru `
    -NoNewWindow
[System.IO.File]::WriteAllText($GrandchildPidPath, $grandchild.Id.ToString())
[Console]::Out.Write('parent-final-without-newline')
[Console]::Out.Flush()
'@ | Set-Content -LiteralPath $parentScript -Encoding UTF8

        $grandchildPid = 0
        try {
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $receivedLines = [System.Collections.Generic.List[string]]::new()
            $exitCode = Invoke-RSStreamingProcess `
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
                -LogPath $logPath `
                -TimeoutMilliseconds 5000 `
                -OutputDrainTimeoutMilliseconds 250 `
                -OutputLineCallback {
                param([string]$Line, [bool]$IsError)
                    [void]$receivedLines.Add($Line)
                }
            $stopwatch.Stop()

            $exitCode | Should Be 0
            $stopwatch.ElapsedMilliseconds | Should BeLessThan 5000
            (Test-Path -LiteralPath $grandchildPidPath) | Should Be $true
            $grandchildPid = [int](Get-Content -LiteralPath $grandchildPidPath -Raw)
            ($receivedLines -contains 'parent-final-without-newline') | Should Be $true
            (Get-Content -LiteralPath $logPath -Raw) | Should Match 'parent-final-without-newline'
            (Get-Content -LiteralPath $logPath -Raw) | Should Match 'closing the containment job to release inherited pipe handles'

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

    It 'bounds process runtime and closes the contained process tree on timeout' {
        $scriptPath = Join-Path $TestDrive 'NeverCompletes.ps1'
        $processIdPath = Join-Path $TestDrive 'timed-out.pid'
        @'
param([Parameter(Mandatory = $true)][string]$ProcessIdPath)
[System.IO.File]::WriteAllText($ProcessIdPath, $PID.ToString())
Write-Output 'ready'
Start-Sleep -Seconds 60
'@ | Set-Content -LiteralPath $scriptPath -Encoding ASCII

        $message = $null
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            Invoke-RSStreamingProcess `
                -FilePath 'powershell.exe' `
                -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, '-ProcessIdPath', $processIdPath) `
                -WorkingDirectory $TestDrive `
                -TimeoutMilliseconds 1000 `
                -OutputDrainTimeoutMilliseconds 250 `
                -OutputLineCallback { param([string]$Line, [bool]$IsError) }
        }
        catch {
            $message = $_.Exception.Message
        }
        finally {
            $stopwatch.Stop()
        }

        $message | Should Match "Timed out after 1000 milliseconds running 'powershell.exe'"
        $stopwatch.ElapsedMilliseconds | Should BeLessThan 6000
        (Test-Path -LiteralPath $processIdPath) | Should Be $true
        $timedOutProcessId = [int](Get-Content -LiteralPath $processIdPath -Raw)
        $exited = $false
        for ($attempt = 0; $attempt -lt 100; $attempt++) {
            if ($null -eq (Get-Process -Id $timedOutProcessId -ErrorAction SilentlyContinue)) {
                $exited = $true
                break
            }
            Start-Sleep -Milliseconds 25
        }
        $exited | Should Be $true
    }

    It 'propagates artifact-operation delegation through the shared launch path' {
        $testRepo = Join-Path $TestDrive 'delegated-artifact-repo'
        $childScript = Join-Path $TestDrive 'ReadArtifactDelegation.ps1'
        [void](New-Item -ItemType Directory -Path $testRepo -Force)
        @'
$hasDelegation = -not [string]::IsNullOrWhiteSpace(
    [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_DELEGATION_TOKEN', 'Process'))
Write-Output "delegated=$hasDelegation"
'@ | Set-Content -LiteralPath $childScript -Encoding ASCII

        $rootLock = Enter-RSArtifactOperationLock `
            -RepoRoot $testRepo `
            -Operation 'streaming-parent' `
            -Scope @{ configuration = 'Debug'; platform = 'x64' }
        try {
            $receivedLines = [System.Collections.Generic.List[string]]::new()
            $exitCode = Invoke-RSStreamingProcess `
                -FilePath 'powershell.exe' `
                -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $childScript) `
                -WorkingDirectory $TestDrive `
                -TimeoutMilliseconds 5000 `
                -DelegateArtifactOperation `
                -OutputLineCallback {
                param([string]$Line, [bool]$IsError)
                    [void]$receivedLines.Add($Line)
                }

            $exitCode | Should Be 0
            $receivedLines | Should Be @('delegated=True')
        }
        finally {
            Exit-RSArtifactOperationLock -Lock $rootLock
        }
    }

    It 'adds timestamped deployment log paths while preserving helper exceptions' {
        $deploymentPath = Join-Path $repoRoot 'Tools\Tests\RedSalamanderPluginDeployment.Tests.ps1'
        $tokens = $null
        $parseErrors = $null
        $deploymentAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $deploymentPath,
            [ref]$tokens,
            [ref]$parseErrors)
        @($parseErrors).Count | Should Be 0
        $wrapperAst = @($deploymentAst.FindAll({
                    param($node)
                    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -eq 'Invoke-RSTargetedBuildForDeploymentTest'
                }, $true))[0]
        $null -eq $wrapperAst | Should Be $false

        $buildLogDir = Join-Path $TestDrive 'deployment-logs'
        $buildScript = Join-Path $TestDrive 'build.ps1'
        $repoRoot = $TestDrive
        Invoke-Expression $wrapperAst.Extent.Text
        function Invoke-RSStreamingProcess {
            param(
                [string]$FilePath,
                [string[]]$Arguments,
                [string]$WorkingDirectory,
                [string]$LogPath,
                [string]$ErrorLogPath,
                [hashtable]$AdditionalEnvironment,
                [int]$TimeoutMilliseconds,
                [int]$OutputDrainTimeoutMilliseconds,
                [switch]$DelegateArtifactOperation,
                [scriptblock]$OutputLineCallback
            )

            throw [System.TimeoutException]::new('simulated streaming timeout')
        }

        $capturedException = $null
        try {
            Invoke-RSTargetedBuildForDeploymentTest -Platform x64 -Configuration Debug
        }
        catch {
            $capturedException = $_.Exception
        }
        finally {
            Remove-Item Function:\Invoke-RSTargetedBuildForDeploymentTest -ErrorAction SilentlyContinue
            Remove-Item Function:\Invoke-RSStreamingProcess -ErrorAction SilentlyContinue
        }

        $null -eq $capturedException | Should Be $false
        $capturedException.Message | Should Match 'simulated streaming timeout'
        $capturedException.Message | Should Match 'stdout: .+pester-targeted-plugin-deployment-\d{8}_\d{6}_\d{3}\.out\.log'
        $capturedException.Message | Should Match 'stderr: .+pester-targeted-plugin-deployment-\d{8}_\d{6}_\d{3}\.err\.log'
        $capturedException.InnerException.GetType().FullName | Should Be 'System.TimeoutException'
        $capturedException.InnerException.Message | Should Be 'simulated streaming timeout'
    }

    It 'keeps targeted plugin deployment on the bounded shared streaming contract' {
        $deploymentSource = Get-Content `
            -LiteralPath (Join-Path $repoRoot 'Tools\Tests\RedSalamanderPluginDeployment.Tests.ps1') `
            -Raw

        $deploymentSource | Should Match 'Invoke-RSStreamingProcess[\s\S]*?-DelegateArtifactOperation'
        $deploymentSource | Should Match '-LogPath \$stdoutLog[\s\S]*?-ErrorLogPath \$stderrLog'
        $deploymentSource | Should Match '-TimeoutMilliseconds \(15 \* 60 \* 1000\)'
        $deploymentSource | Should Match '-OutputDrainTimeoutMilliseconds 5000'
        $deploymentSource | Should Match 'throw \[System\.InvalidOperationException\]::new\(\$message, \$originalException\)'
        $deploymentSource | Should Not Match 'ReadToEndAsync|\.Result'
    }
}
