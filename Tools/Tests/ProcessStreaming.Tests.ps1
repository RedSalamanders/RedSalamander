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

    It 'waits for the hosted MSVC GUI bootstrapper and handles its actual success reboot and failure exits' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($env:REDSALAMANDER_TEST_RUN_ID)) {
            throw 'The managed bootstrapper control requires an existing governed run ID.'
        }
        # The legacy managed WinApp loader has a smaller path budget than the product.
        # Keep its fixture in a unique retained child, not the arbitrarily deep Pester drive.
        # The historical exit -1 had no retained iteration; this budget does not prove its cause.
        $fixtureBase = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'ps' -Case ('m-' + [guid]::NewGuid().ToString('N'))
        $source = Join-Path $fixtureBase 'GuiBootstrapper.cs'
        $gui = Join-Path $fixtureBase 'GUI bootstrapper.exe'
        $compile = Join-Path $fixtureBase 'CompileGui.ps1'
        $fixturePathLimit = 240
        $plannedPaths = @($source, $gui, ($gui + '.config'), $compile, (Join-Path $fixtureBase 'compile.stdout.log'), (Join-Path $fixtureBase 'compile.stderr.log'))
        foreach ($fixtureExit in @(0,3010,17)) {
            $iterationRoot = Join-Path $fixtureBase ("hosted $fixtureExit é日")
            $plannedScratch = Join-Path $iterationRoot ('msvc-14.52-' + ('0' * 32))
            $plannedPaths += @(
                (Join-Path $plannedScratch 'vs_BuildTools.exe'),
                ((Join-Path $plannedScratch 'vs_BuildTools.exe') + '.config'),
                (Join-Path $plannedScratch 'bootstrapper.stdout.log'),
                (Join-Path $plannedScratch 'bootstrapper.stderr.log'),
                (Join-Path $iterationRoot 'VS-Msvc-14.52/installed.args')
            )
        }
        foreach ($path in $plannedPaths) {
            if ($path.Length -ge $fixturePathLimit) { throw "Managed bootstrapper fixture exceeds its conservative path budget: $($path.Length)." }
        }
        $fixtureRecord = [ordered]@{
            schema='red-salamander.managed-bootstrapper-fixture.v1'
            path_limit=$fixturePathLimit
            planned_path_lengths=@($plannedPaths | ForEach-Object { $_.Length })
            phase='before-compile'
        }
        $fixtureRecord | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $fixtureBase 'fixture.json') -Encoding UTF8
        @'
using System;
using System.IO;
using System.Threading;
class GuiBootstrapper {
    static int Main(string[] args) {
        string root = args[Array.IndexOf(args, "--installPath") + 1];
        Thread.Sleep(300);
        Directory.CreateDirectory(root);
        File.WriteAllLines(Path.Combine(root, "installed.args"), args);
        return Int32.Parse(Environment.GetEnvironmentVariable("RS_MsvcFixtureExitCode"));
    }
}
'@ | Set-Content -LiteralPath $source -Encoding UTF8
        'param($SourcePath,$OutputPath); Add-Type -Path $SourcePath -OutputAssembly $OutputPath -OutputType WindowsApplication' |
            Set-Content -LiteralPath $compile -Encoding UTF8
        $compileExit = Invoke-RSStreamingProcess -FilePath 'powershell.exe' -Arguments @('-NoProfile','-File',$compile,$source,$gui) -WorkingDirectory $fixtureBase -LogPath (Join-Path $fixtureBase 'compile.stdout.log') -ErrorLogPath (Join-Path $fixtureBase 'compile.stderr.log')
        $fixtureRecord.compile_exit = $compileExit
        $fixtureRecord.source_sha256 = (Get-FileHash -LiteralPath $source).Hash
        $fixtureRecord.phase = 'after-compile'
        if (Test-Path -LiteralPath $gui) { $fixtureRecord.compiled_sha256 = (Get-FileHash -LiteralPath $gui).Hash }
        $fixtureRecord | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $fixtureBase 'fixture.json') -Encoding UTF8
        $compileExit | Should Be 0
        $image = [IO.File]::ReadAllBytes($gui)
        $peOffset = [BitConverter]::ToInt32($image, 0x3c)
        [BitConverter]::ToUInt16($image, $peOffset + 24 + 68) | Should Be 2

        $tokens = $null
        $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot '.github/actions/setup-msvc/Setup.ps1'),[ref]$tokens,[ref]$errors)
        $errors.Count | Should Be 0
        $branches = @($ast.EndBlock.Statements | Where-Object {
            $_ -is [Management.Automation.Language.IfStatementAst] -and $_.Extent.Text -match '^if\s*\(\s*-not\s+\$selected\s*\)'
        })
        $branches.Count | Should Be 1
        $provision = [scriptblock]::Create($branches[0].Extent.Text)
        $savedEnvironment = @{}
        foreach($key in @('GITHUB_ACTIONS','RUNNER_ENVIRONMENT','RUNNER_TEMP','RS_MsvcFixtureExitCode')) {
            $savedEnvironment[$key] = [Environment]::GetEnvironmentVariable($key,'Process')
        }
        $previousExit = Get-Variable LASTEXITCODE -Scope Global -ErrorAction SilentlyContinue
        $previousExitValue = if($previousExit){$previousExit.Value}else{$null}
        $fixtureHasPrimaryFailure = $false
        try {
            $env:GITHUB_ACTIONS = 'true'
            $env:RUNNER_ENVIRONMENT = 'github-hosted'
            foreach($fixtureExit in @(0,3010,17)) {
                & {
                    param($fixtureExit,$provision,$gui,$fixtureRoot,$repoRoot,$fixturePathLimit)
                    [void][IO.Directory]::CreateDirectory($fixtureRoot)
                    $env:RUNNER_TEMP = $fixtureRoot
                    $env:RS_MsvcFixtureExitCode = [string]$fixtureExit
                    $root = $repoRoot
                    $required = '14.52'
                    $selected = $null
                    $installerExitCode = $null
                    $scratch = $null
                    $VerifyOnly = $false
                    $marker = Join-Path $fixtureRoot 'VS-Msvc-14.52/installed.args'
                    function Invoke-WebRequest {
                        param($Uri,$OutFile)
                        foreach ($path in @($OutFile, ($OutFile + '.config'), (Join-Path (Split-Path $OutFile) 'bootstrapper.stdout.log'), (Join-Path (Split-Path $OutFile) 'bootstrapper.stderr.log'), $marker)) {
                            if ($path.Length -ge $fixturePathLimit) { throw "Actual managed bootstrapper fixture exceeds its path budget: $($path.Length)." }
                        }
                        Copy-Item -LiteralPath $gui -Destination $OutFile
                    }
                    function Get-AuthenticodeSignature { param($LiteralPath); [pscustomobject]@{Status='Valid';SignerCertificate=[pscustomobject]@{Subject='CN=Fixture, O=Microsoft Corporation, C=US'}} }
                    function Get-SelectedTools {
                        if(-not(Test-Path -LiteralPath $marker)){throw 'Compiler verification ran before the GUI installer completed.'}
                        return [ordered]@{version='14.52.0';platform='x64'}
                    }
                    Remove-Variable LASTEXITCODE -Scope Global -ErrorAction SilentlyContinue
                    $message = ''
                    # Evaluate in this already-isolated iteration scope so the real branch's
                    # actual exit and selected-state locals can be retained before assertions.
                    try { . $provision } catch { $message = $_.Exception.Message }
                    if ($message.Length -gt 2048) { $message = $message.Substring(0,2048) }
                    $markerPresent = Test-Path -LiteralPath $marker
                    $installedArguments = if ($markerPresent) { [IO.File]::ReadAllLines($marker) } else { @() }
                    $observed = [ordered]@{
                        fixture_exit=$fixtureExit
                        actual_exit=$installerExitCode
                        marker_present=$markerPresent
                        inner_message=$message
                        selected_present=($null -ne $selected)
                        arguments=@($installedArguments)
                        compiled_sha256=(Get-FileHash -LiteralPath $gui).Hash
                        scratch=$scratch
                        status='observed-before-assertions'
                    }
                    if ($scratch) {
                        $observed.installer_path_length = (Join-Path $scratch 'vs_BuildTools.exe').Length
                        $observed.logs = @(foreach ($name in @('bootstrapper.stdout.log','bootstrapper.stderr.log')) {
                            $path = Join-Path $scratch $name
                            if (Test-Path -LiteralPath $path) { [ordered]@{name=$name;sha256=(Get-FileHash -LiteralPath $path).Hash;bytes=(Get-Item -LiteralPath $path).Length} }
                        })
                    }
                    $record = Join-Path $fixtureRoot 'iteration.json'
                    $observed | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $record -Encoding UTF8
                    Write-Host ('Managed bootstrapper iteration: ' + ($observed | ConvertTo-Json -Depth 5 -Compress -EscapeHandling EscapeNonAscii))
                    if (-not $markerPresent) {
                        $observed.status = 'failed'
                        $observed | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $record -Encoding UTF8
                        throw "Managed bootstrapper fixture $fixtureExit returned '$installerExitCode' without its marker. Inner: $message Evidence: $record"
                    }
                    try {
                        $markerPresent | Should Be $true
                        $installerExitCode | Should Be $fixtureExit
                        if($fixtureExit -eq 17) {
                            $message | Should Match 'VS provisioning failed: 17'
                            ($null -eq $selected) | Should Be $true
                        } else {
                            $message | Should Be ''
                            ($null -ne $selected) | Should Be $true
                        }
                        $installedArguments[[Array]::IndexOf($installedArguments,'--installPath') + 1] | Should Be (Join-Path $fixtureRoot 'VS-Msvc-14.52')
                        ($installedArguments -contains '--wait') | Should Be $true
                        $observed.status = 'passed'
                    } finally {
                        if ($observed.status -ne 'passed') { $observed.status = 'failed' }
                        $observed | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $record -Encoding UTF8
                    }
                } $fixtureExit $provision $gui (Join-Path $fixtureBase ("hosted $fixtureExit é日")) $repoRoot $fixturePathLimit
            }
            $fixtureRecord.phase = 'passed-all-three'
        } catch {
            $fixtureHasPrimaryFailure = $true
            $fixtureRecord.phase = 'failed'
            throw
        } finally {
            try {
                try {
                    $fixtureRecord | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $fixtureBase 'fixture.json') -Encoding UTF8 -ErrorAction Stop
                    Write-Host ("Managed bootstrapper retained evidence: $fixtureBase")
                } catch [IO.IOException], [UnauthorizedAccessException], [Security.SecurityException] {
                    # Only file-access retention failures are secondary to an existing primary.
                    # Fatal or unexpected exceptions propagate; never serialize the ErrorRecord.
                    if (-not $fixtureHasPrimaryFailure) { throw }
                    $secondaryMessage = [string]$_.Exception.Message
                    if ($secondaryMessage.Length -gt 512) { $secondaryMessage = $secondaryMessage.Substring(0,512) }
                    Write-Warning ("Managed bootstrapper secondary retention failure: $secondaryMessage") -WarningAction Continue
                }
            } finally {
                foreach($key in $savedEnvironment.Keys){[Environment]::SetEnvironmentVariable($key,$savedEnvironment[$key],'Process')}
                if($previousExit){$global:LASTEXITCODE=$previousExitValue}else{Remove-Variable LASTEXITCODE -Scope Global -ErrorAction SilentlyContinue}
            }
        }
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

        $sandboxCase = Join-Path $TestDrive 'RedSalamander.Perf\runs\fixture-20261031\scratch\pester-targeted-plugin-deployment\build-capture'
        [void](New-Item -ItemType Directory -Path $sandboxCase -Force)
        $script:scratchAllocation = $null
        function New-RSTestSandboxScratchDirectory {
            param([string]$RepoRoot, [string]$Harness, [string]$Case)
            $script:scratchAllocation = [pscustomobject]@{ RepoRoot = $RepoRoot; Harness = $Harness; Case = $Case }
            return $sandboxCase
        }
        $buildScript = Join-Path $TestDrive 'build.ps1'
        $repoRoot = $TestDrive
        $previousCiRunNumber = [Environment]::GetEnvironmentVariable('GITHUB_RUN_NUMBER', 'Process')
        [Environment]::SetEnvironmentVariable('GITHUB_RUN_NUMBER', '64', 'Process')
        $script:capturedBuildArguments = @()
        $script:capturedBuildEnvironment = $null
        Invoke-Expression $wrapperAst.Extent.Text
        # This fixture executes the real deployment wrapper with explicit policy
        # data, independent of modules or SDK state loaded by an outer Full run.
        function Get-RSBuildScriptArguments {
            param([string]$Suite,[string]$Configuration,[string]$Platform)
            if($Suite -cne 'Full' -or $Configuration -cne 'Debug' -or $Platform -cne 'x64'){throw 'Fixture build policy received a different request.'}
            return @{MaxCpuCount=2}
        }
        function Get-RSBuildEnvironmentOverrides {
            param([string]$Suite)
            if($Suite -cne 'Full'){throw 'Fixture environment policy received a different suite.'}
            return @{RS_DEPLOYMENT_POLICY_FIXTURE='preserved'}
        }
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

            $script:capturedBuildArguments = @($Arguments)
            $script:capturedBuildEnvironment = $AdditionalEnvironment
            throw [System.TimeoutException]::new('simulated streaming timeout')
        }

        $capturedException = $null
        try {
            Invoke-RSTargetedBuildForDeploymentTest -Platform x64 -Configuration Debug -BuildNumber 533
        }
        catch {
            $capturedException = $_.Exception
        }
        finally {
            [Environment]::SetEnvironmentVariable('GITHUB_RUN_NUMBER', $previousCiRunNumber, 'Process')
            Remove-Item Function:\Invoke-RSTargetedBuildForDeploymentTest -ErrorAction SilentlyContinue
            Remove-Item Function:\Invoke-RSStreamingProcess -ErrorAction SilentlyContinue
            Remove-Item Function:\New-RSTestSandboxScratchDirectory -ErrorAction SilentlyContinue
            Remove-Item Function:\Get-RSBuildScriptArguments -ErrorAction SilentlyContinue
            Remove-Item Function:\Get-RSBuildEnvironmentOverrides -ErrorAction SilentlyContinue
        }

        $null -eq $capturedException | Should Be $false
        $capturedException.Message | Should Match 'simulated streaming timeout'
        $numberIndex = [array]::IndexOf($script:capturedBuildArguments, '-BuildNumber')
        $numberIndex | Should BeGreaterThan -1
        $script:capturedBuildArguments[$numberIndex + 1] | Should Be '533'
        $cpuIndex=[array]::IndexOf($script:capturedBuildArguments,'-MaxCpuCount')
        $cpuIndex | Should BeGreaterThan -1
        $script:capturedBuildArguments[$cpuIndex+1] | Should Be '2'
        $script:capturedBuildEnvironment.RS_DEPLOYMENT_POLICY_FIXTURE | Should Be 'preserved'
        [System.IO.Path]::GetFullPath([string]$script:scratchAllocation.RepoRoot) |
            Should Be ([System.IO.Path]::GetFullPath([string]$TestDrive))
        $script:scratchAllocation.Harness | Should Be 'pester-targeted-plugin-deployment'
        $script:scratchAllocation.Case | Should Be 'build-capture'
        $captureRoot = [System.IO.Path]::GetFullPath($sandboxCase).TrimEnd('\') + '\'
        foreach ($capturePath in @(
                [regex]::Match($capturedException.Message, 'stdout: (?<path>.+?\.out\.log)').Groups['path'].Value,
                [regex]::Match($capturedException.Message, 'stderr: (?<path>.+?\.err\.log)').Groups['path'].Value)) {
            [System.IO.Path]::GetFullPath($capturePath).StartsWith($captureRoot, [System.StringComparison]::OrdinalIgnoreCase) | Should Be $true
        }
        $capturedException.InnerException.GetType().FullName | Should Be 'System.TimeoutException'
        $capturedException.InnerException.Message | Should Be 'simulated streaming timeout'
    }

    It 'keeps targeted plugin deployment on the bounded shared streaming contract' {
        $deploymentSource = Get-Content `
            -LiteralPath (Join-Path $repoRoot 'Tools\Tests\RedSalamanderPluginDeployment.Tests.ps1') `
            -Raw

        $deploymentSource | Should Match 'Invoke-RSStreamingProcess[\s\S]*?-DelegateArtifactOperation'
        $deploymentSource | Should Match '-LogPath \$stdoutLog[\s\S]*?-ErrorLogPath \$stderrLog'
        $deploymentSource | Should Match '-TimeoutMilliseconds \(60 \* 60 \* 1000\)'
        $deploymentSource | Should Match '-OutputDrainTimeoutMilliseconds 5000'
        $deploymentSource | Should Match 'throw \[System\.InvalidOperationException\]::new\(\$message, \$originalException\)'
        $deploymentSource | Should Not Match 'ReadToEndAsync|\.Result'
    }
}

Describe 'Background build console policy' {
    BeforeAll {
        Import-Module $artifactOperationLockModule -ErrorAction Stop
        $script:backgroundStreamingModule = Import-Module $helperModule -Force -PassThru -ErrorAction Stop
        $script:backgroundStreamingInvoke = $script:backgroundStreamingModule.ExportedFunctions['Invoke-RSStreamingProcess']
        $script:backgroundSanitizedModule = Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/SanitizedEnvironment.psm1') -PassThru -ErrorAction Stop
    }

    It 'uses the delegated captured route at <Path>' -TestCases @(
        @{ Path = 'vcpkg-install.ps1'; FileVariable = 'vcpkgExePath' },
        @{ Path = 'Tools/Invoke-SanitizedMsbuild.ps1'; FileVariable = 'resolvedMsbuildPath' }
    ) {
        param($Path, $FileVariable)
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot $Path), [ref]$tokens, [ref]$errors)
        $errors.Count | Should Be 0
        $calls = @($ast.FindAll({ param($node)
            $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -ceq 'Invoke-RSStreamingProcess'
        }, $true))
        $calls.Count | Should Be 1
        $script:backgroundRecordedLaunch = $null
        Mock Invoke-RSStreamingProcess {
            param($FilePath, $Arguments, $WorkingDirectory, [switch]$DelegateArtifactOperation)
            $script:backgroundRecordedLaunch = @{ FilePath=$FilePath; Arguments=@($Arguments); WorkingDirectory=$WorkingDirectory; Delegated=[bool]$DelegateArtifactOperation }
            return 17
        }
        $vcpkgExePath = 'fixed-installer-fixture.exe'
        $resolvedMsbuildPath = 'fixed-msbuild-fixture.exe'
        $vcpkgArguments = [Collections.Generic.List[string]]::new()
        $vcpkgArguments.Add('install'); $vcpkgArguments.Add('literal value')
        $MSBuildArguments = @('fixture.vcxproj','/p:Configuration=Debug','/p:Platform=x64')
        $exitCode = & ([scriptblock]::Create($calls[0].Extent.Text))
        $exitCode | Should Be 17
        $script:backgroundRecordedLaunch.FilePath | Should Be (Get-Variable -Name $FileVariable -ValueOnly)
        $script:backgroundRecordedLaunch.WorkingDirectory | Should Be $repoRoot
        $script:backgroundRecordedLaunch.Delegated | Should Be $true
        $expected = if ($FileVariable -ceq 'vcpkgExePath') { $vcpkgArguments.ToArray() } else { $MSBuildArguments }
        ($script:backgroundRecordedLaunch.Arguments -join '|') | Should Be ($expected -join '|')
    }

    It 'preflights the actual PSI then keeps a background child windowless captured and in its owned job' {
        Mock Start-RSContainedProcess -ModuleName ProcessStreaming {
            param($ProcessStartInfo, [switch]$DelegateArtifactOperation)
            # Refuse before the real starter: never exercise a console-enabled regression
            # from a console-free parent on the local desktop.
            if ($ProcessStartInfo.UseShellExecute -or -not $ProcessStartInfo.CreateNoWindow -or
                -not $ProcessStartInfo.RedirectStandardOutput -or -not $ProcessStartInfo.RedirectStandardError) {
                throw 'Refusing unsafe console-policy fixture launch.'
            }
            # Pester's module mock has a different session state. Resolve only the exact
            # canonical path already carried in this fixture's literal child arguments.
            $arguments = @($ProcessStartInfo.ArgumentList)
            if ($arguments.Count -ne 5 -or $arguments[0] -cne '-NoProfile' -or $arguments[1] -cne '-File' -or
                $arguments[3] -cne '-SanitizedModule') { throw 'Unexpected safe-child argument shape.' }
            $canonical = Import-Module $arguments[4] -PassThru -ErrorAction Stop
            if ($canonical.Path -cne $arguments[4] -or -not $canonical.ExportedFunctions.ContainsKey('Start-RSContainedProcess')) {
                throw 'The exact canonical contained starter is unavailable.'
            }
            $starter = $canonical.ExportedFunctions['Start-RSContainedProcess']
            $owned = & $starter -ProcessStartInfo $ProcessStartInfo -DelegateArtifactOperation:$DelegateArtifactOperation
            $transferred = $false
            try {
                [ordered]@{ process=$owned.Id; exact_job=$owned.IsInContainmentJob; delegated=[bool]$DelegateArtifactOperation;
                    no_window=$ProcessStartInfo.CreateNoWindow; stdout_redirected=$ProcessStartInfo.RedirectStandardOutput;
                    stderr_redirected=$ProcessStartInfo.RedirectStandardError } |
                    ConvertTo-Json -Compress | Set-Content -LiteralPath ([IO.Path]::ChangeExtension($arguments[2], '.launch.json')) -Encoding UTF8
                $transferred = $true
                return $owned
            } finally {
                if (-not $transferred) { $owned.Dispose() }
            }
        }
        $child = Join-Path $TestDrive 'ConsoleFreeBuildObservation.ps1'
        @'
param([string]$SanitizedModule)
Import-Module $SanitizedModule -ErrorAction Stop
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
// Fixture-local, read-only console/pipe observations: no shared PowerShell helper owns these APIs.
// Production launch/containment still comes from the canonical module; this type never creates UI or a process.
public static class BackgroundBuildConsoleObservation {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("kernel32.dll", SetLastError=true)] public static extern uint GetConsoleProcessList(uint[] ids, uint count);
    [DllImport("kernel32.dll")] public static extern IntPtr GetStdHandle(int id);
    [DllImport("kernel32.dll")] public static extern uint GetFileType(IntPtr handle);
}
"@
$ids = [uint32[]]::new(1)
$consoleCount = [BackgroundBuildConsoleObservation]::GetConsoleProcessList($ids, 1)
$consoleError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
[ordered]@{
    process=$PID
    parent=[RedSalamander.Tooling.KillOnCloseJob]::GetCurrentParentProcessId()
    started_utc=(Get-Process -Id $PID).StartTime.ToUniversalTime().ToString('o')
    console_window=[BackgroundBuildConsoleObservation]::GetConsoleWindow().ToInt64()
    console_process_count=$consoleCount
    console_query_error=$consoleError
    stdout_type=[BackgroundBuildConsoleObservation]::GetFileType([BackgroundBuildConsoleObservation]::GetStdHandle(-11))
    stderr_type=[BackgroundBuildConsoleObservation]::GetFileType([BackgroundBuildConsoleObservation]::GetStdHandle(-12))
    any_job=[RedSalamander.Tooling.KillOnCloseJob]::IsCurrentProcessInJob()
} | ConvertTo-Json -Compress | Write-Output
Write-Output 'fixed-build-stdout'
[Console]::Error.WriteLine('fixed-build-stderr')
exit 17
'@ | Set-Content -LiteralPath $child -Encoding UTF8
        $rows = [Collections.Generic.List[object]]::new()
        $exitCode = & $script:backgroundStreamingInvoke -FilePath (Get-Process -Id $PID).Path -Arguments @(
            '-NoProfile','-File',$child,'-SanitizedModule',$script:backgroundSanitizedModule.Path
        ) -WorkingDirectory $TestDrive -DelegateArtifactOperation -TimeoutMilliseconds 15000 -OutputLineCallback {
            param($Line, $IsError)
            $rows.Add([pscustomobject]@{ Line=$Line; IsError=$IsError })
        }
        $exitCode | Should Be 17
        $global:LASTEXITCODE | Should Be 17
        $launch = Get-Content -LiteralPath ([IO.Path]::ChangeExtension($child, '.launch.json')) -Raw | ConvertFrom-Json
        $launch.delegated | Should Be $true
        $launch.exact_job | Should Be $true
        @($rows | Where-Object { $_.Line -ceq 'fixed-build-stdout' -and -not $_.IsError }).Count | Should Be 1
        @($rows | Where-Object { $_.Line -ceq 'fixed-build-stderr' -and $_.IsError }).Count | Should Be 1
        $observation = @($rows | Where-Object { $_.Line.StartsWith('{"process":') })
        $observation.Count | Should Be 1
        $facts = $observation[0].Line | ConvertFrom-Json
        $facts.process | Should Be $launch.process
        $facts.parent | Should Be $PID
        [string]::IsNullOrWhiteSpace($facts.started_utc) | Should Be $false
        $facts.console_window | Should Be 0
        # CREATE_NO_WINDOW suppresses a console window/default Terminal UX, not the
        # inbox console server association. Keep process-list facts diagnostic.
        # https://github.com/microsoft/terminal/blob/main/doc/specs/%23492%20-%20Default%20Terminal/spec.md#inbox-console
        Write-Host ('Background build child observation: ' + ($facts | ConvertTo-Json -Compress))
        Write-Host ('Background build launch observation: ' + ($launch | ConvertTo-Json -Compress))
        $facts.stdout_type | Should Be 3
        $facts.stderr_type | Should Be 3
        $facts.any_job | Should Be $true
    }
}

Describe 'Captured VSwhere metadata family' {
    BeforeAll {
        $script:metadataSourceRoot=$repoRoot
        $script:metadataStreaming=Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/ProcessStreaming.psm1') -Force -PassThru
        $script:metadataRealCapture=$script:metadataStreaming.ExportedFunctions['Invoke-RSCapturedProcess']
        function Get-MetadataFunctionText([string]$Relative,[string]$Name) {
            $tokens=$null;$errors=$null
            $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $script:metadataSourceRoot $Relative),[ref]$tokens,[ref]$errors)
            if($errors.Count){throw 'Metadata source does not parse.'}
            $found=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq $Name},$true))
            if($found.Count -ne 1){throw 'Expected one actual metadata function.'};return $found[0].Extent.Text
        }
        function Assert-MetadataCapturedCall([string]$Text) {
            $tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseInput($Text,[ref]$tokens,[ref]$errors)
            $commands=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst]},$true))
            if($errors.Count -or @($commands|Where-Object {$_.InvocationOperator -eq 'Ampersand' -and $_.Extent.Text -match '&\s*\$vswhere\b'}).Count -or
                @($commands|Where-Object {$_.GetCommandName() -ceq 'Invoke-RSCapturedProcess'}).Count -ne 1){throw 'Refusing a bare or unreviewed metadata caller before execution.'}
        }
        $script:metadataBuild=Get-MetadataFunctionText 'build.ps1' 'Find-MSBuild'
        $script:metadataVcpkg=Get-MetadataFunctionText 'Tools/Modules/Build/VcpkgInstallSafety.psm1' 'Get-RSVcpkgVisualStudioInstallation'
        $restore=[IO.File]::ReadAllText((Join-Path $repoRoot 'Tools/Restore-DxUi.ps1')).Replace("`r`n","`n")
        $begin=$restore.IndexOf('$vswhere =');$end=$restore.IndexOf('$msbuildRelative =')
        if($begin -lt 0 -or $end -le $begin){throw 'Expected the actual standalone discovery block.'}
        $script:metadataRestore=$restore.Substring($begin,$end-$begin)
        foreach($text in @($script:metadataBuild,$script:metadataVcpkg,$script:metadataRestore)){Assert-MetadataCapturedCall $text}
    }
    BeforeEach {
        # The canonical decoder inherits Console.OutputEncoding; exercise its explicit UTF-8 caller contract.
        $script:metadataSavedOutputEncoding=[Console]::OutputEncoding
        [Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
        $script:metadataSavedEnvironment=@{}
        foreach($name in @('ProgramFiles','ProgramFiles(x86)','GITHUB_ACTIONS','MSBUILD_EXE_PATH','RS_METADATA_PAYLOAD','RS_METADATA_EXIT','RS_METADATA_OBSERVATION','TEMP','TMP')){
            $script:metadataSavedEnvironment[$name]=[Environment]::GetEnvironmentVariable($name,'Process')
        }
        $script:metadataRoot=Join-Path $TestDrive ([guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($script:metadataRoot)
        $script:metadataProgramFiles=Join-Path $script:metadataRoot 'Program files é雪'
        $script:metadataInstallation=Join-Path $script:metadataProgramFiles 'Microsoft Visual Studio/2026/Enterprise'
        foreach($relative in @('MSBuild/Current/Bin/amd64/MSBuild.exe','MSBuild/Current/Bin/arm64/MSBuild.exe','MSBuild/Current/Bin/MSBuild.exe')){
            $file=Join-Path $script:metadataInstallation $relative;[void][IO.Directory]::CreateDirectory((Split-Path $file));[IO.File]::WriteAllText($file,'data-only marker; never executed')
        }
        $vswhere=Join-Path $script:metadataProgramFiles 'Microsoft Visual Studio/Installer/vswhere.exe'
        [void][IO.Directory]::CreateDirectory((Split-Path $vswhere));[IO.File]::WriteAllText($vswhere,'data-only marker; never executed')
        $script:metadataProbe=Join-Path $script:metadataRoot 'Safe metadata child.ps1'
        $probeLines=@(
            '[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)',
            'Add-Type -TypeDefinition ''using System; using System.Runtime.InteropServices; public static class MetadataConsoleProbe { [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow(); [DllImport("user32.dll")] [return:MarshalAs(UnmanagedType.Bool)] public static extern bool IsWindowVisible(IntPtr h); }''',
            '$console=[MetadataConsoleProbe]::GetConsoleWindow()',
            '[ordered]@{pid=$PID;argv=@($args);cwd=(Get-Location).Path;visibleConsole=($console -ne [IntPtr]::Zero -and [MetadataConsoleProbe]::IsWindowVisible($console));endpoint=$env:_MSPDBSRV_ENDPOINT_}|ConvertTo-Json -Depth 4|Set-Content -LiteralPath $env:RS_METADATA_OBSERVATION -Encoding utf8',
            '[Console]::Out.WriteLine([IO.File]::ReadAllText($env:RS_METADATA_PAYLOAD))',
            '[Console]::Error.WriteLine(''diagnostic stderr é雪; never metadata'')',
            'exit ([int]$env:RS_METADATA_EXIT)')
        [IO.File]::WriteAllText($script:metadataProbe,($probeLines -join "`n"),[Text.UTF8Encoding]::new($false))
        foreach($name in @('ProgramFiles','ProgramFiles(x86)')){[Environment]::SetEnvironmentVariable($name,$script:metadataProgramFiles,'Process')}
        foreach($name in @('GITHUB_ACTIONS','MSBUILD_EXE_PATH')){[Environment]::SetEnvironmentVariable($name,$null,'Process')}
        foreach($name in @('TEMP','TMP')){[Environment]::SetEnvironmentVariable($name,$script:metadataRoot,'Process')}
        $env:RS_METADATA_PAYLOAD=Join-Path $script:metadataRoot 'payload.txt';$env:RS_METADATA_EXIT='0';$env:RS_METADATA_OBSERVATION=Join-Path $script:metadataRoot 'observed.json'
        $starterPath=Join-Path $script:metadataSourceRoot 'Tools/Modules/Build/SanitizedEnvironment.psm1'
        $starters=@(Get-Module -All|Where-Object {$_.Path -ceq $starterPath});if($starters.Count -ne 1){throw 'Expected one actual canonical starter module.'}
        $script:metadataAudit=[pscustomobject]@{Rows=[Collections.Generic.List[object]]::new();Starter=$starters[0].ExportedFunctions['Start-RSContainedProcess'];Active=$true}
        $starterExport=$script:metadataAudit.Starter
        if($starterExport -isnot [Management.Automation.FunctionInfo] -or -not [object]::ReferenceEquals($starterExport.Module,$starters[0])){throw 'Canonical starter export ownership mismatch.'}
        $state=$script:metadataAudit
        Mock Start-RSContainedProcess -ModuleName ProcessStreaming ({
            param($ProcessStartInfo,[switch]$DelegateArtifactOperation)
            if(-not $state.Active -or $ProcessStartInfo.UseShellExecute -or -not $ProcessStartInfo.CreateNoWindow -or
                -not $ProcessStartInfo.RedirectStandardOutput -or -not $ProcessStartInfo.RedirectStandardError -or $DelegateArtifactOperation){throw 'Refusing unsafe metadata start before child creation.'}
            $starter=$state.Starter;$owned=& $starter -ProcessStartInfo $ProcessStartInfo;$transferred=$false
            try {
                $state.Rows.Add([ordered]@{pid=$owned.Id;contained=$owned.IsInContainmentJob;file=$ProcessStartInfo.FileName;cwd=$ProcessStartInfo.WorkingDirectory;
                    argv=@($ProcessStartInfo.ArgumentList);noWindow=$ProcessStartInfo.CreateNoWindow;stdout=$ProcessStartInfo.RedirectStandardOutput;stderr=$ProcessStartInfo.RedirectStandardError;delegated=[bool]$DelegateArtifactOperation;stdoutEncoding=$(if($ProcessStartInfo.StandardOutputEncoding){$ProcessStartInfo.StandardOutputEncoding.CodePage}else{$null});stderrEncoding=$(if($ProcessStartInfo.StandardErrorEncoding){$ProcessStartInfo.StandardErrorEncoding.CodePage}else{$null})})
                $transferred=$true;return $owned
            }finally{if(-not $transferred){$owned.Dispose()}}
        }.GetNewClosure())
    }
    AfterEach {
        [Console]::OutputEncoding=$script:metadataSavedOutputEncoding
        Write-Host ('[METADATA OWNED STARTS] '+($script:metadataAudit.Rows|ConvertTo-Json -Depth 5 -Compress -EscapeHandling EscapeNonAscii))
        $script:metadataAudit.Active=$false
        foreach($name in $script:metadataSavedEnvironment.Keys){
            if($null -eq $script:metadataSavedEnvironment[$name]){ Remove-Item -LiteralPath ('Env:\'+$name) -ErrorAction SilentlyContinue }
            else{ [Environment]::SetEnvironmentVariable($name,$script:metadataSavedEnvironment[$name],'Process') }
            if([Environment]::GetEnvironmentVariable($name,'Process') -cne $script:metadataSavedEnvironment[$name]){throw ("Metadata fixture environment restoration failed: $name expectedNull=$($null -eq $script:metadataSavedEnvironment[$name]) actualNull=$($null -eq [Environment]::GetEnvironmentVariable($name,'Process'))" )}
        }
        Write-Host '[METADATA ENVIRONMENT] restored'
    }
    It 'captures a real owned Unicode CUI child with literal arguments, separate streams, actual exits and cleanup' {
        $beforeEndpoint=[Environment]::GetEnvironmentVariable('_MSPDBSRV_ENDPOINT_','Process')
        foreach($exitCode in @(0,17)){
            [IO.File]::WriteAllText($env:RS_METADATA_PAYLOAD,'{"value":"metadata é雪"}',[Text.UTF8Encoding]::new($false));$env:RS_METADATA_EXIT=[string]$exitCode
            $capture=& $script:metadataRealCapture -FilePath (Join-Path $PSHOME 'pwsh.exe') -Arguments @('-NoLogo','-NoProfile','-File',$script:metadataProbe,'literal space','*') -WorkingDirectory $script:metadataRoot
            $capture.ExitCode|Should Be $exitCode;(($capture.StdOut -join "`n")|ConvertFrom-Json).value|Should Be 'metadata é雪'
            ($capture.StdErr -join '|')|Should Be 'diagnostic stderr é雪; never metadata'
            $observed=Get-Content -LiteralPath $env:RS_METADATA_OBSERVATION -Raw|ConvertFrom-Json
            ($observed.argv -join '|')|Should Be 'literal space|*';$observed.cwd|Should Be $script:metadataRoot;$observed.visibleConsole|Should Be $false
            Write-Host ('[METADATA CHILD] '+([ordered]@{exit=$capture.ExitCode;stdout=$capture.StdOut;stderr=$capture.StdErr;observation=$observed}|ConvertTo-Json -Depth 5 -Compress -EscapeHandling EscapeNonAscii))
            $row=$script:metadataAudit.Rows[$script:metadataAudit.Rows.Count-1];$row.pid|Should Be $observed.pid;$row.contained|Should Be $true;$row.stdoutEncoding|Should BeNullOrEmpty;$row.stderrEncoding|Should BeNullOrEmpty
            (Get-Process -Id $row.pid -ErrorAction SilentlyContinue)|Should BeNullOrEmpty
        }
        $script:metadataAudit.Rows.Count|Should Be 2
        [Environment]::GetEnvironmentVariable('_MSPDBSRV_ENDPOINT_','Process')|Should Be $beforeEndpoint
        $unsafe=[Diagnostics.ProcessStartInfo]::new();$unsafe.FileName=Join-Path $PSHOME 'pwsh.exe'
        {& $script:metadataStreaming {param($psi)Start-RSContainedProcess -ProcessStartInfo $psi} $unsafe}|Should Throw 'Refusing unsafe metadata start'
        $script:metadataAudit.Rows.Count|Should Be 2
    }
    It 'runs all three actual discovery callers through capture and retains selection and exit policies' {
        $calls=[Collections.Generic.List[object]]::new();$real=$script:metadataRealCapture;$probe=$script:metadataProbe;$cwd=$script:metadataRoot
        $installation=$script:metadataInstallation;$build=$script:metadataBuild;$vcpkg=$script:metadataVcpkg;$restore=$script:metadataRestore
        $run={param($kind,$source,$exitCode)
            $env:RS_METADATA_EXIT=[string]$exitCode
            $payload=if($kind -eq 'build'){@(@{installationPath=$installation;installationVersion='18.0.1';displayName='Synthetic VS'})|ConvertTo-Json -Compress}else{$installation}
            [IO.File]::WriteAllText($env:RS_METADATA_PAYLOAD,$payload,[Text.UTF8Encoding]::new($false))
            & {
                function Invoke-RSCapturedProcess {param($FilePath,$Arguments,$WorkingDirectory)
                    $calls.Add([ordered]@{kind=$kind;file=$FilePath;argv=@($Arguments);cwd=$WorkingDirectory})
                    return & $real -FilePath (Join-Path $PSHOME 'pwsh.exe') -Arguments (@('-NoLogo','-NoProfile','-File',$probe)+$Arguments) -WorkingDirectory $WorkingDirectory
                }
                function Get-RSWorkflowSelectedMSBuild{return $null}
                function Get-RSHostOrderedMSBuildRelativePaths{'MSBuild/Current/Bin/amd64/MSBuild.exe';'MSBuild/Current/Bin/arm64/MSBuild.exe';'MSBuild/Current/Bin/MSBuild.exe'}
                if($kind -eq 'restore'){. ([scriptblock]::Create($source));return $installation}
                . ([scriptblock]::Create($source))
                if($kind -eq 'build'){return Find-MSBuild};return Get-RSVcpkgVisualStudioInstallation
            }
        }.GetNewClosure()
        Push-Location $cwd
        try {
            foreach($kind in @('restore','vcpkg','build')){
                $source=switch($kind){restore{$restore}vcpkg{$vcpkg}build{$build}}
                $result=& $run $kind $source 0
                if($kind -eq 'build'){$result.Method|Should Be 'vswhere';$result.Path|Should Be (Join-Path $installation 'MSBuild/Current/Bin/amd64/MSBuild.exe')}else{$result|Should Be $installation}
                $last=$calls[$calls.Count-1];$last.cwd|Should Be $cwd
                $expected=if($kind -eq 'build'){'-all|-products|*|-prerelease|-format|json'}else{'-latest|-prerelease|-products|*|-requires|Microsoft.Component.MSBuild|-property|installationPath'}
                ($last.argv -join '|')|Should Be $expected
                $observed=Get-Content -LiteralPath $env:RS_METADATA_OBSERVATION -Raw|ConvertFrom-Json
                ($observed.argv -join '|')|Should Be $expected;$observed.visibleConsole|Should Be $false
                Write-Host ('[METADATA CALLER] '+([ordered]@{kind=$kind;argv=$last.argv;observation=$observed}|ConvertTo-Json -Depth 5 -Compress -EscapeHandling EscapeNonAscii))
                if($kind -eq 'build'){$fallback=& $run $kind $source 17;$fallback.Method|Should Not Be 'vswhere'}else{{& $run $kind $source 17}|Should Throw 'Visual Studio discovery failed (17)'}
                $failureObserved=Get-Content -LiteralPath $env:RS_METADATA_OBSERVATION -Raw|ConvertFrom-Json
                $failureObserved.visibleConsole|Should Be $false;($failureObserved.argv -join '|')|Should Be $expected
                $failureObserved.pid|Should Be $script:metadataAudit.Rows[$script:metadataAudit.Rows.Count-1].pid
                Write-Host ('[METADATA CALLER FAILURE] '+([ordered]@{kind=$kind;exit=17;observation=$failureObserved}|ConvertTo-Json -Depth 5 -Compress -EscapeHandling EscapeNonAscii))
                $unsafe=$source.Replace('Invoke-RSCapturedProcess -FilePath $vswhere','& $vswhere');($unsafe -cne $source)|Should Be $true
                {Assert-MetadataCapturedCall $unsafe}|Should Throw 'Refusing a bare or unreviewed metadata caller'
            }
            $calls.Count|Should Be 6;$script:metadataAudit.Rows.Count|Should Be 6
            @($script:metadataAudit.Rows|Where-Object {-not $_.contained -or -not $_.noWindow -or -not $_.stdout -or -not $_.stderr -or $_.delegated}).Count|Should Be 0
            foreach($row in $script:metadataAudit.Rows){(Get-Process -Id $row.pid -ErrorAction SilentlyContinue)|Should BeNullOrEmpty}
        }finally{Pop-Location}
    }
    It 'runs each actual graph and HEAD metadata caller with owned capture and refuses failed queries' {
        function Get-GitCallerStatements($Assignment) {
            $siblings=@($Assignment.Parent.Statements);$index=-1
            for($n=0;$n -lt $siblings.Count;$n++){if([object]::ReferenceEquals($siblings[$n],$Assignment)){$index=$n;break}}
            if($index -lt 0 -or $index+2 -ge $siblings.Count){throw 'Git caller statement boundary unavailable.'}
            return (($siblings[$index..($index+2)]|ForEach-Object {$_.Extent.Text}) -join "`n")
        }
        function Assert-GitCallerCapture([string]$Source) {
            $t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseInput($Source,[ref]$t,[ref]$e)
            $commands=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst]},$true))
            if($e.Count -or @($commands|Where-Object {$_.GetCommandName() -ceq 'git'}).Count -or
                @($commands|Where-Object {$_.GetCommandName() -ceq 'Invoke-RSCapturedProcess'}).Count -ne 1){throw 'Refusing an uncaptured Git caller before execution.'}
        }
        $t=$null;$e=$null;$runner=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'Tools/Run-AllTests.ps1'),[ref]$t,[ref]$e)
        $graph=@($runner.FindAll({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -ceq '$graphInputsCapture'},$true))
        $graph.Count|Should Be 3
        $cases=[Collections.Generic.List[object]]::new()
        foreach($assignment in $graph){$statements=@($assignment.Parent.Statements);$source=Get-GitCallerStatements $assignment
            $target=if($source.Contains('$prebuildTrackedInputs =')){'prebuildTrackedInputs'}else{'trackedGraphInputs'}
            $cases.Add(@{kind='graph';source=$source;target=$target})
        }
        $hostText=Get-MetadataFunctionText 'Tools/Modules/Testing/HostedTestDesktop.psm1' 'Invoke-RSHostedArm64StockPreparation'
        $hostAst=[Management.Automation.Language.Parser]::ParseInput($hostText,[ref]$t,[ref]$e)
        $hostAssignment=$hostAst.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -ceq '$headCapture'},$true)
        $cases.Add(@{kind='host-head';source=(Get-GitCallerStatements $hostAssignment);target='sourceHead'})
        $identity=Get-MetadataFunctionText 'Tools/Modules/Build/VcpkgToolIdentity.psm1' 'Assert-RSVcpkgToolIdentity'
        $cases.Add(@{kind='vcpkg-head';source=$identity;target=''})
        $vcpkg=Join-Path $script:metadataRoot 'vcpkg.exe';[IO.File]::WriteAllText($vcpkg,'fixture identity marker; never executed')
        [void][IO.Directory]::CreateDirectory((Join-Path $script:metadataRoot '.git'))
        $calls=[Collections.Generic.List[object]]::new();$real=$script:metadataRealCapture;$probe=$script:metadataProbe;$cwd=$script:metadataRoot
        $run={param($case,$code)
            $env:RS_METADATA_EXIT=[string]$code
            $payload=if($case.kind -eq 'graph'){"Graph é雪.sln`nliteral*.props"}else{'a'*40}
            [IO.File]::WriteAllText($env:RS_METADATA_PAYLOAD,$payload,[Text.UTF8Encoding]::new($false))
            & {
                Set-StrictMode -Version Latest
                function Invoke-RSCapturedProcess {param($FilePath,$Arguments,$WorkingDirectory,[Text.Encoding]$StandardOutputEncoding)
                    if ($null -eq $StandardOutputEncoding -or $StandardOutputEncoding.CodePage -ne 65001) { throw 'Known Git stdout requires per-child UTF8.' }
                    if ($FilePath -cne 'git' -or $WorkingDirectory -cne $cwd) { throw 'Git caller lost literal executable or CWD.' }
                    $calls.Add(@{kind=$case.kind;argv=@($Arguments);cwd=$WorkingDirectory})
                    return & $real -FilePath (Join-Path $PSHOME 'pwsh.exe') -Arguments (@('-NoLogo','-NoProfile','-File',$probe)+$Arguments) -WorkingDirectory $WorkingDirectory -StandardOutputEncoding $StandardOutputEncoding
                }
                function Exit-RSInvalidRunnerInvocation {param($Message)throw $Message}
                $repoRoot=$cwd
                . ([scriptblock]::Create($case.source))
                if($case.kind -eq 'vcpkg-head'){return Assert-RSVcpkgToolIdentity -VcpkgExe $vcpkg -ExpectedCommit ('a'*40)}
                return ,(Get-Variable -Name $case.target -ValueOnly)
            }
        }.GetNewClosure()
        foreach($case in $cases){
            Assert-GitCallerCapture $case.source
            $result=& $run $case 0
            if($case.kind -eq 'graph'){($result -join '|')|Should Be 'Graph é雪.sln|literal*.props'}
            elseif($case.kind -eq 'host-head'){$result|Should Be ('a'*40)}else{$result|Should Be $vcpkg}
            $observed=Get-Content $env:RS_METADATA_OBSERVATION -Raw|ConvertFrom-Json
            $observed.visibleConsole|Should Be $false;$observed.pid|Should Be $script:metadataAudit.Rows[$script:metadataAudit.Rows.Count-1].pid
            $expected=if($case.kind -eq 'graph'){'-C|'+$cwd+'|ls-files|*.sln|*.slnx|*.vcxproj|*.props|*.targets|RuntimeDependencies.props'}else{'-C|'+$cwd+'|rev-parse|HEAD'}
            ($calls[$calls.Count-1].argv -join '|')|Should Be $expected;($observed.argv -join '|')|Should Be $expected
            $failure=if($case.kind -eq 'graph'){'Cannot enumerate tracked validation build-graph inputs.'}elseif($case.kind -eq 'host-head'){'Cannot bind stock preparation to the exact checkout head.'}else{'Unable to read the vcpkg Git identity'}
            {& $run $case 17}|Should Throw $failure
            $unsafe=$case.source.Replace('Invoke-RSCapturedProcess','& git')
            ($unsafe -cne $case.source)|Should Be $true
            {Assert-GitCallerCapture $unsafe}|Should Throw 'Refusing an uncaptured Git caller'
        }
        $calls.Count|Should Be 10;$script:metadataAudit.Rows.Count|Should Be 10
        foreach($row in $script:metadataAudit.Rows){$row.contained|Should Be $true;$row.delegated|Should Be $false;(Get-Process -Id $row.pid -ErrorAction SilentlyContinue)|Should BeNullOrEmpty}
    }
}