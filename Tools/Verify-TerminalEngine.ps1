<#
.SYNOPSIS
Verifies one built lane against an explicit Terminal-engine lock.

.DESCRIPTION
Compatibility entrypoint for generic Terminal qualification and upgrade work. It binds the build manifest and output identities to the supplied lock, optionally runs locked smoke hooks, and writes a verification manifest. The normal product Ghostty runtime is verified by Build-GhosttyTerminalRuntime.ps1.

.PARAMETER Platform
Selects the x64 or ARM64 lane.

.PARAMETER Configuration
Selects Debug, Release, or ASan Debug.

.PARAMETER LockFile
Specifies the explicit Terminal-engine evaluation or candidate lock.

.PARAMETER OutputRoot
Selects the safe repository build root containing the built lane.

.PARAMETER RunSmoke
Runs each output's exact lock-declared smoke hook and verifies its output identities.

.PARAMETER Offline
Runs verification under attested host-wide network isolation.

.PARAMETER ExclusiveDisposableRunner
Attests disposable-runner ownership; required with -Offline.

.PARAMETER PassThruManifest
Returns the generated verification-manifest path.

.OUTPUTS
With -PassThruManifest, a string containing the manifest path. Otherwise no pipeline output. Writes terminal-engine-verify-manifest.json below the lane output root.

.NOTES
Prerequisites: a valid explicit lock and a matching completed Build-TerminalEngine lane. Smoke verification must execute on the requested native architecture. Side effects: may launch locked smoke processes and writes a verification manifest; offline mode creates and removes worker state. Exit is nonzero on lock, identity, architecture, isolation, smoke, or manifest mismatch. Primary consumers: External/TerminalEngine/TerminalEngine.proj and maintainer-operated upgrade qualification.

.EXAMPLE
.\Tools\Verify-TerminalEngine.ps1 -LockFile .\.build\candidate-lock.json -Platform x64 -Configuration Release -RunSmoke -PassThruManifest
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('x64', 'ARM64')]
    [string]$Platform,

    [Parameter(Mandatory)]
    [ValidateSet('Debug', 'Release', 'ASan Debug')]
    [string]$Configuration,

    [string]$LockFile = 'External\terminal-engine.lock.json',

    [string]$OutputRoot = '.build\TerminalEngine',

    [switch]$RunSmoke,

    [switch]$Offline,

    [switch]$ExclusiveDisposableRunner,

    [switch]$PassThruManifest,

    [Parameter(DontShow)]
    [switch]$InternalWorker,

    [Parameter(DontShow)]
    [string]$WorkerResultFile = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$lifecycleModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEngineLifecycle.psm1'
Import-Module $lifecycleModule -Force -ErrorAction Stop

function Invoke-RSVerificationWorker {
    param(
        [Parameter(Mandatory)]
        [object]$Lock,

        [Parameter(Mandatory)]
        [object]$Lane,

        [Parameter(Mandatory)]
        [string]$LaneOutputRoot,

        [switch]$RunSmoke
    )

    $buildManifestPath = Join-Path $LaneOutputRoot 'terminal-engine-build-manifest.json'
    if (-not (Test-Path -LiteralPath $buildManifestPath -PathType Leaf)) {
        throw "Build manifest is missing: $buildManifestPath"
    }
    $buildManifest = Get-Content -LiteralPath $buildManifestPath -Raw |
        ConvertFrom-Json -Depth 100 -ErrorAction Stop
    if ([string]$buildManifest.formatId -cne 'red-salamander-terminal-engine-build' -or
        [int]$buildManifest.schemaVersion -ne 1 -or
        [string]$buildManifest.lock.sha256 -cne $Lock.lockSha256 -or
        [string]$buildManifest.platform -cne $Platform -or
        [string]$buildManifest.configuration -cne $Configuration) {
        throw 'Build manifest is not bound to the requested lock and lane.'
    }

    $outputs = Test-RSTerminalOutputFiles `
        -Lock $Lock `
        -Lane $Lane `
        -LaneOutputRoot $LaneOutputRoot
    $runtimeRecords = [Collections.Generic.List[object]]::new()
    $smokeRecords = [Collections.Generic.List[object]]::new()
    $native = Get-RSTerminalNativeArchitecture
    $expectedNative = if ($Platform -eq 'x64') { 'X64' } else { 'Arm64' }

    foreach ($output in @(Get-RSTerminalLaneOutputs -Lock $Lock -Lane $Lane)) {
        $artifactPath = Resolve-RSTerminalPathInRoot `
            -Root $LaneOutputRoot `
            -Path ([string]$output.relativePath) `
            -MustExist
        $runtimeExpected = Get-RSObjectProperty `
            -Object $output `
            -Name runtimeIdentity `
            -Default $null
        if ($null -ne $runtimeExpected) {
            if ([string]$native.process -cne $expectedNative -or
                [string]$native.os -cne $expectedNative) {
                throw "Runtime identity requires a native $Platform process and OS."
            }
            $actualRuntime = Invoke-RSTerminalRuntimeIdentity `
                -Path $artifactPath `
                -Expected $runtimeExpected `
                -Abi $Lock.abi
            $runtimeRecords.Add([ordered]@{
                    outputId = [string]$output.id
                    identitySha256 = Get-RSCanonicalObjectSha256 -InputObject $actualRuntime
                    verified = $true
                })
        }

        if ($RunSmoke) {
            $smoke = Get-RSObjectProperty -Object $output -Name smoke -Default $null
            if ($null -eq $smoke) {
                throw "Output '$($output.id)' has no locked smoke hook."
            }
            $smokeResult = Invoke-RSTerminalSmokeHook `
                -Lock $Lock `
                -Smoke $smoke `
                -ArtifactPath $artifactPath
            $expectedStdout = Get-RSObjectProperty `
                -Object $smoke `
                -Name expectedStdoutSha256 `
                -Default $null
            $expectedStderr = Get-RSObjectProperty `
                -Object $smoke `
                -Name expectedStderrSha256 `
                -Default $null
            if ($null -ne $expectedStdout -and
                [string]$smokeResult.stdoutSha256 -cne [string]$expectedStdout) {
                throw "Output '$($output.id)' smoke stdout digest differs from the lock."
            }
            if ($null -ne $expectedStderr -and
                [string]$smokeResult.stderrSha256 -cne [string]$expectedStderr) {
                throw "Output '$($output.id)' smoke stderr digest differs from the lock."
            }
            $smokeRecords.Add([ordered]@{
                    outputId = [string]$output.id
                    exitCode = [int]$smokeResult.exitCode
                    stdoutSha256 = [string]$smokeResult.stdoutSha256
                    stderrSha256 = [string]$smokeResult.stderrSha256
                    verified = $true
                })
        }
    }

    return [pscustomobject]@{
        buildManifestSha256 = Get-RSSha256 -Path $buildManifestPath
        outputs = $outputs
        native = $native
        runtimeIdentity = $runtimeRecords.ToArray()
        smoke = $smokeRecords.ToArray()
    }
}

$lock = Read-RSTerminalEngineLock -LockFile $LockFile -ValidateInputs
$resolvedOutputRoot = Assert-RSTerminalSafeBuildRoot `
    -RepoRoot $repoRoot `
    -BuildRoot $OutputRoot
$lane = Get-RSTerminalLane `
    -Lock $lock `
    -Platform $Platform `
    -Configuration $Configuration
$laneOutputRoot = Join-Path (Join-Path $resolvedOutputRoot $Platform) $Configuration

if ($InternalWorker) {
    if ($Offline) {
        throw 'The internal offline worker must not recursively request isolation.'
    }
    if ([string]::IsNullOrWhiteSpace($WorkerResultFile)) {
        throw 'InternalWorker requires WorkerResultFile.'
    }
    $workerPath = [IO.Path]::GetFullPath($WorkerResultFile)
    [void](Assert-RSTerminalSafeBuildRoot -RepoRoot $repoRoot -BuildRoot (
            Split-Path -Parent $workerPath))
    $result = Invoke-RSVerificationWorker `
        -Lock $lock `
        -Lane $lane `
        -LaneOutputRoot $laneOutputRoot `
        -RunSmoke:$RunSmoke
    Write-RSTerminalJsonFile -Path $workerPath -Value $result
    exit 0
}

$isolationSummary = [pscustomobject]@{
    requested = $false
    provider = $null
    providerStatus = $null
    isolationVerified = $false
    descendantCoverage = $null
    rootStdoutSha256 = $null
    rootStderrSha256 = $null
}
if ($Offline) {
    if (-not $ExclusiveDisposableRunner) {
        throw '-Offline requires -ExclusiveDisposableRunner because isolation is host-wide.'
    }
    $isolationScript = Join-Path $repoRoot 'Tools\TerminalEngine\Invoke-RSIsolatedProcess.ps1'
    . $isolationScript -LoadOnly
    $workerDirectory = Join-Path $resolvedOutputRoot '.worker'
    [void](New-Item -ItemType Directory -Path $workerDirectory -Force)
    $workerResult = Join-Path $workerDirectory (
        "verify-$([Guid]::NewGuid().ToString('N')).json")
    $pwshPath = [IO.Path]::GetFullPath((Get-Process -Id $PID).Path)
    $arguments = [Collections.Generic.List[string]]::new()
    foreach ($argument in @(
            '-NoLogo',
            '-NoProfile',
            '-NonInteractive',
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            $PSCommandPath,
            '-Platform',
            $Platform,
            '-Configuration',
            $Configuration,
            '-LockFile',
            $lock.lockPath,
            '-OutputRoot',
            $resolvedOutputRoot,
            '-InternalWorker',
            '-WorkerResultFile',
            $workerResult)) {
        $arguments.Add($argument)
    }
    if ($RunSmoke) {
        $arguments.Add('-RunSmoke')
    }
    $attestation = Invoke-RSIsolatedProcess `
        -FilePath $pwshPath `
        -ArgumentList $arguments.ToArray() `
        -CanaryExecutablePath $pwshPath `
        -WorkingDirectory $repoRoot `
        -ExclusiveDisposableRunner
    if ($attestation.providerStatus -cne 'pass' -or
        -not $attestation.isolationVerified -or
        -not $attestation.rootCommandExecuted -or
        $attestation.rootCommand.result.exitCode -ne 0) {
        throw 'Runner-enforced offline terminal-engine verification failed.'
    }
    if (-not (Test-Path -LiteralPath $workerResult -PathType Leaf)) {
        throw 'Offline verification worker did not emit its exact result.'
    }
    $result = Get-Content -LiteralPath $workerResult -Raw |
        ConvertFrom-Json -Depth 100 -ErrorAction Stop
    Remove-Item -LiteralPath $workerResult -Force
    $isolationSummary = [pscustomobject]@{
        requested = $true
        provider = [string]$attestation.provider
        providerStatus = [string]$attestation.providerStatus
        isolationVerified = [bool]$attestation.isolationVerified
        descendantCoverage = [string]$attestation.scopeContract.descendantCoverage
        rootStdoutSha256 = [string]$attestation.rootCommand.result.stdoutSha256
        rootStderrSha256 = [string]$attestation.rootCommand.result.stderrSha256
    }
}
else {
    $result = Invoke-RSVerificationWorker `
        -Lock $lock `
        -Lane $lane `
        -LaneOutputRoot $laneOutputRoot `
        -RunSmoke:$RunSmoke
}

$manifestPath = Join-Path $laneOutputRoot 'terminal-engine-verify-manifest.json'
$manifest = [ordered]@{
    formatId = 'red-salamander-terminal-engine-verification'
    schemaVersion = 1
    createdUtc = [DateTime]::UtcNow.ToString(
        'o',
        [Globalization.CultureInfo]::InvariantCulture)
    lock = [ordered]@{
        path = [IO.Path]::GetRelativePath($repoRoot, $lock.lockPath).Replace('\', '/')
        sha256 = $lock.lockSha256
        candidateId = [string]$lock.engine.candidateId
        pin = [string]$lock.engine.pin
    }
    platform = $Platform
    configuration = $Configuration
    buildManifestSha256 = [string]$result.buildManifestSha256
    offline = $isolationSummary
    runSmoke = [bool]$RunSmoke
    native = $result.native
    outputs = @($result.outputs)
    runtimeIdentity = @($result.runtimeIdentity)
    smoke = @($result.smoke)
    status = 'pass'
}
Write-RSTerminalJsonFile -Path $manifestPath -Value $manifest

if ($PassThruManifest) {
    $manifestPath
}
else {
    Write-Host "Terminal engine verification passed: $Platform/$Configuration"
    Write-Host "Manifest: $manifestPath"
}
