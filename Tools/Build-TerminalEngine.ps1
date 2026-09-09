<#
.SYNOPSIS
Builds one lane from an explicit Terminal-engine evaluation lock.

.DESCRIPTION
Compatibility entrypoint for the generic lock-based qualification lifecycle. It validates a supplied lock, builds its requested architecture/configuration lane, and writes a lock-bound build manifest. The current product runtime is built by Build-GhosttyTerminalRuntime.ps1; this command is retained for reproducible upgrade and historical qualification work.

.PARAMETER Platform
Selects the x64 or ARM64 lane.

.PARAMETER Configuration
Selects Debug, Release, or ASan Debug.

.PARAMETER LockFile
Specifies the explicit evaluation or candidate lock. The historical default path is not present in a normal product checkout, so callers must normally supply this parameter.

.PARAMETER OutputRoot
Selects the safe repository build root for lane outputs and manifests.

.PARAMETER Clean
Removes the selected lane's prior output before building.

.PARAMETER Offline
Runs the build under verified host-wide network isolation.

.PARAMETER ExclusiveDisposableRunner
Attests that the offline operation owns a disposable runner; required with -Offline.

.PARAMETER PassThruManifest
Returns the generated build-manifest path.

.OUTPUTS
With -PassThruManifest, a string containing the manifest path. Otherwise no pipeline output. Build outputs and terminal-engine-build-manifest.json are written below OutputRoot.

.NOTES
Prerequisites: a valid explicit Terminal-engine lock and all toolchains/inputs it declares. Side effects: builds external code and writes below OutputRoot; offline mode creates and removes a worker result. Exit is nonzero on invalid locks, unsafe paths, isolation failure, build failure, or output mismatch. Primary consumers: External/TerminalEngine/TerminalEngine.proj and the reviewed upgrade workflow. This is not the canonical product Ghostty builder.

.EXAMPLE
.\Tools\Build-TerminalEngine.ps1 -LockFile .\.build\candidate-lock.json -Platform x64 -Configuration Release -Clean -PassThruManifest
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

    [switch]$Clean,

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

$lock = Read-RSTerminalEngineLock -LockFile $LockFile -ValidateInputs
$resolvedOutputRoot = Assert-RSTerminalSafeBuildRoot `
    -RepoRoot $repoRoot `
    -BuildRoot $OutputRoot
[void](Get-RSTerminalLane `
        -Lock $lock `
        -Platform $Platform `
        -Configuration $Configuration)

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
    $result = Invoke-RSTerminalBuildLane `
        -Lock $lock `
        -Platform $Platform `
        -Configuration $Configuration `
        -OutputRoot $resolvedOutputRoot `
        -Clean:$Clean
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
        "build-$([Guid]::NewGuid().ToString('N')).json")
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
    if ($Clean) {
        $arguments.Add('-Clean')
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
        throw 'Runner-enforced offline terminal-engine build failed.'
    }
    if (-not (Test-Path -LiteralPath $workerResult -PathType Leaf)) {
        throw 'Offline build worker did not emit its exact result.'
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
    $result = Invoke-RSTerminalBuildLane `
        -Lock $lock `
        -Platform $Platform `
        -Configuration $Configuration `
        -OutputRoot $resolvedOutputRoot `
        -Clean:$Clean
}

$laneOutputRoot = Join-Path (Join-Path $resolvedOutputRoot $Platform) $Configuration
$manifestPath = Join-Path $laneOutputRoot 'terminal-engine-build-manifest.json'
$manifest = [ordered]@{
    formatId = 'red-salamander-terminal-engine-build'
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
    clean = [bool]$Clean
    offline = $isolationSummary
    inputs = @($lock.verifiedInputs | ForEach-Object {
            [ordered]@{
                id = $_.id
                kind = $_.kind
                sha256 = $_.sha256
                sizeBytes = $_.sizeBytes
            }
        })
    command = [ordered]@{
        reused = [bool](Get-RSObjectProperty `
                -Object $result.command `
                -Name reused `
                -Default $false)
        stdoutSha256 = Get-RSObjectProperty `
            -Object $result.command `
            -Name stdoutSha256 `
            -Default $null
        stderrSha256 = Get-RSObjectProperty `
            -Object $result.command `
            -Name stderrSha256 `
            -Default $null
    }
    outputs = @($result.outputs)
}
Write-RSTerminalJsonFile -Path $manifestPath -Value $manifest

if ($PassThruManifest) {
    $manifestPath
}
else {
    Write-Host "Terminal engine built and verified: $Platform/$Configuration"
    Write-Host "Manifest: $manifestPath"
}
