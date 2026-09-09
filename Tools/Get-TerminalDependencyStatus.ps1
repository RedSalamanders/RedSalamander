<#
.SYNOPSIS
Checks dependency and cache status for an explicit Terminal-engine lock.

.DESCRIPTION
Compatibility entrypoint for generic Terminal qualification and upgrade work. It validates the lock's dependency graph, optionally checks canonical online endpoints and cached inputs, and can write a machine-readable status report. It is not the current Ghostty product lifecycle command.

.PARAMETER Mode
Uses Locked for declarative/cache validation or Online to probe declared canonical endpoints.

.PARAMETER EngineLock
Specifies the explicit Terminal-engine lock. LockFile is accepted as an alias.

.PARAMETER PackageToolchainLock
Optionally specifies the package/toolchain lock to include in the status record.

.PARAMETER RequireOnline
Fails when any canonical dependency endpoint is unreachable.

.PARAMETER RequireCache
Fails when a required locked input is not present in the cache.

.PARAMETER PassThruReport
Writes a JSON report and returns its path.

.PARAMETER ReportRoot
Selects the safe repository build directory for status reports.

.OUTPUTS
With -PassThruReport, a string containing the generated JSON path. Otherwise no pipeline output.

.NOTES
Prerequisites: a valid explicit Terminal-engine lock; Online mode also requires network access. Side effects: Online mode performs network probes; PassThruReport writes below .build. Exit is nonzero for invalid locks, unsafe output roots, missing required cache content, or required endpoint failures. Primary consumer: External/TerminalEngine/TerminalEngine.proj and reviewed upgrade diagnostics.

.EXAMPLE
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock .\.build\candidate-lock.json -RequireCache
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Locked', 'Online')]
    [string]$Mode,

    [Alias('LockFile')]
    [string]$EngineLock = 'External\terminal-engine.lock.json',

    [string]$PackageToolchainLock = '',

    [switch]$RequireOnline,

    [switch]$RequireCache,

    [switch]$PassThruReport,

    [string]$ReportRoot = '.build\TerminalDependencyStatus'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$lifecycleModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEngineLifecycle.psm1'
Import-Module $lifecycleModule -Force -ErrorAction Stop

if ($RequireOnline -and $Mode -ne 'Online') {
    throw '-RequireOnline is valid only in Online mode.'
}

$lock = Read-RSTerminalEngineLock `
    -LockFile $EngineLock `
    -ValidateInputs:$RequireCache
$cacheRecords = [Collections.Generic.List[object]]::new()
foreach ($inputRecord in @($lock.inputs)) {
    $sourceType = [string]$inputRecord.source.type
    $present = $true
    if ($sourceType -ceq 'https') {
        $cachePath = Get-RSTerminalInputCachePath `
            -RepoRoot $repoRoot `
            -Sha256 ([string]$inputRecord.sha256)
        $present = Test-Path -LiteralPath $cachePath -PathType Leaf
    }
    elseif ($sourceType -ceq 'repository-file') {
        $repositoryPath = Resolve-RSTerminalPathInRoot `
            -Root $repoRoot `
            -Path ([string]$inputRecord.source.path)
        $present = Test-Path -LiteralPath $repositoryPath -PathType Leaf
    }
    $cacheRecords.Add([ordered]@{
            id = [string]$inputRecord.id
            sourceType = $sourceType
            sha256 = [string]$inputRecord.sha256
            status = if (-not $present) {
                'missing'
            }
            elseif ($RequireCache) {
                'verified'
            }
            else {
                'present-unverified'
            }
        })
}
$cacheStatus = if (@($cacheRecords | Where-Object { $_.status -eq 'missing' }).Count -eq 0) {
    if ($RequireCache) { 'verified' } else { 'present-unverified' }
}
else {
    'incomplete'
}
$licenseClosure = @(Get-RSTerminalLicenseClosure -Lock $lock)
$noticeClosure = @(Get-RSTerminalNoticeClosure -Lock $lock)
$inputIds = @{}
foreach ($inputRecord in @($lock.inputs)) {
    $inputIds[[string]$inputRecord.id] = $true
}

$dependencyIds = @{}
foreach ($dependency in @($lock.dependencies)) {
    $dependencyId = [string](Get-RSObjectProperty `
            -Object $dependency `
            -Name id `
            -Required)
    if ([string]::IsNullOrWhiteSpace($dependencyId) -or
        $dependencyIds.ContainsKey($dependencyId)) {
        throw "Dependency id is empty or duplicated: '$dependencyId'."
    }
    $dependencyIds[$dependencyId] = $true
    [void](Get-RSObjectProperty -Object $dependency -Name pin -Required)
    foreach ($inputId in @(Get-RSObjectProperty `
            -Object $dependency `
            -Name inputIds `
            -Required)) {
        if (-not $inputIds.ContainsKey([string]$inputId)) {
            throw "Dependency '$dependencyId' refers to unknown input '$inputId'."
        }
    }
    $license = Get-RSObjectProperty -Object $dependency -Name license -Required
    if ([string]::IsNullOrWhiteSpace(
            [string](Get-RSObjectProperty -Object $license -Name spdx -Required))) {
        throw "Dependency '$dependencyId' has no SPDX identity."
    }
    foreach ($field in @('licenseInputIds', 'noticeInputIds')) {
        foreach ($inputId in @(Get-RSObjectProperty `
                -Object $license `
                -Name $field `
                -Default @())) {
            if (-not $inputIds.ContainsKey([string]$inputId)) {
                throw "Dependency '$dependencyId' $field refers to unknown input '$inputId'."
            }
        }
    }
}

$closureKeys = @{}
foreach ($entry in @($lock.runtimeClosure)) {
    $platform = [string](Get-RSObjectProperty -Object $entry -Name platform -Required)
    $configuration = [string](Get-RSObjectProperty `
            -Object $entry `
            -Name configuration `
            -Required)
    $module = [string](Get-RSObjectProperty -Object $entry -Name module -Required)
    $system = [bool](Get-RSObjectProperty -Object $entry -Name system -Required)
    $key = "$platform`0$configuration`0$module"
    if ($closureKeys.ContainsKey($key)) {
        throw "Duplicate runtime closure entry '$platform/$configuration/$module'."
    }
    $closureKeys[$key] = $true
    $outputId = Get-RSObjectProperty -Object $entry -Name outputId -Default $null
    $inputId = Get-RSObjectProperty -Object $entry -Name inputId -Default $null
    if (-not $system -and $null -eq $outputId -and $null -eq $inputId) {
        throw "Non-system runtime '$module' must bind an outputId or inputId."
    }
    if ($null -ne $outputId -and -not $lock.outputById.ContainsKey([string]$outputId)) {
        throw "Runtime closure entry '$module' refers to unknown output '$outputId'."
    }
    if ($null -ne $inputId -and -not $inputIds.ContainsKey([string]$inputId)) {
        throw "Runtime closure entry '$module' refers to unknown input '$inputId'."
    }
}

$packageLockRecord = $null
if (-not [string]::IsNullOrWhiteSpace($PackageToolchainLock)) {
    $packagePath = if ([IO.Path]::IsPathFullyQualified($PackageToolchainLock)) {
        [IO.Path]::GetFullPath($PackageToolchainLock)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $repoRoot $PackageToolchainLock))
    }
    if (-not (Test-RSPathIsWithin -Path $packagePath -Root $repoRoot) -or
        -not (Test-Path -LiteralPath $packagePath -PathType Leaf)) {
        throw 'PackageToolchainLock must be an existing file inside the repository.'
    }
    $packageLock = Get-Content -LiteralPath $packagePath -Raw |
        ConvertFrom-Json -Depth 100 -ErrorAction Stop
    foreach ($file in @(Get-RSObjectProperty `
            -Object $packageLock `
            -Name inputs `
            -Default @())) {
        $path = Resolve-RSTerminalPathInRoot `
            -Root $repoRoot `
            -Path ([string](Get-RSObjectProperty -Object $file -Name path -Required)) `
            -MustExist
        $expectedSha = [string](Get-RSObjectProperty -Object $file -Name sha256 -Required)
        if ((Get-RSSha256 -Path $path) -cne $expectedSha) {
            throw "Package toolchain input '$($file.id)' differs from its lock."
        }
    }
    $packageLockRecord = [ordered]@{
        sha256 = Get-RSSha256 -Path $packagePath
        validated = $true
    }
}

function Get-RSDiscoveryPayload {
    param(
        [Parameter(Mandatory)]
        [object]$Discovery
    )

    $kind = [string](Get-RSObjectProperty -Object $Discovery -Name kind -Required)
    $uri = [string](Get-RSObjectProperty -Object $Discovery -Name uri -Default '')
    if ($kind -eq 'manual') {
        return [pscustomobject]@{
            status = 'required-manual'
            value = $null
            endpointKind = $kind
        }
    }
    if ($kind -in @('file-text', 'file-json')) {
        $filePath = if ([Uri]::IsWellFormedUriString($uri, [UriKind]::Absolute)) {
            $parsed = [Uri]$uri
            if (-not $parsed.IsFile) {
                throw "Discovery '$kind' requires a file URI or repository-relative path."
            }
            [IO.Path]::GetFullPath($parsed.LocalPath)
        }
        else {
            [IO.Path]::GetFullPath((Join-Path $repoRoot $uri))
        }
        if (-not (Test-RSPathIsWithin -Path $filePath -Root $repoRoot) -or
            -not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
            throw 'File discovery endpoint must be an existing file inside the repository.'
        }
        $payload = [IO.File]::ReadAllText($filePath, [Text.Encoding]::UTF8)
    }
    elseif ($kind -in @('https-text', 'https-json')) {
        $parsed = [Uri]$uri
        if ($parsed.Scheme -cne 'https' -or
            -not [string]::IsNullOrEmpty($parsed.UserInfo)) {
            throw 'Online discovery endpoints must be credential-free HTTPS URLs.'
        }
        $handler = [Net.Http.HttpClientHandler]::new()
        $handler.AllowAutoRedirect = $false
        $handler.UseDefaultCredentials = $false
        $handler.Credentials = $null
        $handler.UseProxy = $false
        $handler.Proxy = $null
        $client = [Net.Http.HttpClient]::new($handler)
        $client.Timeout = [TimeSpan]::FromSeconds(15)
        try {
            $response = $client.GetAsync($parsed).GetAwaiter().GetResult()
            if (-not $response.IsSuccessStatusCode) {
                throw "HTTP status $([int]$response.StatusCode)."
            }
            $payload = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        }
        finally {
            $client.Dispose()
            $handler.Dispose()
        }
    }
    else {
        throw "Unsupported update-discovery kind '$kind'."
    }

    if ($kind.EndsWith('-json')) {
        $json = $payload | ConvertFrom-Json -Depth 50 -ErrorAction Stop
        $propertyPath = [string](Get-RSObjectProperty `
                -Object $Discovery `
                -Name jsonProperty `
                -Required)
        $value = $json
        foreach ($part in $propertyPath.Split('.')) {
            $property = $value.PSObject.Properties[$part]
            if ($null -eq $property) {
                throw "JSON discovery property '$propertyPath' was not found."
            }
            $value = $property.Value
        }
        $payload = [string]$value
    }
    return [pscustomobject]@{
        status = $null
        value = $payload.Trim()
        endpointKind = $kind
    }
}

$statusRecords = [Collections.Generic.List[object]]::new()
if ($Mode -eq 'Online') {
    foreach ($dependency in @($lock.dependencies)) {
        $discoveries = @(Get-RSObjectProperty `
                -Object $dependency `
                -Name updateDiscovery `
                -Default @())
        if ($discoveries.Count -eq 0) {
            $statusRecords.Add([ordered]@{
                    dependencyId = [string]$dependency.id
                    discoveryId = 'missing'
                    status = 'required-manual'
                    endpointKind = 'manual'
                    observedIdentitySha256 = $null
                })
            continue
        }
        foreach ($discovery in $discoveries) {
            $discoveryId = [string](Get-RSObjectProperty `
                    -Object $discovery `
                    -Name id `
                    -Required)
            try {
                $observation = Get-RSDiscoveryPayload -Discovery $discovery
                if ($observation.status -eq 'required-manual') {
                    $status = 'required-manual'
                    $observedSha = $null
                }
                else {
                    $current = [string](Get-RSObjectProperty `
                            -Object $discovery `
                            -Name currentValue `
                            -Required)
                    $status = if ([string]::Equals(
                            [string]$observation.value,
                            $current,
                            [StringComparison]::Ordinal)) {
                        'current'
                    }
                    else {
                        'update-observed'
                    }
                    $observedSha = Get-RSCanonicalObjectSha256 `
                        -InputObject ([string]$observation.value)
                    if ([bool](Get-RSObjectProperty `
                            -Object $discovery `
                            -Name securityAdvisory `
                            -Default $false) -and
                        $status -eq 'update-observed') {
                        $status = 'security-review'
                    }
                }
            }
            catch {
                $status = 'unreachable'
                $observedSha = $null
            }
            $statusRecords.Add([ordered]@{
                    dependencyId = [string]$dependency.id
                    discoveryId = $discoveryId
                    status = $status
                    endpointKind = [string](Get-RSObjectProperty `
                            -Object $discovery `
                            -Name kind `
                            -Required)
                    observedIdentitySha256 = $observedSha
                })
        }
    }
}

$report = [ordered]@{
    formatId = 'red-salamander-terminal-dependency-status'
    schemaVersion = 3
    createdUtc = [DateTime]::UtcNow.ToString(
        'o',
        [Globalization.CultureInfo]::InvariantCulture)
    mode = $Mode
    engineLockSha256 = $lock.lockSha256
    candidateId = [string]$lock.engine.candidateId
    pin = [string]$lock.engine.pin
    bootstrapInputSetSha256 = Get-RSCanonicalObjectSha256 -InputObject @(
        $lock.inputs | ForEach-Object {
            [ordered]@{
                id = [string]$_.id
                sourceType = [string]$_.source.type
                sha256 = [string]$_.sha256
            }
        })
    cacheRequired = [bool]$RequireCache
    cacheStatus = $cacheStatus
    cacheInputs = $cacheRecords.ToArray()
    dependencySetSha256 = Get-RSCanonicalObjectSha256 -InputObject @($lock.dependencies)
    sourceAdaptationSha256 = Get-RSCanonicalObjectSha256 `
        -InputObject $lock.sourceAdaptation
    derivedInputSetSha256 = Get-RSCanonicalObjectSha256 `
        -InputObject @($lock.derivedInputs)
    installedToolchainSetSha256 = Get-RSCanonicalObjectSha256 `
        -InputObject @($lock.installedToolchains)
    licenseClosureSha256 = Get-RSCanonicalObjectSha256 `
        -InputObject $licenseClosure
    noticeClosureSha256 = Get-RSCanonicalObjectSha256 `
        -InputObject $noticeClosure
    abiSha256 = Get-RSCanonicalObjectSha256 -InputObject $lock.abi
    runtimeClosureSha256 = Get-RSCanonicalObjectSha256 -InputObject @($lock.runtimeClosure)
    runtimeLoadGraphSha256 = Get-RSCanonicalObjectSha256 -InputObject @(
        $lock.runtimeClosure |
            Sort-Object -Stable -CaseSensitive -Property `
                platform,
                configuration,
                module |
            ForEach-Object {
                [ordered]@{
                    platform = [string]$_.platform
                    configuration = [string]$_.configuration
                    module = [string]$_.module
                    system = [bool]$_.system
                    loadSource = [string]$_.loadSource
                    loadedBy = [string]$_.loadedBy
                }
            })
    runtimeClosureEvidence = 'declarative-lock-graph'
    packageToolchain = $packageLockRecord
    observations = $statusRecords.ToArray()
    lockedStatus = 'pass'
}

$reportPath = $null
if ($PassThruReport) {
    $resolvedReportRoot = Assert-RSTerminalSafeBuildRoot `
        -RepoRoot $repoRoot `
        -BuildRoot $ReportRoot
    [void](New-Item -ItemType Directory -Path $resolvedReportRoot -Force)
    $reportPath = Join-Path $resolvedReportRoot (
        "status-$([DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss'))-$([Guid]::NewGuid().ToString('N')).json")
    Write-RSTerminalJsonFile -Path $reportPath -Value $report
}

if ($RequireOnline -and
    @($statusRecords | Where-Object { $_.status -eq 'unreachable' }).Count -gt 0) {
    throw 'One or more canonical online dependency endpoints were unreachable.'
}

if ($PassThruReport) {
    $reportPath
}
else {
    Write-Host "Terminal dependency lock is structurally valid: $($lock.engine.candidateId) $($lock.engine.pin)"
    Write-Host "  input cache: $cacheStatus (required=$([bool]$RequireCache))"
    foreach ($record in $statusRecords) {
        Write-Host "  $($record.dependencyId)/$($record.discoveryId): $($record.status)"
    }
}
