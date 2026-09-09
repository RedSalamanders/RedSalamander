Set-StrictMode -Version Latest

$textIdentityModule = Join-Path $PSScriptRoot 'TerminalEngineTextIdentity.psm1'
$exportIdentityModule = Join-Path $PSScriptRoot 'TerminalRuntimeExportIdentity.psm1'
$canonicalJsonModule = Join-Path (Split-Path -Parent $PSScriptRoot) 'Modules\Testing\ValidationFingerprint.psm1'
Import-Module $textIdentityModule -Scope Local -ErrorAction Stop
Import-Module $exportIdentityModule -Scope Local -ErrorAction Stop
Import-Module $canonicalJsonModule -Scope Local -ErrorAction Stop

function Assert-RSGhosttyRuntimeLockHash {
    param(
        [Parameter(Mandatory = $true)][string]$Field,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value,
        [Parameter(Mandatory = $true)][ValidateSet(40, 64)][int]$Length
    )

    $pattern = if ($Length -eq 40) { '^[0-9a-f]{40}$' } else { '^[0-9a-f]{64}$' }
    if ($Value -cnotmatch $pattern) {
        throw "Ghostty runtime lock field '$Field' must be a lowercase $Length-character hexadecimal digest."
    }
}

function Assert-RSGhosttyRuntimeLockSortedUniqueStrings {
    param(
        [Parameter(Mandatory = $true)][string]$Field,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Values
    )

    $strings = [string[]]@($Values | ForEach-Object { [string]$_ })
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($value in $strings) {
        if (-not $seen.Add($value)) {
            throw "Ghostty runtime lock field '$Field' contains duplicate value '$value'."
        }
    }

    $sorted = [string[]]$strings.Clone()
    [Array]::Sort($sorted, [StringComparer]::Ordinal)
    if (($strings -join "`0") -cne ($sorted -join "`0")) {
        throw "Ghostty runtime lock field '$Field' must be sorted with ordinal case-sensitive ordering."
    }
}

function Assert-RSGhosttyRuntimeLockUniqueStrings {
    param(
        [Parameter(Mandatory = $true)][string]$Field,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Values,
        [switch]$CaseInsensitive
    )

    $comparer = if ($CaseInsensitive) {
        [StringComparer]::OrdinalIgnoreCase
    }
    else {
        [StringComparer]::Ordinal
    }
    $seen = [Collections.Generic.HashSet[string]]::new($comparer)
    foreach ($value in $Values) {
        $stringValue = [string]$value
        if (-not $seen.Add($stringValue)) {
            throw "Ghostty runtime lock field '$Field' contains duplicate value '$stringValue'."
        }
    }
}

function Resolve-RSGhosttyRuntimeLockPath {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)][string]$Field
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $candidate = [IO.Path]::GetFullPath((Join-Path $root $RelativePath))
    $rootPrefix = $root + '\'
    if (-not $candidate.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Ghostty runtime lock field '$Field' escapes the repository root: $RelativePath"
    }
    return $candidate
}

function Assert-RSGhosttyRuntimeLockShape {
    param(
        [Parameter(Mandatory = $true)][Collections.IDictionary]$LockTable
    )

    $required = @(
        'formatId',
        'schemaVersion',
        'upstream',
        'toolchain',
        'dependencies',
        'overlay',
        'build',
        'abi',
        'notices',
        'runtimeIdentity')
    foreach ($name in $required) {
        if (-not $LockTable.Contains($name)) {
            throw "Ghostty runtime lock is missing required property '$name'."
        }
    }

    $allowed = [Collections.Generic.HashSet[string]]::new([string[]]$required, [StringComparer]::Ordinal)
    foreach ($name in $LockTable.Keys) {
        if (-not $allowed.Contains([string]$name)) {
            throw "Ghostty runtime lock contains unsupported additional property '$name'."
        }
    }

    if ([int]$LockTable.schemaVersion -ne 1) {
        throw "Ghostty runtime lock field 'schemaVersion' must be exactly 1."
    }
}

function Assert-RSGhosttyRuntimeLockSemantics {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Lock,
        [Parameter(Mandatory = $true)][string]$RepoRoot
    )

    Assert-RSGhosttyRuntimeLockHash -Field 'upstream.commit' -Value ([string]$Lock.upstream.commit) -Length 40
    Assert-RSGhosttyRuntimeLockHash -Field 'upstream.tree' -Value ([string]$Lock.upstream.tree) -Length 40
    Assert-RSGhosttyRuntimeLockHash -Field 'upstream.archive.sha256' -Value ([string]$Lock.upstream.archive.sha256) -Length 64
    foreach ($archive in @($Lock.toolchain.archives)) {
        Assert-RSGhosttyRuntimeLockHash -Field "toolchain.archives[$($archive.hostPlatform)].sha256" -Value ([string]$archive.sha256) -Length 64
        Assert-RSGhosttyRuntimeLockHash -Field "toolchain.archives[$($archive.hostPlatform)].executableSha256" -Value ([string]$archive.executableSha256) -Length 64
    }
    foreach ($dependency in @($Lock.dependencies)) {
        Assert-RSGhosttyRuntimeLockHash -Field "dependencies[$($dependency.id)].sourceArchive.sha256" -Value ([string]$dependency.sourceArchive.sha256) -Length 64
        Assert-RSGhosttyRuntimeLockHash -Field "dependencies[$($dependency.id)].contentArchive.sha256" -Value ([string]$dependency.contentArchive.sha256) -Length 64
        Assert-RSGhosttyRuntimeLockHash -Field "dependencies[$($dependency.id)].licenseSha256" -Value ([string]$dependency.licenseSha256) -Length 64
    }
    Assert-RSGhosttyRuntimeLockHash -Field 'overlay.normalizedTextSha256' -Value ([string]$Lock.overlay.normalizedTextSha256) -Length 64
    Assert-RSGhosttyRuntimeLockHash -Field 'abi.exportSetSha256' -Value ([string]$Lock.abi.exportSetSha256) -Length 64
    foreach ($notice in @($Lock.notices)) {
        Assert-RSGhosttyRuntimeLockHash -Field "notices[$($notice.id)].normalizedTextSha256" -Value ([string]$notice.normalizedTextSha256) -Length 64
    }

    Assert-RSGhosttyRuntimeLockSortedUniqueStrings -Field 'toolchain.archives.hostPlatform' -Values @($Lock.toolchain.archives.hostPlatform)
    Assert-RSGhosttyRuntimeLockSortedUniqueStrings -Field 'dependencies.id' -Values @($Lock.dependencies.id)
    Assert-RSGhosttyRuntimeLockSortedUniqueStrings -Field 'overlay.concernIds' -Values @($Lock.overlay.concernIds)
    Assert-RSGhosttyRuntimeLockSortedUniqueStrings -Field 'build.environmentAllowlist' -Values @($Lock.build.environmentAllowlist)
    Assert-RSGhosttyRuntimeLockSortedUniqueStrings -Field 'build.platforms.platform' -Values @($Lock.build.platforms.platform)
    Assert-RSGhosttyRuntimeLockSortedUniqueStrings -Field 'abi.importModules' -Values @($Lock.abi.importModules)
    Assert-RSGhosttyRuntimeLockSortedUniqueStrings -Field 'abi.exports' -Values @($Lock.abi.exports)
    Assert-RSGhosttyRuntimeLockSortedUniqueStrings -Field 'abi.loaderRequiredExports' -Values @($Lock.abi.loaderRequiredExports)
    Assert-RSGhosttyRuntimeLockSortedUniqueStrings -Field 'notices.id' -Values @($Lock.notices.id)
    Assert-RSGhosttyRuntimeLockUniqueStrings `
        -Field 'dependency/toolchain download fileName' `
        -CaseInsensitive `
        -Values @(
            @($Lock.toolchain.archives | ForEach-Object { $_.fileName })
            @($Lock.dependencies | ForEach-Object { $_.sourceArchive.fileName; $_.contentArchive.fileName })
        )
    Assert-RSGhosttyRuntimeLockUniqueStrings `
        -Field 'dependencies.packageName' `
        -CaseInsensitive `
        -Values @($Lock.dependencies.packageName)
    Assert-RSGhosttyRuntimeLockUniqueStrings `
        -Field 'notices.outputRelativePath' `
        -CaseInsensitive `
        -Values @($Lock.notices.outputRelativePath)

    $exportIdentity = Get-RSTerminalExportSetIdentity -ExportNames @($Lock.abi.exports)
    if ([int]$Lock.abi.exportCount -ne [int]$exportIdentity.count) {
        throw "Ghostty runtime lock field 'abi.exportCount' is $($Lock.abi.exportCount), but abi.exports contains $($exportIdentity.count) names."
    }
    if ([string]$Lock.abi.exportSetSha256 -cne [string]$exportIdentity.sha256) {
        throw "Ghostty runtime lock field 'abi.exportSetSha256' does not match the derived sorted export identity $($exportIdentity.sha256)."
    }
    foreach ($requiredExport in @($Lock.abi.loaderRequiredExports)) {
        if (@($Lock.abi.exports) -cnotcontains [string]$requiredExport) {
            throw "Ghostty runtime lock field 'abi.loaderRequiredExports' names '$requiredExport', which is absent from abi.exports."
        }
    }

    $platforms = @($Lock.build.platforms)
    foreach ($expected in @(
            [pscustomobject]@{ Platform = 'ARM64'; Target = 'aarch64-windows-msvc'; Machine = 'ARM64' },
            [pscustomobject]@{ Platform = 'x64'; Target = 'x86_64-windows-msvc'; Machine = 'AMD64' })) {
        $matches = @($platforms | Where-Object { [string]$_.platform -ceq $expected.Platform })
        if ($matches.Count -ne 1) {
            throw "Ghostty runtime lock field 'build.platforms' must contain exactly one '$($expected.Platform)' record."
        }
        $record = $matches[0]
        if ([string]$record.target -cne $expected.Target) {
            throw "Ghostty runtime lock field 'build.platforms[$($expected.Platform)].target' must be '$($expected.Target)'."
        }
        if ([string]$record.machine -cne $expected.Machine) {
            throw "Ghostty runtime lock field 'build.platforms[$($expected.Platform)].machine' must be '$($expected.Machine)'."
        }
        if ([long]$record.minimumDllSizeBytes -gt [long]$record.maximumDllSizeBytes) {
            throw "Ghostty runtime lock fields 'build.platforms[$($expected.Platform)].minimumDllSizeBytes' and 'maximumDllSizeBytes' form an invalid envelope."
        }
    }

    $overlayPath = Resolve-RSGhosttyRuntimeLockPath -RepoRoot $RepoRoot -RelativePath ([string]$Lock.overlay.path) -Field 'overlay.path'
    if (-not (Test-Path -LiteralPath $overlayPath -PathType Leaf)) {
        throw "Ghostty runtime lock field 'overlay.path' names a missing file: $($Lock.overlay.path)"
    }
    $overlayIdentity = Get-RSTerminalLfTextIdentity -LiteralPath $overlayPath
    if ([string]$overlayIdentity.Sha256 -cne [string]$Lock.overlay.normalizedTextSha256) {
        throw "Ghostty runtime lock field 'overlay.normalizedTextSha256' does not match '$($Lock.overlay.path)': $($overlayIdentity.Sha256)"
    }

    $dependencyIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($dependency in @($Lock.dependencies)) {
        [void]$dependencyIds.Add([string]$dependency.id)
    }
    foreach ($requiredNoticeId in @('ghostty', 'unicode', 'uucode', 'uucode-bjoern-hoehrmann', 'zig')) {
        $matches = @($Lock.notices | Where-Object { [string]$_.id -ceq $requiredNoticeId })
        if ($matches.Count -ne 1) {
            throw "Ghostty runtime lock field 'notices' must contain exactly one '$requiredNoticeId' record."
        }
    }
    foreach ($notice in @($Lock.notices)) {
        $outputPath = Resolve-RSGhosttyRuntimeLockPath -RepoRoot $RepoRoot -RelativePath ([string]$notice.outputRelativePath) -Field "notices[$($notice.id)].outputRelativePath"
        if ([string]$notice.source.kind -ceq 'dependency') {
            if (-not $dependencyIds.Contains([string]$notice.source.dependencyId)) {
                throw "Ghostty runtime lock field 'notices[$($notice.id)].source.dependencyId' names unknown dependency '$($notice.source.dependencyId)'."
            }
            $dependency = @($Lock.dependencies | Where-Object { [string]$_.id -ceq [string]$notice.source.dependencyId })[0]
            if ([string]$notice.source.path -cne [string]$dependency.licenseArchiveRelativePath -or
                [string]$notice.normalizedTextSha256 -cne [string]$dependency.licenseSha256) {
                throw "Ghostty runtime lock notice '$($notice.id)' does not match dependency '$($dependency.id)' license identity."
            }
        }
        elseif ([string]$notice.source.kind -ceq 'repository') {
            $sourcePath = Resolve-RSGhosttyRuntimeLockPath -RepoRoot $RepoRoot -RelativePath ([string]$notice.source.path) -Field "notices[$($notice.id)].source.path"
            if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
                throw "Ghostty runtime lock field 'notices[$($notice.id)].source.path' names a missing license input: $($notice.source.path)"
            }
            $noticeIdentity = Get-RSTerminalLfTextIdentity -LiteralPath $sourcePath
            if ([string]$noticeIdentity.Sha256 -cne [string]$notice.normalizedTextSha256) {
                throw "Ghostty runtime lock field 'notices[$($notice.id)].normalizedTextSha256' does not match '$($notice.source.path)': $($noticeIdentity.Sha256)"
            }
        }
        [void]$outputPath
    }
}

function Read-RSGhosttyRuntimeLock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter(Mandatory = $true)][string]$SchemaPath,
        [Parameter(Mandatory = $true)][string]$RepoRoot
    )

    if (-not (Test-Path -LiteralPath $LiteralPath -PathType Leaf)) {
        throw "Ghostty runtime lock file is missing: $LiteralPath"
    }
    if (-not (Test-Path -LiteralPath $SchemaPath -PathType Leaf)) {
        throw "Ghostty runtime lock schema is missing: $SchemaPath"
    }

    $raw = Get-Content -LiteralPath $LiteralPath -Raw -ErrorAction Stop
    $table = $raw | ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String -ErrorAction Stop
    Assert-RSGhosttyRuntimeLockShape -LockTable $table

    $schemaErrors = @()
    $schemaValid = Test-Json -Json $raw -SchemaFile $SchemaPath -ErrorVariable +schemaErrors -ErrorAction SilentlyContinue
    if (-not $schemaValid) {
        $details = @($schemaErrors | ForEach-Object { $_.Exception.Message }) -join '; '
        $details = [regex]::Replace(
            $details,
            "'/(?<pointer>[^']+)'",
            { param($match) "'$($match.Groups['pointer'].Value.Replace('/', '.'))'" })
        throw "Ghostty runtime lock schema validation failed: $details"
    }

    $lock = $raw | ConvertFrom-Json -Depth 100 -NoEnumerate -DateKind String -ErrorAction Stop
    Assert-RSGhosttyRuntimeLockSemantics -Lock $lock -RepoRoot $RepoRoot
    return $lock
}

function Get-RSGhosttyRuntimeLockDigest {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][pscustomobject]$Lock)

    return Get-RSCanonicalJsonDigest -InputObject $Lock
}

function Get-RSGhosttyRuntimeLockPlatform {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Lock,
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform
    )

    $matches = @($Lock.build.platforms | Where-Object { [string]$_.platform -ceq $Platform })
    if ($matches.Count -ne 1) {
        throw "Ghostty runtime lock field 'build.platforms' does not contain exactly one '$Platform' record."
    }
    return $matches[0]
}

function Get-RSGhosttyRuntimeLockDependency {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Lock,
        [Parameter(Mandatory = $true)][string]$Id
    )

    $matches = @($Lock.dependencies | Where-Object { [string]$_.id -ceq $Id })
    if ($matches.Count -ne 1) {
        throw "Ghostty runtime lock field 'dependencies' does not contain exactly one '$Id' record."
    }
    return $matches[0]
}

function Test-RSGhosttyRuntimePathWithinRoot {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$AllowRoot
    )

    $fullRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $fullPath = [IO.Path]::GetFullPath($Path)
    if ($AllowRoot -and $fullPath.Equals($fullRoot, [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    return $fullPath.StartsWith($fullRoot + '\', [StringComparison]::OrdinalIgnoreCase)
}

function Resolve-RSGhosttyRuntimeBuildPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform,
        [string]$CandidateLockPath,
        [string]$CandidateOutputRoot
    )

    $resolvedRepoRoot = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $schemaPath = Join-Path $resolvedRepoRoot 'Specs\Terminal\GhosttyRuntimeLock.schema.json'
    $canonicalLockPath = Join-Path $resolvedRepoRoot 'External\TerminalEngine\GhosttyRuntimeLock.v1.json'
    $productStateRoot = Join-Path $resolvedRepoRoot '.build\TerminalEngine'
    $productOutputRoot = Join-Path (Join-Path $productStateRoot 'Product') $Platform
    $upgradeRoot = Join-Path $resolvedRepoRoot '.build\TerminalEngineUpgrade'
    $hasCandidateLock = -not [string]::IsNullOrWhiteSpace($CandidateLockPath)
    $hasCandidateOutput = -not [string]::IsNullOrWhiteSpace($CandidateOutputRoot)
    if ($hasCandidateLock -ne $hasCandidateOutput) {
        throw 'Candidate mode requires both -CandidateLockPath and -CandidateOutputRoot; no canonical-lock or output fallback is allowed.'
    }

    if (-not $hasCandidateLock) {
        return [pscustomobject]@{
            Mode = 'Product'
            LockPath = $canonicalLockPath
            SchemaPath = $schemaPath
            StateRoot = $productStateRoot
            OutputRoot = $productOutputRoot
            WritePairingHeader = $true
        }
    }

    $resolvedLockPath = if ([IO.Path]::IsPathFullyQualified($CandidateLockPath)) {
        [IO.Path]::GetFullPath($CandidateLockPath)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $resolvedRepoRoot $CandidateLockPath))
    }
    $resolvedOutputRoot = if ([IO.Path]::IsPathFullyQualified($CandidateOutputRoot)) {
        [IO.Path]::GetFullPath($CandidateOutputRoot)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $resolvedRepoRoot $CandidateOutputRoot))
    }

    if (-not (Test-RSGhosttyRuntimePathWithinRoot -Root $upgradeRoot -Path $resolvedLockPath)) {
        throw "Candidate lock must be a contained file below '$upgradeRoot': $resolvedLockPath"
    }
    if (-not (Test-RSGhosttyRuntimePathWithinRoot -Root $upgradeRoot -Path $resolvedOutputRoot)) {
        throw "Candidate output must be a contained child below '$upgradeRoot': $resolvedOutputRoot"
    }
    if (-not (Test-Path -LiteralPath $resolvedLockPath -PathType Leaf)) {
        throw "Candidate lock file is missing; no canonical-lock fallback is allowed: $resolvedLockPath"
    }
    if (Test-RSGhosttyRuntimePathWithinRoot -Root $productStateRoot -Path $resolvedOutputRoot -AllowRoot) {
        throw "Candidate output cannot target Product state: $resolvedOutputRoot"
    }

    return [pscustomobject]@{
        Mode = 'Candidate'
        LockPath = $resolvedLockPath
        SchemaPath = $schemaPath
        StateRoot = $upgradeRoot
        OutputRoot = $resolvedOutputRoot
        WritePairingHeader = $false
    }
}

function Assert-RSGhosttyRuntimeOutputReplaceable {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$OutputRoot)

    $runtimePath = Join-Path ([IO.Path]::GetFullPath($OutputRoot)) 'bin\ghostty-vt.dll'
    if (-not (Test-Path -LiteralPath $runtimePath -PathType Leaf)) {
        return
    }

    $probe = $null
    try {
        $probe = [IO.File]::Open(
            $runtimePath,
            [IO.FileMode]::Open,
            [IO.FileAccess]::ReadWrite,
            [IO.FileShare]::None)
    }
    catch [UnauthorizedAccessException] {
        throw "Cannot replace the exact Ghostty runtime output '$runtimePath'. An independently owned or protected file blocks publication; close its owner and retry. No process was terminated."
    }
    catch [IO.IOException] {
        throw "Cannot replace the exact Ghostty runtime output '$runtimePath'. An independently owned or protected file blocks publication; close its owner and retry. No process was terminated."
    }
    finally {
        if ($null -ne $probe) {
            $probe.Dispose()
        }
    }
}

function Invoke-RSGhosttyRuntimeBuild {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform,
        [switch]$Force,
        [switch]$Offline,
        [switch]$PassThru,
        [string]$CandidateLockPath,
        [string]$CandidateOutputRoot
    )

    $repoRoot = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $buildPlan = Resolve-RSGhosttyRuntimeBuildPlan `
        -RepoRoot $repoRoot `
        -Platform $Platform `
        -CandidateLockPath $CandidateLockPath `
        -CandidateOutputRoot $CandidateOutputRoot
    $runtimeLock = Read-RSGhosttyRuntimeLock `
        -LiteralPath $buildPlan.LockPath `
        -SchemaPath $buildPlan.SchemaPath `
        -RepoRoot $repoRoot
    $runtimeLockDigest = Get-RSGhosttyRuntimeLockDigest -Lock $runtimeLock
    $runtimePlatform = Get-RSGhosttyRuntimeLockPlatform -Lock $runtimeLock -Platform $Platform
    $dependencyLocks = @($runtimeLock.dependencies)
    $uucodeLock = Get-RSGhosttyRuntimeLockDependency -Lock $runtimeLock -Id 'uucode'
    $hostToolchainArchives = @($runtimeLock.toolchain.archives | Where-Object { [string]$_.hostPlatform -ceq 'x64' })
    if ($hostToolchainArchives.Count -ne 1) {
        throw "Ghostty runtime lock field 'toolchain.archives' must contain exactly one x64 host toolchain."
    }
    $zigArchiveLock = $hostToolchainArchives[0]

    $ghosttyCommit = [string]$runtimeLock.upstream.commit
    $ghosttyTree = [string]$runtimeLock.upstream.tree
    $ghosttyVersion = [string]$runtimeLock.upstream.version
    $ghosttyRepository = [string]$runtimeLock.upstream.repository
    $zigVersion = [string]$runtimeLock.toolchain.zigVersion
    $zigArchiveName = [string]$zigArchiveLock.fileName
    $zigArchiveUri = [string]$zigArchiveLock.url
    $zigArchiveSha256 = [string]$zigArchiveLock.sha256
    $zigExecutableSha256 = [string]$zigArchiveLock.executableSha256
    $uucodePackageName = [string]$uucodeLock.packageName
    $uucodeSourceArchiveName = [string]$uucodeLock.sourceArchive.fileName
    $uucodeSourceArchiveUri = [string]$uucodeLock.sourceArchive.url
    $uucodeSourceArchiveSha256 = [string]$uucodeLock.sourceArchive.sha256
    $uucodeSourceInternalRoot = [string]$uucodeLock.sourceArchive.internalRoot
    $uucodePackageCacheFileName = [string]$uucodeLock.contentArchive.fileName
    $uucodePackageArchiveSha256 = [string]$uucodeLock.contentArchive.sha256
    $uucodePackageInternalRoot = [string]$uucodeLock.contentArchive.internalRoot
    $uucodeLicenseArchiveRelativePath = [string]$uucodeLock.licenseArchiveRelativePath
    $uucodeLicenseSha256 = [string]$uucodeLock.licenseSha256
    $buildArguments = [string[]]@($runtimeLock.build.arguments)
    $contentHashNormalization = [string]$runtimeLock.build.contentHashNormalization
    $canonicalBuildDrive = [string]$runtimeLock.build.canonicalDrive
    $target = [string]$runtimePlatform.target
    $expectedMachine = [string]$runtimePlatform.machine
    $minimumDllSizeBytes = [long]$runtimePlatform.minimumDllSizeBytes
    $maximumDllSizeBytes = [long]$runtimePlatform.maximumDllSizeBytes
    $pairingContract = [string]$runtimeLock.runtimeIdentity.pairingContract
    $expectedExports = [string[]]@($runtimeLock.abi.exports)
    $requiredExports = [string[]]@($runtimeLock.abi.loaderRequiredExports)
    $expectedExportCount = [int]$runtimeLock.abi.exportCount
    $expectedExportSetSha256 = [string]$runtimeLock.abi.exportSetSha256
    $generatedIdentityHeaderRelativePath = ([string]$runtimeLock.runtimeIdentity.generatedIdentityHeader).Replace('/', '\')
    $expectedImportModules = [string[]]@($runtimeLock.abi.importModules)

    $textIdentityModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEngineTextIdentity.psm1'
    Import-Module $textIdentityModule -Force -ErrorAction Stop
    $sanitizedEnvironmentModule = Join-Path $repoRoot 'Tools\Modules\Build\SanitizedEnvironment.psm1'
    Import-Module $sanitizedEnvironmentModule -Force -ErrorAction Stop
    $terminalBuildRoot = [string]$buildPlan.StateRoot
    $sourcesRoot = Join-Path $terminalBuildRoot 'Sources'
    $downloadsRoot = Join-Path $terminalBuildRoot 'Downloads'
    $toolchainsRoot = Join-Path $terminalBuildRoot 'Toolchains'
    $cacheRoot = Join-Path $terminalBuildRoot 'Cache'
    $sourceRoot = Join-Path $sourcesRoot "ghostty-$ghosttyCommit"
    $zigArchive = Join-Path $downloadsRoot $zigArchiveName
    $zigExecutable = Join-Path $toolchainsRoot ([string]$zigArchiveLock.executableRelativePath).Replace('/', '\')
    $zigRoot = Split-Path -Parent $zigExecutable
    $productRoot = [string]$buildPlan.OutputRoot
    $identityPath = Join-Path $productRoot 'terminal-runtime-identity.json'
    $dependencyStates = @($dependencyLocks | ForEach-Object {
            [pscustomobject]@{
                Id = [string]$_.id
                Lock = $_
                SourceArchive = Join-Path $downloadsRoot ([string]$_.sourceArchive.fileName)
                ContentArchive = Join-Path $downloadsRoot ([string]$_.contentArchive.fileName)
                FetchCacheRoot = Join-Path $cacheRoot ("dependency-fetch-{0}-{1}" -f [string]$_.id, $zigVersion)
            }
        })
    $uucodeState = @($dependencyStates | Where-Object { $_.Id -ceq 'uucode' })[0]
    $uucodeSourceArchive = [string]$uucodeState.SourceArchive
    $uucodePackageArchive = [string]$uucodeState.ContentArchive
    $legacyUucodePackageArchive = Join-Path $cacheRoot "zig-global-$zigVersion\p\$uucodePackageName.tar.gz"
    $privateRuntimeOverlayPatch = Join-Path $repoRoot ([string]$runtimeLock.overlay.path).Replace('/', '\')
    if (-not (Test-Path -LiteralPath $privateRuntimeOverlayPatch -PathType Leaf)) {
        throw "Ghostty private-runtime Windows overlay patch is missing: $privateRuntimeOverlayPatch"
    }
    $privateRuntimeOverlayIdentity = Get-RSTerminalLfTextIdentity `
        -LiteralPath $privateRuntimeOverlayPatch
    $privateRuntimeOverlayPatchBytes = $privateRuntimeOverlayIdentity.Utf8Bytes
    $privateRuntimeOverlayPatchSha256 = $privateRuntimeOverlayIdentity.Sha256
    $normalizedPrivateRuntimeOverlayPatch = Join-Path $cacheRoot 'Ghostty-WindowsPrivateRuntime.lf.patch'
    $archiveModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEngineArchive.psm1'
    Import-Module $archiveModule -Force -ErrorAction Stop

    function Test-PathInsideTerminalBuildRoot {
        param([Parameter(Mandatory)][string]$Path)

        $fullPath = [IO.Path]::GetFullPath($Path)
        $rootWithSeparator = $terminalBuildRoot.TrimEnd('\') + '\'
        return $fullPath.StartsWith($rootWithSeparator, [StringComparison]::OrdinalIgnoreCase)
    }

    function Get-RSGhosttyRuntimeRepositoryMutexName {
        param(
            [Parameter(Mandatory)]
            [string]$RepoRoot
        )

        $normalizedRoot = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\').ToUpperInvariant()
        $sha256 = [Security.Cryptography.SHA256]::Create()
        try {
            $hashBytes = $sha256.ComputeHash([Text.Encoding]::UTF8.GetBytes($normalizedRoot))
        }
        finally {
            $sha256.Dispose()
        }

        $hash = ([BitConverter]::ToString($hashBytes)).Replace('-', '')
        return "Local\RedSalamander.GhosttyRuntime.Repository.$hash"
    }

    function Remove-TerminalBuildDirectory {
        param([Parameter(Mandatory)][string]$Path)

        $fullPath = [IO.Path]::GetFullPath($Path)
        if (-not (Test-PathInsideTerminalBuildRoot -Path $fullPath)) {
            throw "Refusing to remove a path outside the terminal build root: $fullPath"
        }
        if ($fullPath -eq $terminalBuildRoot) {
            throw 'Refusing to remove the terminal build root itself.'
        }
        if (Test-Path -LiteralPath $fullPath) {
            Remove-Item -LiteralPath $fullPath -Recurse -Force
        }
    }

    function Get-TerminalRuntimeIdentityHeaderText {
        param(
            [Parameter(Mandatory)][long]$SizeBytes,
            [Parameter(Mandatory)][string]$Sha256
        )

        if ($SizeBytes -le 0 -or $Sha256 -cnotmatch '^[0-9a-f]{64}$') {
            throw 'Cannot generate the Terminal runtime identity header from an invalid artifact identity.'
        }
        $header = @"
#pragma once

#include <cstdint>
#include <string_view>

namespace RSTerminalRuntimeIdentity
{
inline constexpr uint64_t kSizeBytes = $($SizeBytes)ULL;
inline constexpr std::wstring_view kSha256 = L"$Sha256";
} // namespace RSTerminalRuntimeIdentity
"@
        return $header.Replace("`r`n", "`n").Replace("`r", "`n") + "`n"
    }

    function Invoke-CheckedProcess {
        param(
            [Parameter(Mandatory)][string]$FilePath,
            [Parameter(Mandatory)][string[]]$ArgumentList,
            [Parameter(Mandatory)][string]$WorkingDirectory,
            [hashtable]$Environment = @{}
        )

        $startInfo = New-RSProcessStartInfo `
            -FilePath $FilePath `
            -Arguments $ArgumentList `
            -WorkingDirectory $WorkingDirectory `
            -AdditionalEnvironment $Environment
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true

        $process = $null
        try {
            $process = Start-RSContainedProcess -ProcessStartInfo $startInfo
            $stdoutTask = $process.StandardOutput.ReadToEndAsync()
            $stderrTask = $process.StandardError.ReadToEndAsync()
            $process.WaitForExit()
            $stdout = $stdoutTask.GetAwaiter().GetResult()
            $stderr = $stderrTask.GetAwaiter().GetResult()
            if ($process.ExitCode -ne 0) {
                throw "'$FilePath' failed with exit code $($process.ExitCode).`n$stderr"
            }
            return [pscustomobject]@{ Stdout = $stdout; Stderr = $stderr }
        }
        finally {
            Close-RSContainedProcess -Process $process
        }
    }

    function Test-GitApply {
        param(
            [Parameter(Mandatory)][string[]]$Arguments,
            [Parameter(Mandatory)][string]$WorkingDirectory
        )

        try {
            [void](Invoke-CheckedProcess -FilePath $gitCommand.Source -ArgumentList $Arguments -WorkingDirectory $WorkingDirectory)
            return $true
        } catch {
            return $false
        }
    }

    $repositoryMutexName = Get-RSGhosttyRuntimeRepositoryMutexName -RepoRoot $repoRoot
    $repositoryMutex = [Threading.Mutex]::new($false, $repositoryMutexName)
    $repositoryMutexAcquired = $false
    try {
        try {
            $repositoryMutexAcquired = $repositoryMutex.WaitOne([TimeSpan]::FromMinutes(10))
        } catch [Threading.AbandonedMutexException] {
            $repositoryMutexAcquired = $true
            Write-Warning "Recovered abandoned Ghostty runtime repository state lease '$repositoryMutexName'. Exact source-overlay and partial-output recovery checks will run before cached output is trusted."
        }
        if (-not $repositoryMutexAcquired) {
            throw "Timed out waiting for this repository's shared Ghostty runtime source, cache, and products: $repoRoot"
        }

    foreach ($directory in @($terminalBuildRoot, $sourcesRoot, $downloadsRoot, $toolchainsRoot, $cacheRoot)) {
        [void](New-Item -ItemType Directory -Path $directory -Force)
    }
    [IO.File]::WriteAllBytes($normalizedPrivateRuntimeOverlayPatch, $privateRuntimeOverlayPatchBytes)
    $actualNormalizedOverlaySha256 =
        (Get-FileHash -Algorithm SHA256 -LiteralPath $normalizedPrivateRuntimeOverlayPatch).Hash.ToLowerInvariant()
    if ($actualNormalizedOverlaySha256 -cne $privateRuntimeOverlayPatchSha256) {
        throw "Normalized Ghostty private-runtime overlay SHA-256 mismatch: $actualNormalizedOverlaySha256"
    }

    if (-not (Test-Path -LiteralPath $zigArchive -PathType Leaf)) {
        if ($Offline) {
            throw "Offline verification requires the pinned Zig archive cache: $zigArchive"
        }
        $temporaryArchive = "$zigArchive.partial"
        if (Test-Path -LiteralPath $temporaryArchive) {
            Remove-Item -LiteralPath $temporaryArchive -Force
        }
        $curlCommand = Get-Command curl.exe -ErrorAction Stop
        [void](Invoke-CheckedProcess `
                -FilePath $curlCommand.Source `
                -ArgumentList @('--fail', '--location', '--silent', '--show-error', '--output', $temporaryArchive, $zigArchiveUri) `
                -WorkingDirectory $downloadsRoot)
        Move-Item -LiteralPath $temporaryArchive -Destination $zigArchive
    }

    $actualZigSha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $zigArchive).Hash.ToLowerInvariant()
    if ($actualZigSha256 -cne $zigArchiveSha256) {
        throw "Zig archive SHA-256 mismatch: $actualZigSha256"
    }

    if (-not (Test-Path -LiteralPath $zigExecutable -PathType Leaf)) {
        Expand-Archive -LiteralPath $zigArchive -DestinationPath $toolchainsRoot -Force
    }
    $actualZigExecutableSha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $zigExecutable).Hash.ToLowerInvariant()
    if ($actualZigExecutableSha256 -cne $zigExecutableSha256) {
        throw "Pinned Zig executable SHA-256 mismatch: $actualZigExecutableSha256"
    }
    $actualZigVersion = (Invoke-CheckedProcess -FilePath $zigExecutable -ArgumentList @('version') -WorkingDirectory $repoRoot).Stdout.Trim()
    if ($actualZigVersion -cne $zigVersion) {
        throw "Expected Zig $zigVersion but found '$actualZigVersion'."
    }

    $gitCommand = Get-Command git.exe -ErrorAction Stop
    if (-not (Test-Path -LiteralPath (Join-Path $sourceRoot '.git') -PathType Container)) {
        if ($Offline) {
            throw "Offline verification requires the pinned Ghostty source checkout: $sourceRoot"
        }
        $partialSource = "$sourceRoot.partial"
        Remove-TerminalBuildDirectory -Path $partialSource
        [void](New-Item -ItemType Directory -Path $partialSource -Force)
        [void](Invoke-CheckedProcess -FilePath $gitCommand.Source -ArgumentList @('init', '--quiet') -WorkingDirectory $partialSource)
        # The pinned Ghostty tree contains corpus paths beyond legacy MAX_PATH.
        # Configure the private checkout before fetch/checkout so a clean Windows
        # build materializes every tracked input rather than silently deleting the
        # long-path subset and failing the integrity gate later.
        [void](Invoke-CheckedProcess -FilePath $gitCommand.Source -ArgumentList @('config', 'core.longpaths', 'true') -WorkingDirectory $partialSource)
        [void](Invoke-CheckedProcess -FilePath $gitCommand.Source -ArgumentList @('remote', 'add', 'origin', $ghosttyRepository) -WorkingDirectory $partialSource)
        [void](Invoke-CheckedProcess -FilePath $gitCommand.Source -ArgumentList @('fetch', '--quiet', '--depth', '1', 'origin', $ghosttyCommit) -WorkingDirectory $partialSource)
        [void](Invoke-CheckedProcess -FilePath $gitCommand.Source -ArgumentList @('checkout', '--quiet', '--detach', 'FETCH_HEAD') -WorkingDirectory $partialSource)
        Move-Item -LiteralPath $partialSource -Destination $sourceRoot
    }

    [void](Invoke-CheckedProcess -FilePath $gitCommand.Source -ArgumentList @('config', 'core.longpaths', 'true') -WorkingDirectory $sourceRoot)

    $actualCommit = (Invoke-CheckedProcess -FilePath $gitCommand.Source -ArgumentList @('rev-parse', 'HEAD') -WorkingDirectory $sourceRoot).Stdout.Trim()
    if ($actualCommit -cne $ghosttyCommit) {
        throw "Ghostty source pin mismatch: $actualCommit"
    }
    $actualTree = (Invoke-CheckedProcess -FilePath $gitCommand.Source -ArgumentList @('rev-parse', 'HEAD^{tree}') -WorkingDirectory $sourceRoot).Stdout.Trim()
    if ($actualTree -cne $ghosttyTree) {
        throw "Ghostty source tree mismatch: $actualTree"
    }
    $buildZonPath = Join-Path $sourceRoot 'build.zig.zon'
    $versionMatch = [regex]::Match([IO.File]::ReadAllText($buildZonPath), '(?m)^\s*\.version\s*=\s*"(?<version>[^"]+)"')
    if (-not $versionMatch.Success -or $versionMatch.Groups['version'].Value -cne $ghosttyVersion) {
        $actualVersion = if ($versionMatch.Success) { $versionMatch.Groups['version'].Value } else { '<missing>' }
        throw "Ghostty source version mismatch: expected '$ghosttyVersion', found '$actualVersion'."
    }
    # deps.files.ghostty.org serves source tarballs, while Zig stores deterministic,
    # path-filtered package archives under p/<content-hash>.tar.gz. Preserve both
    # identities for every locked build/runtime dependency so candidate Offline
    # builds cannot fetch an undeclared transitive input.
    if (-not (Test-Path -LiteralPath $uucodePackageArchive -PathType Leaf) -and
        (Test-Path -LiteralPath $legacyUucodePackageArchive -PathType Leaf) -and
        (Get-FileHash -Algorithm SHA256 -LiteralPath $legacyUucodePackageArchive).Hash.ToLowerInvariant() -ceq
            $uucodePackageArchiveSha256) {
        Copy-Item -LiteralPath $legacyUucodePackageArchive -Destination $uucodePackageArchive
    }
    $tarCommand = Get-Command tar.exe -ErrorAction Stop
    foreach ($dependencyState in $dependencyStates) {
        $dependency = $dependencyState.Lock
        $dependencyId = [string]$dependency.id
        $sourceArchive = [string]$dependencyState.SourceArchive
        $contentArchive = [string]$dependencyState.ContentArchive
        $sourceSha256 = [string]$dependency.sourceArchive.sha256
        $contentSha256 = [string]$dependency.contentArchive.sha256
        $packageName = [string]$dependency.packageName
        $licenseRelativePath = [string]$dependency.licenseArchiveRelativePath

        if (Test-Path -LiteralPath $contentArchive -PathType Leaf) {
            $actualContentSha256 =
                (Get-FileHash -Algorithm SHA256 -LiteralPath $contentArchive).Hash.ToLowerInvariant()
            if ($actualContentSha256 -cne $contentSha256) {
                if ($dependencyId -ceq 'uucode' -and
                    $actualContentSha256 -ceq $sourceSha256 -and
                    -not (Test-Path -LiteralPath $sourceArchive -PathType Leaf)) {
                    # Recover the representation written by the previous uucode
                    # bootstrap bug; exact source and content checks still run.
                    Move-Item -LiteralPath $contentArchive -Destination $sourceArchive
                }
                else {
                    throw "$dependencyId canonical package archive SHA-256 mismatch: $actualContentSha256"
                }
            }
        }

        if (-not (Test-Path -LiteralPath $contentArchive -PathType Leaf)) {
            if (-not (Test-Path -LiteralPath $sourceArchive -PathType Leaf)) {
                if ($Offline) {
                    throw "Offline verification requires the pinned $dependencyId package or source archive cache: $contentArchive"
                }
                $temporarySourceArchive = "$sourceArchive.partial"
                if (Test-Path -LiteralPath $temporarySourceArchive) {
                    Remove-Item -LiteralPath $temporarySourceArchive -Force
                }
                $curlCommand = Get-Command curl.exe -ErrorAction Stop
                [void](Invoke-CheckedProcess `
                        -FilePath $curlCommand.Source `
                        -ArgumentList @('--fail', '--location', '--silent', '--show-error', '--output', $temporarySourceArchive, [string]$dependency.sourceArchive.url) `
                        -WorkingDirectory $downloadsRoot)
                Move-Item -LiteralPath $temporarySourceArchive -Destination $sourceArchive
            }

            [void](Assert-RSTerminalEngineArchiveLayout `
                    -LiteralPath $sourceArchive `
                    -ExpectedSha256 $sourceSha256 `
                    -InternalRoot ([string]$dependency.sourceArchive.internalRoot) `
                    -RequiredFileRelativePath $licenseRelativePath `
                    -Label "$dependencyId source archive")

            $fetchCacheRoot = [string]$dependencyState.FetchCacheRoot
            Remove-TerminalBuildDirectory -Path $fetchCacheRoot
            [void](New-Item -ItemType Directory -Path $fetchCacheRoot -Force)
            try {
                $actualPackageName = (Invoke-CheckedProcess `
                        -FilePath $zigExecutable `
                        -ArgumentList @('fetch', '--global-cache-dir', $fetchCacheRoot, $sourceArchive) `
                        -WorkingDirectory $sourceRoot).Stdout.Trim()
                if ($actualPackageName -cne $packageName) {
                    throw "$dependencyId package content hash mismatch: $actualPackageName"
                }

                $canonicalPackageArchive = Join-Path $fetchCacheRoot "p\$packageName.tar.gz"
                if (-not (Test-Path -LiteralPath $canonicalPackageArchive -PathType Leaf)) {
                    throw "Zig did not materialize the canonical $dependencyId package archive: $canonicalPackageArchive"
                }
                $actualCanonicalSha256 =
                    (Get-FileHash -Algorithm SHA256 -LiteralPath $canonicalPackageArchive).Hash.ToLowerInvariant()
                if ($actualCanonicalSha256 -cne $contentSha256) {
                    throw "$dependencyId canonical package archive SHA-256 mismatch: $actualCanonicalSha256"
                }

                $temporaryPackageArchive = "$contentArchive.partial"
                if (Test-Path -LiteralPath $temporaryPackageArchive) {
                    Remove-Item -LiteralPath $temporaryPackageArchive -Force
                }
                Copy-Item -LiteralPath $canonicalPackageArchive -Destination $temporaryPackageArchive
                Move-Item -LiteralPath $temporaryPackageArchive -Destination $contentArchive
            }
            finally {
                Remove-TerminalBuildDirectory -Path $fetchCacheRoot
            }
        }

        [void](Assert-RSTerminalEngineArchiveLayout `
                -LiteralPath $contentArchive `
                -ExpectedSha256 $contentSha256 `
                -InternalRoot ([string]$dependency.contentArchive.internalRoot) `
                -RequiredFileRelativePath $licenseRelativePath `
                -Label "$dependencyId canonical package archive")
        $licenseText = (Invoke-CheckedProcess `
                -FilePath $tarCommand.Source `
                -ArgumentList @('-xOf', $contentArchive, "$($dependency.contentArchive.internalRoot)/$licenseRelativePath") `
                -WorkingDirectory $repoRoot).Stdout
        $licenseIdentity = Get-RSTerminalLfStringIdentity -Text $licenseText
        if ($licenseIdentity.Sha256 -cne [string]$dependency.licenseSha256) {
            throw "$dependencyId license semantic SHA-256 mismatch: $($licenseIdentity.Sha256)"
        }
        $dependencyState | Add-Member -NotePropertyName LicenseText -NotePropertyValue $licenseIdentity.Text
    }

    $privateRuntimeOverlayCanApply = Test-GitApply -Arguments @('apply', '--check', $normalizedPrivateRuntimeOverlayPatch) -WorkingDirectory $sourceRoot
    $privateRuntimeOverlayCanReverse = Test-GitApply -Arguments @('apply', '--reverse', '--check', $normalizedPrivateRuntimeOverlayPatch) -WorkingDirectory $sourceRoot
    if (-not $privateRuntimeOverlayCanApply -and $privateRuntimeOverlayCanReverse) {
        # Recover only the exact tracked overlay after an interrupted prior build.
        [void](Invoke-CheckedProcess `
                -FilePath $gitCommand.Source `
                -ArgumentList @('apply', '--reverse', $normalizedPrivateRuntimeOverlayPatch) `
                -WorkingDirectory $sourceRoot)
        $privateRuntimeOverlayCanApply = $true
    }
    if (-not $privateRuntimeOverlayCanApply) {
        throw 'The pinned Ghostty private-runtime Windows overlay no longer applies cleanly to the pinned source.'
    }
    $windowsUnmaterializablePaths = @(
        'test/fuzz-libghostty/corpus/parser-cmin/id_000123,time_0,execs_0,orig_id_000940,src_000935,time_9023,execs_524829,op_quick,pos_55,val_+13,+cov',
        'test/fuzz-libghostty/corpus/parser-cmin/id_000213,time_0,execs_0,orig_id_001278,src_001266,time_20982,execs_1128131,op_quick,pos_31,val_+2,+cov'
    )
    foreach ($relativePath in $windowsUnmaterializablePaths) {
        $materializedPath = Join-Path $sourceRoot $relativePath
        if (Test-Path -LiteralPath $materializedPath -PathType Leaf) {
            Remove-Item -LiteralPath $materializedPath -Force
        }
    }
    $allowedStatus = @($windowsUnmaterializablePaths | ForEach-Object { " D $_" })
    $sourceStatus = @(
        (Invoke-CheckedProcess `
                -FilePath $gitCommand.Source `
                -ArgumentList @('status', '--porcelain=v1', '--untracked-files=no') `
                -WorkingDirectory $sourceRoot).Stdout -split "`r?`n" | Where-Object { $_ -ne '' }
    )
    $unexpectedStatus = @($sourceStatus | Where-Object { $_ -cnotin $allowedStatus })
    if ($unexpectedStatus.Count -ne 0 -or $sourceStatus.Count -ne $allowedStatus.Count) {
        throw "Ghostty source checkout has unexpected tracked drift: $($sourceStatus -join '; ')"
    }

    $noticeStates = @($runtimeLock.notices | ForEach-Object {
            $notice = $_
            $identityKey = if ([string]$notice.id -ceq 'uucode-bjoern-hoehrmann') {
                'uucodeBjoernHoehrmann'
            }
            else {
                [string]$notice.id
            }
            $sourceKind = [string]$notice.source.kind
            $sourcePath = $null
            $text = $null
            if ($sourceKind -ceq 'dependency') {
                $dependencyState = @($dependencyStates | Where-Object {
                        $_.Id -ceq [string]$notice.source.dependencyId
                    })[0]
                $text = [string]$dependencyState.LicenseText
            }
            else {
                $sourceBase = switch ($sourceKind) {
                    'repository' { $repoRoot }
                    'upstream' { $sourceRoot }
                    'toolchain' { $zigRoot }
                    default { throw "Unsupported Terminal runtime notice source kind '$sourceKind'." }
                }
                $sourcePath = Join-Path $sourceBase ([string]$notice.source.path).Replace('/', '\')
                if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
                    throw "Terminal runtime license input is missing: $sourcePath"
                }
                $sourceIdentity = Get-RSTerminalLfTextIdentity -LiteralPath $sourcePath
                if ($sourceIdentity.Sha256 -cne [string]$notice.normalizedTextSha256) {
                    throw "Terminal runtime notice '$($notice.id)' SHA-256 mismatch: $($sourceIdentity.Sha256)"
                }
                $text = [string]$sourceIdentity.Text
            }
            [pscustomobject]@{
                Id = [string]$notice.id
                IdentityKey = $identityKey
                OutputRelativePath = ([string]$notice.outputRelativePath).Replace('/', '\')
                ExpectedSha256 = [string]$notice.normalizedTextSha256
                Text = $text
            }
        })
    $expectedLicenseHashes = [ordered]@{}
    $licenseRelativePaths = [ordered]@{}
    foreach ($noticeState in $noticeStates) {
        $expectedLicenseHashes[$noticeState.IdentityKey] = $noticeState.ExpectedSha256
        $licenseRelativePaths[$noticeState.IdentityKey] = $noticeState.OutputRelativePath
    }
    $dependencyIdentities = @($dependencyStates | ForEach-Object {
            $dependency = $_.Lock
            [ordered]@{
                id = [string]$dependency.id
                packageName = [string]$dependency.packageName
                sourceArchiveSha256 = [string]$dependency.sourceArchive.sha256
                contentArchiveSha256 = [string]$dependency.contentArchive.sha256
                licenseSha256 = [string]$dependency.licenseSha256
            }
        })
    $expectedIdentity = [ordered]@{
        formatId = [string]$runtimeLock.runtimeIdentity.formatId
        schemaVersion = [int]$runtimeLock.runtimeIdentity.schemaVersion
        ghosttyCommit = $ghosttyCommit
        zigVersion = $zigVersion
        zigArchiveSha256 = $zigArchiveSha256
        uucodePackageName = $uucodePackageName
        uucodeSourceArchiveUri = $uucodeSourceArchiveUri
        uucodeSourceArchiveCacheFileName = $uucodeSourceArchiveName
        uucodeSourceArchiveSha256 = $uucodeSourceArchiveSha256
        uucodeSourceInternalRoot = $uucodeSourceInternalRoot
        uucodePackageCacheFileName = $uucodePackageCacheFileName
        uucodePackageArchiveSha256 = $uucodePackageArchiveSha256
        uucodePackageInternalRoot = $uucodePackageInternalRoot
        uucodeLicenseArchiveRelativePath = $uucodeLicenseArchiveRelativePath
        uucodeLicenseSha256 = $uucodeLicenseSha256
        dependencies = $dependencyIdentities
        target = $target
        buildArguments = $buildArguments
        contentHashNormalization = $contentHashNormalization
        privateRuntimeOverlayPatchSha256 = $privateRuntimeOverlayPatchSha256
        pairingContract = $pairingContract
        generatedIdentityHeader = $generatedIdentityHeaderRelativePath.Replace('\', '/')
        dllNormalizedSha256 = $null
        dllSha256 = $null
        dllSizeBytes = $null
        machine = $expectedMachine
        importModules = $expectedImportModules
        exportCount = $expectedExportCount
        exportSetSha256 = $expectedExportSetSha256
        licenseSha256 = $expectedLicenseHashes
    }
    $uucodeIdentityLockNames = @(
        'uucodePackageName',
        'uucodeSourceArchiveUri',
        'uucodeSourceArchiveCacheFileName',
        'uucodeSourceArchiveSha256',
        'uucodeSourceInternalRoot',
        'uucodePackageCacheFileName',
        'uucodePackageArchiveSha256',
        'uucodePackageInternalRoot',
        'uucodeLicenseArchiveRelativePath',
        'uucodeLicenseSha256')
    $expectedDependencyIdentitySha256 = Get-RSCanonicalJsonDigest -InputObject $dependencyIdentities
    $peIdentityModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEnginePeIdentity.psm1'
    Import-Module $peIdentityModule -Force -ErrorAction Stop
    $importPolicyModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalRuntimeImportPolicy.psm1'
    Import-Module $importPolicyModule -Force -ErrorAction Stop
    $exportIdentityModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalRuntimeExportIdentity.psm1'
    Import-Module $exportIdentityModule -Force -ErrorAction Stop

    $canReuse = $false
    if (-not $Force -and
        (Test-Path -LiteralPath (Join-Path $productRoot 'bin\ghostty-vt.dll') -PathType Leaf) -and
        (Test-Path -LiteralPath (Join-Path $productRoot 'include\ghostty\vt.h') -PathType Leaf) -and
        (-not $buildPlan.WritePairingHeader -or
            (Test-Path -LiteralPath (Join-Path $productRoot $generatedIdentityHeaderRelativePath) -PathType Leaf)) -and
        (Test-Path -LiteralPath $identityPath -PathType Leaf)) {
        $currentIdentity = Get-Content -LiteralPath $identityPath -Raw | ConvertFrom-Json -Depth 20
        $currentExportCountProperty = $currentIdentity.PSObject.Properties['exportCount']
        $currentExportSetSha256Property = $currentIdentity.PSObject.Properties['exportSetSha256']
        $currentDependenciesProperty = $currentIdentity.PSObject.Properties['dependencies']
        $currentDependenciesMatch = $null -ne $currentDependenciesProperty -and
            (Get-RSCanonicalJsonDigest -InputObject @($currentDependenciesProperty.Value)) -ceq
                $expectedDependencyIdentitySha256
        $currentUucodeLocksMatch = $true
        foreach ($lockName in $uucodeIdentityLockNames) {
            $lockProperty = $currentIdentity.PSObject.Properties[$lockName]
            if ($null -eq $lockProperty -or
                [string]$lockProperty.Value -cne [string]$expectedIdentity[$lockName]) {
                $currentUucodeLocksMatch = $false
                break
            }
        }
        $canReuse = $currentIdentity.formatId -ceq $expectedIdentity.formatId -and
            [int]$currentIdentity.schemaVersion -eq [int]$runtimeLock.runtimeIdentity.schemaVersion -and
            $currentIdentity.ghosttyCommit -ceq $ghosttyCommit -and
            $currentIdentity.zigVersion -ceq $zigVersion -and
            $currentIdentity.zigArchiveSha256 -ceq $zigArchiveSha256 -and
            $currentUucodeLocksMatch -and
            $currentDependenciesMatch -and
            $currentIdentity.target -ceq $target -and
            (@($currentIdentity.buildArguments) -join "`0") -ceq ($buildArguments -join "`0") -and
            $currentIdentity.contentHashNormalization -ceq $contentHashNormalization -and
            $currentIdentity.privateRuntimeOverlayPatchSha256 -ceq $privateRuntimeOverlayPatchSha256 -and
            $currentIdentity.pairingContract -ceq $pairingContract -and
            $currentIdentity.generatedIdentityHeader -ceq $generatedIdentityHeaderRelativePath.Replace('\', '/') -and
            [string]$currentIdentity.dllNormalizedSha256 -cmatch '^[0-9a-f]{64}$' -and
            [string]$currentIdentity.dllSha256 -cmatch '^[0-9a-f]{64}$' -and
            [long]$currentIdentity.dllSizeBytes -ge $minimumDllSizeBytes -and
            [long]$currentIdentity.dllSizeBytes -le $maximumDllSizeBytes -and
            $currentIdentity.machine -ceq $expectedMachine -and
            (@($currentIdentity.importModules) -join "`0") -ceq ($expectedImportModules -join "`0") -and
            $null -ne $currentExportCountProperty -and
            [int]$currentExportCountProperty.Value -eq $expectedExportCount -and
            $null -ne $currentExportSetSha256Property -and
            [string]$currentExportSetSha256Property.Value -ceq $expectedExportSetSha256
        if ($canReuse) {
            foreach ($licenseName in $expectedLicenseHashes.Keys) {
                $identityLicenseProperty = $currentIdentity.licenseSha256.PSObject.Properties[$licenseName]
                if ($null -eq $identityLicenseProperty -or
                    [string]$identityLicenseProperty.Value -cne $expectedLicenseHashes[$licenseName]) {
                    $canReuse = $false
                    break
                }
            }
        }
        if ($canReuse) {
            $productDll = Join-Path $productRoot 'bin\ghostty-vt.dll'
            $productPeIdentity = Get-RSPeIdentity -Path $productDll
            $productImportModules = Get-RSTerminalRuntimeImportModules -Identity $productPeIdentity
            $productExportNames = [string[]]@(
                $productPeIdentity.semantic.exports | ForEach-Object { $_.name })
            $productExportSet = Get-RSTerminalExportSetIdentity -ExportNames $productExportNames
            $productExportsMatch = $productExportNames.Count -eq $expectedExports.Count -and
                @($expectedExports | Where-Object { $productExportNames -cnotcontains $_ }).Count -eq 0
            $canReuse = $productPeIdentity.machine -ceq $expectedMachine -and
                [long]$productPeIdentity.sizeBytes -eq [long]$currentIdentity.dllSizeBytes -and
                $productPeIdentity.normalizedSha256 -ceq [string]$currentIdentity.dllNormalizedSha256 -and
                $productPeIdentity.sha256 -ceq [string]$currentIdentity.dllSha256 -and
                (Test-RSTerminalRuntimeImportModules `
                    -Identity $productPeIdentity `
                    -ExpectedModules $expectedImportModules) -and
                $productExportsMatch -and
                [int]$productExportSet.count -eq $expectedExportCount -and
                $productExportSet.sha256 -ceq $expectedExportSetSha256
        }
        if ($canReuse -and $buildPlan.WritePairingHeader) {
            $expectedHeader = Get-TerminalRuntimeIdentityHeaderText `
                -SizeBytes ([long]$currentIdentity.dllSizeBytes) `
                -Sha256 ([string]$currentIdentity.dllSha256)
            $actualHeader = [IO.File]::ReadAllText((Join-Path $productRoot $generatedIdentityHeaderRelativePath))
            $canReuse = $actualHeader -ceq $expectedHeader
        }
        if ($canReuse) {
            foreach ($licenseName in $licenseRelativePaths.Keys) {
                $licensePath = Join-Path $productRoot $licenseRelativePaths[$licenseName]
                if (-not (Test-Path -LiteralPath $licensePath -PathType Leaf) -or
                    (Get-FileHash -Algorithm SHA256 -LiteralPath $licensePath).Hash.ToLowerInvariant() -cne $expectedLicenseHashes[$licenseName]) {
                    $canReuse = $false
                    break
                }
            }
        }
    }

    if (-not $canReuse) {
        $partialProduct = "$productRoot.partial"
        Remove-TerminalBuildDirectory -Path $partialProduct
        [void](New-Item -ItemType Directory -Path $partialProduct -Force)
        Remove-TerminalBuildDirectory -Path (Join-Path $sourceRoot '.zig-cache')

        $reproducibleCacheName = "ghostty-$ghosttyCommit-$Platform-reproducible-v1"
        $reproducibleCacheRoot = Join-Path $cacheRoot $reproducibleCacheName
        Remove-TerminalBuildDirectory -Path $reproducibleCacheRoot
        $reproducibleGlobalCache = Join-Path $reproducibleCacheRoot 'global'
        $reproducibleLocalCache = Join-Path $reproducibleCacheRoot 'local'
        $reproduciblePackageCache = Join-Path $reproducibleGlobalCache 'p'
        [void](New-Item -ItemType Directory -Path $reproduciblePackageCache -Force)
        [void](New-Item -ItemType Directory -Path $reproducibleLocalCache -Force)
        foreach ($dependencyState in $dependencyStates) {
            Copy-Item `
                -LiteralPath ([string]$dependencyState.ContentArchive) `
                -Destination (Join-Path $reproduciblePackageCache "$($dependencyState.Lock.packageName).tar.gz")
        }

        $canonicalDriveMutexName = 'Local\RedSalamander.GhosttyRuntime.CanonicalDrive.R'
        $canonicalDriveMutex = [Threading.Mutex]::new($false, $canonicalDriveMutexName)
        $canonicalDriveMutexAcquired = $false
        $canonicalDriveMutexWasAbandoned = $false
        try {
            try {
                $canonicalDriveMutexAcquired = $canonicalDriveMutex.WaitOne([TimeSpan]::FromMinutes(10))
            }
            catch [Threading.AbandonedMutexException] {
                $canonicalDriveMutexAcquired = $true
                $canonicalDriveMutexWasAbandoned = $true
            }
            if (-not $canonicalDriveMutexAcquired) {
                throw "Timed out waiting for the session-wide Ghostty canonical build-drive lease '$canonicalDriveMutexName'. Cache-hit validation and other worktrees that do not rebuild remain independent."
            }

            $mappingCreated = $false
            $privateRuntimeOverlayApplied = $false
            try {
                if (Test-Path -LiteralPath "$canonicalBuildDrive\") {
                    $abandonedContext = if ($canonicalDriveMutexWasAbandoned) {
                        ' The lease was abandoned, so an interrupted prior builder may have left this mapping.'
                    }
                    else { '' }
                    throw ("$canonicalBuildDrive must be free while rebuilding the reproducible Ghostty runtime." +
                           "$abandonedContext No mapping was changed. Inspect it with 'subst $canonicalBuildDrive'; " +
                           "after verifying no process uses that path, remove it with 'subst $canonicalBuildDrive /D' and retry.")
                }

                $substCommand = Get-Command subst.exe -ErrorAction Stop
                [void](Invoke-CheckedProcess `
                        -FilePath $substCommand.Source `
                        -ArgumentList @($canonicalBuildDrive, $repoRoot) `
                        -WorkingDirectory $repoRoot)
                $mappingCreated = $true

                [void](Invoke-CheckedProcess `
                        -FilePath $gitCommand.Source `
                        -ArgumentList @('apply', $normalizedPrivateRuntimeOverlayPatch) `
                        -WorkingDirectory $sourceRoot)
                $privateRuntimeOverlayApplied = $true
                $mappedRepoRoot = "$canonicalBuildDrive\"
                $mappedSourceRoot = Join-Path $mappedRepoRoot ([IO.Path]::GetRelativePath($repoRoot, $sourceRoot))
                $mappedZigExecutable = Join-Path $mappedRepoRoot ([IO.Path]::GetRelativePath($repoRoot, $zigExecutable))
                $mappedPartialProduct = Join-Path $mappedRepoRoot ([IO.Path]::GetRelativePath($repoRoot, $partialProduct))
                $mappedReproducibleCacheRoot = Join-Path $mappedRepoRoot ([IO.Path]::GetRelativePath($repoRoot, $reproducibleCacheRoot))
                $mappedGlobalCache = Join-Path $mappedReproducibleCacheRoot 'global'
                $mappedLocalCache = Join-Path $mappedReproducibleCacheRoot 'local'
                $arguments = @('build') + $buildArguments + @(
                    '--cache-dir', $mappedLocalCache,
                    '--global-cache-dir', $mappedGlobalCache,
                    '--prefix', $mappedPartialProduct,
                    "-Dtarget=$target"
                )
                [void](Invoke-CheckedProcess `
                        -FilePath $mappedZigExecutable `
                        -ArgumentList $arguments `
                        -WorkingDirectory $mappedSourceRoot)
            }
            finally {
                try {
                    if ($privateRuntimeOverlayApplied) {
                        [void](Invoke-CheckedProcess `
                                -FilePath $gitCommand.Source `
                                -ArgumentList @('apply', '--reverse', $normalizedPrivateRuntimeOverlayPatch) `
                                -WorkingDirectory $sourceRoot)
                    }
                }
                finally {
                    try {
                        if ($mappingCreated) {
                            [void](Invoke-CheckedProcess `
                                    -FilePath $substCommand.Source `
                                    -ArgumentList @($canonicalBuildDrive, '/D') `
                                    -WorkingDirectory $repoRoot)
                        }
                    }
                    finally {
                        Remove-TerminalBuildDirectory -Path $reproducibleCacheRoot
                    }
                }
            }
        }
        finally {
            if ($canonicalDriveMutexAcquired) {
                $canonicalDriveMutex.ReleaseMutex()
            }
            $canonicalDriveMutex.Dispose()
        }

        $dllPath = Join-Path $partialProduct 'bin\ghostty-vt.dll'
        if (-not (Test-Path -LiteralPath $dllPath -PathType Leaf)) {
            throw "Ghostty build did not produce '$dllPath'."
        }

        $peIdentity = Get-RSPeIdentity -Path $dllPath
        if ($peIdentity.machine -cne $expectedMachine) {
            throw "Ghostty runtime machine mismatch: $($peIdentity.machine)"
        }
        if ([long]$peIdentity.sizeBytes -lt $minimumDllSizeBytes -or
            [long]$peIdentity.sizeBytes -gt $maximumDllSizeBytes) {
            throw "Ghostty runtime size $($peIdentity.sizeBytes) is outside the reviewed $minimumDllSizeBytes..$maximumDllSizeBytes byte production envelope."
        }
        $exportNames = [string[]]@($peIdentity.semantic.exports | ForEach-Object { $_.name })
        $exportSet = Get-RSTerminalExportSetIdentity -ExportNames $exportNames
        $exportsMatch = $exportNames.Count -eq $expectedExports.Count -and
            @($expectedExports | Where-Object { $exportNames -cnotcontains $_ }).Count -eq 0
        if (-not $exportsMatch -or
            [int]$exportSet.count -ne $expectedExportCount -or
            $exportSet.sha256 -cne $expectedExportSetSha256) {
            throw "Ghostty runtime export set changed: count=$($exportSet.count), SHA-256=$($exportSet.sha256). Missing, duplicate, and additional exports are rejected."
        }
        foreach ($requiredExport in $requiredExports) {
            if ($exportNames -cnotcontains $requiredExport) {
                throw "Ghostty runtime lacks required export '$requiredExport'."
            }
        }
        $importModules = Get-RSTerminalRuntimeImportModules -Identity $peIdentity
        if (-not (Test-RSTerminalRuntimeImportModules `
                -Identity $peIdentity `
                -ExpectedModules $expectedImportModules)) {
            throw "Ghostty runtime has an unexpected private dependency: $($importModules -join ', ')"
        }

        $licenseDirectory = Join-Path $partialProduct 'share\licenses\terminal-runtime'
        [void](New-Item -ItemType Directory -Path $licenseDirectory -Force)
        foreach ($noticeState in $noticeStates) {
            $licensePath = Join-Path $partialProduct $noticeState.OutputRelativePath
            [void](New-Item -ItemType Directory -Path (Split-Path -Parent $licensePath) -Force)
            [IO.File]::WriteAllText(
                $licensePath,
                [string]$noticeState.Text,
                [Text.UTF8Encoding]::new($false))
        }
        foreach ($licenseName in $licenseRelativePaths.Keys) {
            $licensePath = Join-Path $partialProduct $licenseRelativePaths[$licenseName]
            $actualLicenseHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $licensePath).Hash.ToLowerInvariant()
            if ($actualLicenseHash -cne $expectedLicenseHashes[$licenseName]) {
                throw "Terminal runtime license '$licenseName' SHA-256 mismatch: $actualLicenseHash"
            }
        }
        if ($buildPlan.WritePairingHeader) {
            $generatedIdentityHeader = Join-Path $partialProduct $generatedIdentityHeaderRelativePath
            [void](New-Item -ItemType Directory -Path (Split-Path $generatedIdentityHeader -Parent) -Force)
            [IO.File]::WriteAllText(
                $generatedIdentityHeader,
                (Get-TerminalRuntimeIdentityHeaderText -SizeBytes ([long]$peIdentity.sizeBytes) -Sha256 $peIdentity.sha256),
                [Text.UTF8Encoding]::new($false))
        }
        $expectedIdentity.dllNormalizedSha256 = $peIdentity.normalizedSha256
        $expectedIdentity.dllSha256 = $peIdentity.sha256
        $expectedIdentity.dllSizeBytes = [string]$peIdentity.sizeBytes
        $expectedIdentity | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath (Join-Path $partialProduct 'terminal-runtime-identity.json') -Encoding utf8NoBOM

        Assert-RSGhosttyRuntimeOutputReplaceable -OutputRoot $productRoot
        Remove-TerminalBuildDirectory -Path $productRoot
        Move-Item -LiteralPath $partialProduct -Destination $productRoot
    }

    $productDll = Join-Path $productRoot 'bin\ghostty-vt.dll'
    $result = Get-Content -LiteralPath $identityPath -Raw | ConvertFrom-Json -Depth 20
    $actualProductIdentity = Get-RSPeIdentity -Path $productDll
    $actualProductImportModules = Get-RSTerminalRuntimeImportModules -Identity $actualProductIdentity
    if ($actualProductIdentity.machine -cne $expectedMachine -or
        $actualProductIdentity.normalizedSha256 -cne [string]$result.dllNormalizedSha256 -or
        $actualProductIdentity.sha256 -cne [string]$result.dllSha256 -or
        [long]$actualProductIdentity.sizeBytes -ne [long]$result.dllSizeBytes -or
        [long]$actualProductIdentity.sizeBytes -lt $minimumDllSizeBytes -or
        [long]$actualProductIdentity.sizeBytes -gt $maximumDllSizeBytes) {
        throw "Cached Ghostty runtime differs from its atomic plugin-pair identity: size=$($actualProductIdentity.sizeBytes) raw SHA-256=$($actualProductIdentity.sha256) normalized SHA-256=$($actualProductIdentity.normalizedSha256)"
    }
    if (-not (Test-RSTerminalRuntimeImportModules `
            -Identity $actualProductIdentity `
            -ExpectedModules $expectedImportModules) -or
        (@($result.importModules) -join "`0") -cne ($expectedImportModules -join "`0")) {
        throw "Cached Ghostty runtime has an unexpected private dependency: $($actualProductImportModules -join ', ')"
    }
    foreach ($lockName in $uucodeIdentityLockNames) {
        if ([string]$result.$lockName -cne [string]$expectedIdentity[$lockName]) {
            throw "Cached Ghostty runtime identity has the wrong '$lockName' lock."
        }
    }
    $resultDependenciesProperty = $result.PSObject.Properties['dependencies']
    if ($null -eq $resultDependenciesProperty -or
        (Get-RSCanonicalJsonDigest -InputObject @($resultDependenciesProperty.Value)) -cne
            $expectedDependencyIdentitySha256) {
        throw 'Cached Ghostty runtime identity has the wrong dependency lock set.'
    }
    if ($buildPlan.WritePairingHeader) {
        $generatedIdentityHeader = Join-Path $productRoot $generatedIdentityHeaderRelativePath
        $expectedHeader = Get-TerminalRuntimeIdentityHeaderText `
            -SizeBytes ([long]$actualProductIdentity.sizeBytes) `
            -Sha256 $actualProductIdentity.sha256
        if (-not (Test-Path -LiteralPath $generatedIdentityHeader -PathType Leaf) -or
            [IO.File]::ReadAllText($generatedIdentityHeader) -cne $expectedHeader) {
            throw 'The generated Terminal runtime identity header does not match the exact private DLL.'
        }
    }
    $finalExportNames = [string[]]@($actualProductIdentity.semantic.exports | ForEach-Object { $_.name })
    $finalExportSet = Get-RSTerminalExportSetIdentity -ExportNames $finalExportNames
    $finalExportsMatch = $finalExportNames.Count -eq $expectedExports.Count -and
        @($expectedExports | Where-Object { $finalExportNames -cnotcontains $_ }).Count -eq 0
    if (-not $finalExportsMatch -or
        [int]$finalExportSet.count -ne $expectedExportCount -or
        $finalExportSet.sha256 -cne $expectedExportSetSha256) {
        throw 'Cached Ghostty runtime exact export set differs from the product lock.'
    }
    foreach ($licenseName in $licenseRelativePaths.Keys) {
        $licensePath = Join-Path $productRoot $licenseRelativePaths[$licenseName]
        if (-not (Test-Path -LiteralPath $licensePath -PathType Leaf)) {
            throw "Cached Ghostty runtime license is missing: $licensePath"
        }
        $actualLicenseHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $licensePath).Hash.ToLowerInvariant()
        if ($actualLicenseHash -cne $expectedLicenseHashes[$licenseName]) {
            throw "Cached Ghostty runtime license '$licenseName' differs from the product lock: $actualLicenseHash"
        }
        $identityLicenseProperty = $result.licenseSha256.PSObject.Properties[$licenseName]
        if ($null -eq $identityLicenseProperty -or
            [string]$identityLicenseProperty.Value -cne $expectedLicenseHashes[$licenseName]) {
            throw "Cached Ghostty runtime identity has the wrong '$licenseName' license lock."
        }
    }
    if ($PassThru) {
        $result
    }
    else {
        Write-Host "Ghostty terminal runtime ready: mode=$($buildPlan.Mode) platform=$Platform pairedRaw=$($result.dllSha256) normalizedDiagnostic=$($result.dllNormalizedSha256)"
    }
    } finally {
        if ($repositoryMutexAcquired) {
            $repositoryMutex.ReleaseMutex()
        }
        $repositoryMutex.Dispose()
    }


}

Export-ModuleMember -Function @(
    'Read-RSGhosttyRuntimeLock',
    'Get-RSGhosttyRuntimeLockDigest',
    'Get-RSGhosttyRuntimeLockPlatform',
    'Get-RSGhosttyRuntimeLockDependency',
    'Resolve-RSGhosttyRuntimeBuildPlan',
    'Assert-RSGhosttyRuntimeOutputReplaceable',
    'Invoke-RSGhosttyRuntimeBuild')
