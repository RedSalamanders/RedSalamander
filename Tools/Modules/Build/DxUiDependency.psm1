# Product policy around the canonical library's restore identity and imports.
Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'ArtifactOperationLock.psm1') -Scope Local
Import-Module (Join-Path $PSScriptRoot 'ProcessStreaming.psm1') -Scope Local

# The restore leaves out what the library keeps for its own development: retained measurements, specifications and the
# generated gallery sheets are most of its tracked files and none of its consumer interface. Omitted paths carry the
# skip-worktree bit, so `git status` stays clean and the clean-source check below still holds.
$script:RSDxUiSparseCheckoutPatterns = @('/*', '!/Measurements/', '!/docs/gallery/', '!/Specs/')
# What the restore and the build read from the checkout. A sparse checkout that lacks one of them has failed.
$script:RSDxUiRequiredSourcePaths = @('capabilities.json', 'vcpkg-install.ps1', 'Tools/ConsumerBuild.psm1',
    'Build/DxUi.Consumer.props', 'Build/DxUi.Consumer.targets', 'src/DxUi.vcxproj', 'include')

function Get-RSDxUiPin {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RepoRoot)
    $path = Join-Path $RepoRoot 'Dependencies/DxUi.lock.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    $pin = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    # The revision the lock records is the one this product has adopted. Restore compares it with what the pinned source
    # reports (its capabilities.json, through the identity), so a pin that moves across a revision fails there.
    if ($pin.repository -cne 'https://github.com/RedSalamanders/DxUi' -or $pin.commit -cnotmatch '^[0-9a-f]{40}$' -or
        -not ($pin.apiRevision -is [int] -or $pin.apiRevision -is [long]) -or $pin.apiRevision -lt 1 -or
        @($pin.targets).Count -ne 1 -or $pin.targets[0] -cne 'DxUi') {
        throw 'DxUi.lock.json must name the canonical repository, exact commit, a positive integer API revision and single DxUi target.'
    }
    return $pin
}

function Get-RSDxUiRecordedApiRevision {
    # The API revision an identity records: a dictionary key in a fresh identity, a property in its JSON file. $null when absent.
    param([Parameter(Mandatory)][AllowNull()]$Identity)
    if ($null -eq $Identity) { return $null }
    if ($Identity -is [System.Collections.IDictionary]) { return $Identity['apiRevision'] }
    $property = $Identity.PSObject.Properties['apiRevision']
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Assert-RSDxUiApiRevision {
    # A revision change is a consumer-facing change (the library's rule: a consumer built against the previous revision could
    # stop compiling, linking or behaving as before). The lock records the revision this product adopted; the pinned source
    # states the one it implements. They differ when a pin moved without the adaptation the revision's change note lists.
    param([Parameter(Mandatory)]$Pin, [Parameter(Mandatory)][AllowNull()]$Recorded, [Parameter(Mandatory)][string]$Context)
    if ($null -eq $Recorded -or $Recorded -ne $Pin.apiRevision) {
        $reported = if ($null -eq $Recorded) { 'no API revision' } else { "API revision $Recorded" }
        throw ("$Context reports $reported, but Dependencies/DxUi.lock.json records API revision $($Pin.apiRevision). " +
            'Read CHANGELOG.md and any pending Changes/ notes in the pinned DxUi library, adapt the product to them, then set apiRevision in the lock.')
    }
}

function Get-RSDxUiSourceIdentity {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RepoRoot, [switch]$AllowMissing)
    $pin = Get-RSDxUiPin -RepoRoot $RepoRoot
    if ($null -eq $pin) { return $null }
    $source = Resolve-RSGuardedRepositoryPath -RepoRoot $RepoRoot -GuardedRelativeRoot '.build/dependencies/DxUi' `
        -Path (Join-Path $RepoRoot ".build/dependencies/DxUi/source/$($pin.commit)")
    if (-not (Test-Path -LiteralPath $source)) {
        if ($AllowMissing) { return $null }
        throw 'The pinned DxUi source is missing. Run Tools/Restore-DxUi.ps1 for this profile.'
    }
    $head = (& git -C $source rev-parse HEAD).Trim()
    if ($LASTEXITCODE -ne 0 -or $head -cne $pin.commit) { throw 'The restored DxUi checkout does not match the product pin.' }
    $dirty = & git -C $source status --porcelain --untracked-files=normal
    if ($LASTEXITCODE -ne 0 -or $dirty) { throw 'The restored DxUi source is dirty; build and test receipt reuse is forbidden.' }
    $tree = (& git -C $source rev-parse HEAD:include).Trim()
    if ($LASTEXITCODE -ne 0) { throw 'Cannot identify the pinned DxUi public headers.' }
    return [pscustomobject]@{ Pin=$pin; Source=$source; PublicHeadersGitTree=$tree; DisableStlAnnotations=$true }
}

function Get-RSDxUiOutputKey {
    param([Parameter(Mandatory)][string]$Fingerprint)
    # The directory is only a lookup key: DxUi.identity.<Platform>.json and module sidecars keep the full
    # fingerprint. A 64-bit prefix stays collision-free and shortens the deep vcpkg build-tree paths.
    if ($Fingerprint -cnotmatch '^[0-9a-f]{64}$') { throw 'The DxUi build fingerprint must be a lowercase SHA-256 digest.' }
    return $Fingerprint.Substring(0,16)
}

function Invoke-RSDxUiRestoreProcess {
    param([string]$RepoRoot, [string]$FilePath, [string[]]$Arguments)
    $logs = Join-Path $RepoRoot '.build/logs'
    [void](New-Item -ItemType Directory -Path $logs -Force)
    $log = Join-Path $logs ('dxui-restore-' + [guid]::NewGuid().ToString('N') + '.log')
    $code = Invoke-RSStreamingProcess -FilePath $FilePath -Arguments $Arguments -WorkingDirectory $RepoRoot -LogPath $log `
        -OutputLineCallback { param($Line, $IsError) Write-Host $Line }
    if ($code -ne 0) { throw "DxUi restore failed ($code). See $log" }
}

function Install-RSDxUiPinnedSource {
    <# Clones the library at exactly Commit into Source, with Git long paths and without the paths the restore does not need
       (RSDxUiSparseCheckoutPatterns). Private: Restore-RSDxUiDependency passes the lock's repository, and the restore test
       passes a local origin from a deep root. A failure removes the partial checkout this call created, so the next restore
       clones again instead of reusing a half-written source that fails its clean-source check. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [Parameter(Mandatory)][string]$Repository,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{40}$')][string]$Commit,
        [Parameter(Mandatory)][string]$Source
    )
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $Source) -Force)
    try {
        # A consumer root can be deep enough that library paths, and Git's own template hooks, pass MAX_PATH. Two settings
        # serve two moments: the process-level -c covers `git init`, which copies the hooks before the new repository's
        # configuration exists, and clone's own -c stores the setting in that configuration for every later git command on
        # this checkout (the clean-source check). No user or global Git configuration changes.
        Invoke-RSDxUiRestoreProcess $RepoRoot 'git.exe' @('-c', 'core.longpaths=true', 'clone', '-c', 'core.longpaths=true', '--no-checkout', $Repository, $Source)
        Invoke-RSDxUiRestoreProcess $RepoRoot 'git.exe' (@('-C', $Source, 'sparse-checkout', 'set', '--no-cone') + $script:RSDxUiSparseCheckoutPatterns)
        Invoke-RSDxUiRestoreProcess $RepoRoot 'git.exe' @('-C', $Source, '-c', 'core.longpaths=true', 'checkout', '--detach', $Commit)
        # git checkout can name files it could not write and still exit 0; the status shows each as a deletion.
        $dirty = @(& git -C $Source status --porcelain --untracked-files=normal)
        if ($LASTEXITCODE -ne 0 -or $dirty.Count -gt 0) {
            throw "The DxUi checkout is incomplete: $($dirty.Count) path(s) were not written, such as '$($dirty | Select-Object -First 1)'."
        }
        foreach ($required in $script:RSDxUiRequiredSourcePaths) {
            if (-not (Test-Path -LiteralPath (Join-Path $Source $required))) {
                throw "The sparse DxUi checkout lacks '$required'; its patterns must keep everything the restore and the build read."
            }
        }
    }
    catch {
        $failure = $_
        # Remove only what this call created, and only beneath the guarded dependency root.
        try {
            $guarded = Resolve-RSGuardedRepositoryPath -RepoRoot $RepoRoot -GuardedRelativeRoot '.build/dependencies/DxUi' -Path $Source
            if (Test-Path -LiteralPath $guarded) { Remove-Item -LiteralPath $guarded -Recurse -Force -ErrorAction Stop }
        }
        catch { Write-Warning "The partial DxUi checkout could not be removed; delete '$Source' before restoring again. $($_.Exception.Message)" }
        throw $failure
    }
}

function Get-RSDxUiConsumerCompilerHost {
    param([string]$RepoRoot, [string]$MSBuildPath, [string]$Platform, [string]$Configuration)
    # Restore cannot depend on BuildEvidence, which imports this module. Use its same first-party probe project.
    $hostOutput = & $MSBuildPath (Join-Path $RepoRoot 'RedSalamander/RedSalamander.vcxproj') /nologo `
        "/p:Configuration=$Configuration" "/p:Platform=$Platform" '-getProperty:PreferredToolArchitecture'
    if ($LASTEXITCODE -ne 0) { throw ('Cannot evaluate the consumer compiler host before DxUi restore: ' + ($hostOutput -join "`n")) }
    $compilerHost = ($hostOutput -join "`n").Trim()
    if ($compilerHost -notin @('x86','x64','arm64')) { throw "Unsupported consumer compiler host: $compilerHost" }
    return $compilerHost
}

function Get-RSDxUiUpdateNoticePresentation {
    [CmdletBinding()]
    param([AllowNull()][string]$Notice)

    if ([string]::IsNullOrWhiteSpace($Notice)) { return $null }

    # The pinned DxUi helper owns the read-only update decision. This consumer adapter
    # owns only the interactive severity presentation and product upgrade commands.
    $foregroundColor = if ($Notice -like 'DxUi main * has no successful completed validation;*') {
        [ConsoleColor]::Red
    } elseif ($Notice -like 'DxUi update available:*') {
        [ConsoleColor]::Yellow
    } else {
        [ConsoleColor]::DarkYellow
    }
    $displayNotice = if ($Notice -like 'DxUi update available:*') {
        $summary = $Notice -replace ' \. Update .* on a branch and run the product regressions\.$', ''
        "$summary.`nTo upgrade and run local validation: .\Tools\Update-DxUi.ps1`nTo upgrade only after equivalent product validation passed elsewhere: .\Tools\Update-DxUi.ps1 -UpdateOnly"
    }
    else {
        $Notice
    }
    return [pscustomobject]@{ Notice = $displayNotice; ForegroundColor = $foregroundColor }
}

function Restore-RSDxUiDependency {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [Parameter(Mandatory)][string]$MSBuildPath,
        [Parameter(Mandatory)][ValidateSet('x64','ARM64')][string]$Platform,
        [Parameter(Mandatory)][ValidateSet('Debug','Release','ASan Debug')][string]$Configuration,
        [switch]$CheckUpdates,
        [switch]$Rebuild,
        [string]$ProjectName
    )
    $ErrorActionPreference = 'Stop'
    $pin = Get-RSDxUiPin -RepoRoot $RepoRoot
    if ($null -eq $pin) { throw 'The product DxUi lock is missing.' }
    $scope = @{configuration=$Configuration;platform=$Platform;coordination='shared-dependency'}
    $lease = Enter-RSArtifactOperationLock -RepoRoot $RepoRoot -Operation 'Restore pinned DxUi' -Scope $scope
    try {
        if ($lease.WasAbandoned) {
            [void](Set-RSArtifactOperationContaminated -RepoRoot $RepoRoot -Scope $scope -AbandonedOwner $lease.AbandonedOwner `
                -Reason 'The previous DxUi restore owner exited while writing shared dependency inputs.')
        }
        $contamination = Read-RSArtifactOperationContamination -RepoRoot $RepoRoot -Scope $scope
        if ($null -ne $contamination -and -not (Test-RSArtifactOperationRepairAllowed -Contamination $contamination `
            -Rebuild:$Rebuild -ProjectName $ProjectName -Configuration $Configuration -Platform $Platform)) {
            throw "The shared dependency scope requires a full-solution build.ps1 -Rebuild for platform $Platform before restoring DxUi."
        }
        # The root build keeps the marker until its complete rebuild and receipt publication succeed.
        Assert-RSNoResidualArtifactToolProcesses -RepoRoot $RepoRoot -Scope $scope
        $dependency = Resolve-RSGuardedRepositoryPath -RepoRoot $RepoRoot -GuardedRelativeRoot '.build/dependencies/DxUi' `
            -Path (Join-Path $RepoRoot '.build/dependencies/DxUi') -CreateDirectory
        $source = Join-Path $dependency "source/$($pin.commit)"
        if (-not (Test-Path -LiteralPath $source)) {
            Install-RSDxUiPinnedSource -RepoRoot $RepoRoot -Repository "$($pin.repository).git" -Commit $pin.commit -Source $source
        }
        $sourceIdentity = Get-RSDxUiSourceIdentity -RepoRoot $RepoRoot
        Import-Module (Join-Path $source 'Tools/ConsumerBuild.psm1') -Scope Local
        $compilerHost = Get-RSDxUiConsumerCompilerHost -RepoRoot $RepoRoot -MSBuildPath $MSBuildPath -Platform $Platform -Configuration $Configuration
        $previousHost = $env:PreferredToolArchitecture
        try {
            # The pinned helper accepts the standard MSBuild environment property. Limit it to identity evaluation.
            $env:PreferredToolArchitecture = $compilerHost
            $identity = Get-DxUiConsumerBuildIdentity -DxUiRoot $source -MSBuildPath $MSBuildPath -Platform $Platform -DisableStlAnnotations
        } finally { $env:PreferredToolArchitecture = $previousHost }
        # A changed compiler policy owns a separate writable archive output, even when the API/source pin stays unchanged.
        [xml]$msvcPolicy = Get-Content -Raw -LiteralPath (Join-Path $RepoRoot 'Directory.Build.props')
        $requiredTools = $msvcPolicy.SelectSingleNode('/Project/PropertyGroup/RequiredMsvcToolsVersion').InnerText
        if (-not ([string]$identity.Identity.vcToolsVersion).StartsWith($requiredTools + '.', [StringComparison]::Ordinal)) {
            throw "This product requires MSVC $requiredTools.x; selected $($identity.Identity.vcToolsVersion)."
        }
        $policyIdentity = [ordered]@{}
        if ($identity.Identity -is [Collections.IDictionary]) {
            foreach ($entry in $identity.Identity.GetEnumerator()) { $policyIdentity[$entry.Key] = $entry.Value }
        } else {
            foreach ($entry in $identity.Identity.PSObject.Properties) { $policyIdentity[$entry.Name] = $entry.Value }
        }
        $policyIdentity['releaseOptimizationPolicy'] = 'Gy-REF-GL-LTCG-full-v1'
        $policyHasher = [Security.Cryptography.SHA256]::Create()
        try { $policyFingerprint = ([BitConverter]::ToString($policyHasher.ComputeHash([Text.Encoding]::UTF8.GetBytes(($policyIdentity | ConvertTo-Json -Depth 8 -Compress))))).Replace('-','').ToLowerInvariant() }
        finally { $policyHasher.Dispose() }
        $identity = [pscustomobject]@{ Fingerprint=$policyFingerprint; Identity=$policyIdentity }
        # Before the library's own dependency restore: a pin that crossed an API revision is adopted deliberately.
        Assert-RSDxUiApiRevision -Pin $pin -Recorded (Get-RSDxUiRecordedApiRevision $identity.Identity) -Context 'The pinned DxUi source'
        $output = (Join-Path $dependency (Get-RSDxUiOutputKey -Fingerprint $identity.Fingerprint)) + [IO.Path]::DirectorySeparatorChar
        Invoke-RSDxUiRestoreProcess $RepoRoot 'powershell.exe' @('-NoProfile','-ExecutionPolicy','Bypass','-File',
            (Join-Path $source 'vcpkg-install.ps1'),'-Platform',$Platform,'-OutputRoot',$output)
        $values = [ordered]@{
            DxUiRoot=$source; DxUiConsumerOutputRoot=$output; DxUiConsumerLockFile=(Join-Path $RepoRoot 'Dependencies/DxUi.lock.json');
            DxUiDisableStlAnnotations='true'; DxUiConsumerToolset=$identity.Identity.toolset;
            DxUiConsumerVCToolsVersion=$identity.Identity.vcToolsVersion; DxUiConsumerSdkVersion=$identity.Identity.windowsSdkVersion;
            DxUiConsumerPreferredToolArchitecture=$identity.Identity.preferredToolArchitecture
        }
        $properties = foreach ($name in $values.Keys) { '<' + $name + '>' + [Security.SecurityElement]::Escape($values[$name]) + '</' + $name + '>' }
        $props = '<Project><PropertyGroup>' + ($properties -join '') + '</PropertyGroup></Project>'
        foreach ($record in @(@{path="DxUi.resolved.$Platform.props";text=$props},
            @{path="DxUi.identity.$Platform.json";text=($identity | ConvertTo-Json -Depth 5)})) {
            $path = Join-Path $dependency $record.path
            if (-not (Test-Path -LiteralPath $path) -or [IO.File]::ReadAllText($path) -cne $record.text) {
                [IO.File]::WriteAllText($path,$record.text,[Text.UTF8Encoding]::new($false))
            }
        }
        if ($CheckUpdates) {
            Import-Module (Join-Path $source 'Tools/ConsumerUpdate.psm1') -Scope Local
            # Consume the pinned helper's information record and re-render only its display.
            # This preserves the pin's network/authentication and update-selection behavior.
            foreach ($record in @(Show-DxUiUpdateNotice -LockFile $values.DxUiConsumerLockFile 6>&1)) {
                $notice = if ($record -is [Management.Automation.InformationRecord]) { [string]$record.MessageData } else { [string]$record }
                $presentation = Get-RSDxUiUpdateNoticePresentation -Notice $notice
                if ($presentation) {
                    Write-Host $presentation.Notice -ForegroundColor $presentation.ForegroundColor
                }
            }
        }
        Write-Host "DxUi restored at $($sourceIdentity.Pin.commit) for $Platform; output identity $($identity.Fingerprint)."
    } finally { Exit-RSArtifactOperationLock -Lock $lease }
}

function Write-RSDxUiModuleProvenance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [Parameter(Mandatory)][string]$ArtifactManifestDir,
        [Parameter(Mandatory)][ValidateSet('x64','ARM64')][string]$Platform,
        [Parameter(Mandatory)][ValidateSet('Debug','Release','ASan Debug')][string]$Configuration
    )
    $source = Get-RSDxUiSourceIdentity -RepoRoot $RepoRoot
    if ($null -eq $source) { return }
    $dependency = Join-Path $RepoRoot '.build/dependencies/DxUi'
    $identity = Get-Content -LiteralPath (Join-Path $dependency "DxUi.identity.$Platform.json") -Raw | ConvertFrom-Json
    if ($identity.Identity.commit -cne $source.Pin.commit -or $identity.Identity.platform -cne $Platform -or
        -not $identity.Identity.disableStlAnnotations) { throw 'The resolved DxUi build identity is stale or uses incompatible STL annotations.' }
    Assert-RSDxUiApiRevision -Pin $source.Pin -Recorded (Get-RSDxUiRecordedApiRevision $identity.Identity) -Context 'The resolved DxUi build identity'
    $archive = Join-Path $dependency "$(Get-RSDxUiOutputKey -Fingerprint $identity.Fingerprint)/$Platform/$Configuration/DxUi.lib"
    foreach ($manifest in Get-ChildItem -LiteralPath $ArtifactManifestDir -Filter '*.manifest' -File) {
        $lines = Get-Content -LiteralPath $manifest.FullName
        $linkLine = @($lines | Where-Object { $_.StartsWith('dxui_link_tlog=') })
        if ($linkLine.Count -eq 0) { continue }
        $targetLine = @($lines | Where-Object { $_.StartsWith('target=') })
        if ($linkLine.Count -ne 1 -or $targetLine.Count -ne 1) { throw 'Ambiguous DxUi build artifact manifest.' }
        $target = Resolve-RSGuardedRepositoryPath -RepoRoot $RepoRoot -GuardedRelativeRoot ".build/$Platform/$Configuration" -Path $targetLine[0].Substring(7)
        $tlog = $linkLine[0].Substring(15)
        $link = [IO.File]::ReadAllText($tlog)
        if ($link.IndexOf($archive.Replace('/','\'),[StringComparison]::OrdinalIgnoreCase) -lt 0) { throw "The module did not link the current pinned DxUi archive: $target" }
        if ($link -match '(?i)Common[\\/]DxUi[\\/]' -or $link -match '(?i)[" ]DxUi\.lib[" ]') {
            throw "The module contains a legacy or ambiguous DxUi linker input: $target"
        }
        $record = [ordered]@{
            repository=$source.Pin.repository; commit=$source.Pin.commit; apiRevision=$source.Pin.apiRevision;
            platform=$Platform; configuration=$Configuration; publicHeadersGitTree=$source.PublicHeadersGitTree;
            buildIdentity=$identity; archiveSha256=(Get-FileHash -LiteralPath $archive).Hash;
            module=[IO.Path]::GetFileName($target); moduleSha256=(Get-FileHash -LiteralPath $target).Hash;
            linkCommandSha256=(Get-FileHash -LiteralPath $tlog).Hash
        }
        $path = $target + '.DxUi.json'
        $json = $record | ConvertTo-Json -Depth 7
        if (-not (Test-Path -LiteralPath $path) -or [IO.File]::ReadAllText($path) -cne $json) { [IO.File]::WriteAllText($path,$json,[Text.UTF8Encoding]::new($false)) }
    }
}

function Copy-RSDxUiPackageProvenance {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RepoRoot, [Parameter(Mandatory)][string]$BuildOutputDir,
        [Parameter(Mandatory)][string]$PackageRoot)
    $pin = Get-RSDxUiPin -RepoRoot $RepoRoot
    if ($null -eq $pin) { return }
    # The packaging caller owns its guarded staging directory and the artifact lease.
    # Enumerate attested source sidecars, then copy only those whose module is shipped.
    foreach ($sidecar in Get-ChildItem -LiteralPath $BuildOutputDir -Recurse -File -Filter '*.DxUi.json') {
        $rootPath = [IO.Path]::GetFullPath($BuildOutputDir).TrimEnd('\','/')
        $relative = $sidecar.FullName.Substring($rootPath.Length + 1)
        $moduleRelative = $relative.Substring(0,$relative.Length - '.DxUi.json'.Length)
        $modulePath = Join-Path $PackageRoot $moduleRelative
        if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) { continue }
        $record = Get-Content -LiteralPath $sidecar.FullName -Raw | ConvertFrom-Json
        if ($record.commit -cne $pin.commit -or $record.repository -cne $pin.repository -or
            $record.module -cne [IO.Path]::GetFileName($moduleRelative) -or
            $record.moduleSha256 -ine (Get-FileHash -LiteralPath $modulePath -Algorithm SHA256).Hash) {
            throw "Shipped DxUi module does not match the exact pinned build provenance: $moduleRelative"
        }
        Copy-Item -LiteralPath $sidecar.FullName -Destination (Join-Path $PackageRoot $relative) -Force
    }
}

Export-ModuleMember -Function Get-RSDxUiPin, Get-RSDxUiSourceIdentity, Restore-RSDxUiDependency, Write-RSDxUiModuleProvenance, Copy-RSDxUiPackageProvenance
