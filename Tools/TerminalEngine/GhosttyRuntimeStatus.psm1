Set-StrictMode -Version Latest

$lifecycleModule = Join-Path $PSScriptRoot 'GhosttyRuntimeLifecycle.psm1'
$canonicalJsonModule = Join-Path (Split-Path -Parent $PSScriptRoot) 'Modules\Testing\ValidationFingerprint.psm1'
Import-Module $lifecycleModule -ErrorAction Stop
Import-Module $canonicalJsonModule -ErrorAction Stop

function Get-RSGhosttyStatusUtcText {
    param([Parameter(Mandatory = $true)][DateTime]$Value)

    return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-RSGhosttyStatusRepositoryCommit {
    param([Parameter(Mandatory = $true)][string]$RepoRoot)

    $output = @(& git.exe -C $RepoRoot rev-parse HEAD 2>$null)
    $exitCode = $LASTEXITCODE
    $value = $output | Select-Object -First 1
    if ($exitCode -ne 0 -or [string]$value -cnotmatch '^[0-9a-f]{40}$') {
        throw 'Unable to bind Ghostty status evidence to the current repository commit.'
    }
    return [string]$value
}

function Get-RSGhosttyStatusVersionFromBuildZon {
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }
    $match = [regex]::Match($Text, '(?m)^\s*\.version\s*=\s*"(?<version>[^"]+)"')
    if (-not $match.Success) {
        return $null
    }
    return $match.Groups['version'].Value
}

function Invoke-RSGhosttyStatusWebRequest {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('github', 'osv')][string]$Source,
        [Parameter(Mandatory = $true)][string]$Uri,
        [ValidateSet('GET', 'POST')][string]$Method = 'GET',
        [AllowNull()][string]$Body,
        [Parameter(Mandatory = $true)][DateTime]$AttemptedUtc
    )

    try {
        $parameters = @{
            Uri = $Uri
            Method = $Method
            Headers = @{ 'User-Agent' = 'RedSalamander-Ghostty-Status/1' }
            TimeoutSec = 30
            ErrorAction = 'Stop'
        }
        if ($Method -ceq 'POST') {
            $parameters.ContentType = 'application/json'
            $parameters.Body = $Body
        }
        $response = Invoke-WebRequest @parameters
        $value = $null
        try {
            $value = [string]$response.Content | ConvertFrom-Json -Depth 100 -NoEnumerate -DateKind String -ErrorAction Stop
        }
        catch [System.ArgumentException] {
            return [pscustomobject]@{
                Attempt = [ordered]@{ source = $Source; outcome = 'malformed'; statusCode = [int]$response.StatusCode; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
                Value = $null
            }
        }
        catch [System.Text.Json.JsonException] {
            return [pscustomobject]@{
                Attempt = [ordered]@{ source = $Source; outcome = 'malformed'; statusCode = [int]$response.StatusCode; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
                Value = $null
            }
        }
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = $Source; outcome = 'success'; statusCode = [int]$response.StatusCode; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = $value
        }
    }
    catch [System.TimeoutException] {
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = $Source; outcome = 'timeout'; statusCode = $null; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = $null
        }
    }
    catch [System.Threading.Tasks.TaskCanceledException] {
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = $Source; outcome = 'timeout'; statusCode = $null; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = $null
        }
    }
    catch [System.OperationCanceledException] {
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = $Source; outcome = 'timeout'; statusCode = $null; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = $null
        }
    }
    catch [Microsoft.PowerShell.Commands.HttpResponseException] {
        $statusCode = if ($null -ne $_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { $null }
        $outcome = if ($statusCode -eq 403 -or $statusCode -eq 429) { 'rate-limit' } else { 'refused' }
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = $Source; outcome = $outcome; statusCode = $statusCode; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = $null
        }
    }
    catch [System.Net.Http.HttpRequestException] {
        $statusCode = if ($null -ne $_.Exception.StatusCode) { [int]$_.Exception.StatusCode } else { $null }
        $outcome = if ($statusCode -eq 403 -or $statusCode -eq 429) { 'rate-limit' }
            elseif ($statusCode -eq 404) { 'refused' }
            else { 'refused' }
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = $Source; outcome = $outcome; statusCode = $statusCode; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = $null
        }
    }
    catch [System.Net.WebException] {
        $outcome = if ($_.Exception.Status -eq [System.Net.WebExceptionStatus]::Timeout) { 'timeout' } else { 'refused' }
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = $Source; outcome = $outcome; statusCode = $null; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = $null
        }
    }
}

function Read-RSGhosttyStatusFixture {
    param(
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter(Mandatory = $true)][DateTime]$AttemptedUtc
    )

    try {
        $value = Get-Content -LiteralPath $LiteralPath -Raw -ErrorAction Stop |
            ConvertFrom-Json -Depth 100 -NoEnumerate -DateKind String -ErrorAction Stop
        $outcome = [string]$value.outcome
        if ($outcome -notin @('success', 'malformed', 'rate-limit', 'timeout', 'refused')) {
            throw [System.ArgumentException]::new('Fixture outcome is invalid.')
        }
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = 'fixture'; outcome = $outcome; statusCode = $null; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = if ($outcome -ceq 'success') { $value } else { $null }
        }
    }
    catch [System.ArgumentException] {
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = 'fixture'; outcome = 'malformed'; statusCode = $null; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = $null
        }
    }
    catch [System.Text.Json.JsonException] {
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = 'fixture'; outcome = 'malformed'; statusCode = $null; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = $null
        }
    }
    catch [System.Management.Automation.RuntimeException] {
        return [pscustomobject]@{
            Attempt = [ordered]@{ source = 'fixture'; outcome = 'malformed'; statusCode = $null; attemptedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc }
            Value = $null
        }
    }
}

function Get-RSGhosttyStatusRealObservation {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Online', 'Candidate')][string]$Mode,
        [Parameter(Mandatory = $true)][string]$LockedCommit,
        [AllowNull()][string]$CandidateCommit,
        [Parameter(Mandatory = $true)][DateTime]$AttemptedUtc
    )

    $attempts = [Collections.Generic.List[object]]::new()
    $apiRoot = 'https://api.github.com/repos/ghostty-org/ghostty'
    $targetCommit = $CandidateCommit
    if ($Mode -ceq 'Online') {
        $repository = Invoke-RSGhosttyStatusWebRequest -Source github -Uri $apiRoot -AttemptedUtc $AttemptedUtc
        $attempts.Add($repository.Attempt)
        if ($null -eq $repository.Value -or [string]::IsNullOrWhiteSpace([string]$repository.Value.default_branch)) {
            return [pscustomobject]@{ Attempts = @($attempts); Value = $null }
        }
        $targetCommit = [string]$repository.Value.default_branch
    }

    $commitResponse = Invoke-RSGhosttyStatusWebRequest -Source github -Uri "$apiRoot/commits/$targetCommit" -AttemptedUtc $AttemptedUtc
    $attempts.Add($commitResponse.Attempt)
    if ($null -eq $commitResponse.Value) {
        return [pscustomobject]@{ Attempts = @($attempts); Value = $null }
    }
    $resolvedCommit = [string]$commitResponse.Value.sha
    $resolvedTree = [string]$commitResponse.Value.commit.tree.sha
    if ($resolvedCommit -cnotmatch '^[0-9a-f]{40}$' -or $resolvedTree -cnotmatch '^[0-9a-f]{40}$') {
        $attempts[$attempts.Count - 1].outcome = 'malformed'
        return [pscustomobject]@{ Attempts = @($attempts); Value = $null }
    }
    if ($Mode -ceq 'Candidate' -and $resolvedCommit -cne $CandidateCommit) {
        return [pscustomobject]@{
            Attempts = @($attempts)
            Value = [pscustomobject]@{ substitutionRefused = $true; commit = $null; tree = $null; version = $null; aheadCount = $null; behindCount = $null; advisories = @() }
        }
    }

    $contents = Invoke-RSGhosttyStatusWebRequest -Source github -Uri "$apiRoot/contents/build.zig.zon?ref=$resolvedCommit" -AttemptedUtc $AttemptedUtc
    $attempts.Add($contents.Attempt)
    $version = $null
    if ($null -ne $contents.Value -and [string]$contents.Value.encoding -ceq 'base64') {
        try {
            $zonText = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(([string]$contents.Value.content).Replace("`n", '')))
            $version = Get-RSGhosttyStatusVersionFromBuildZon -Text $zonText
        }
        catch [FormatException] {
            $attempts[$attempts.Count - 1].outcome = 'malformed'
        }
    }

    $compare = Invoke-RSGhosttyStatusWebRequest -Source github -Uri "$apiRoot/compare/$LockedCommit...$resolvedCommit" -AttemptedUtc $AttemptedUtc
    $attempts.Add($compare.Attempt)
    $aheadCount = $null
    $behindCount = $null
    $comparisonStatus = $null
    if ($null -ne $compare.Value) {
        $aheadCount = [int]$compare.Value.ahead_by
        $behindCount = [int]$compare.Value.behind_by
        $comparisonStatus = [string]$compare.Value.status
    }

    $advisories = @()
    if ($Mode -ceq 'Online') {
        $osvBody = ConvertTo-RSCanonicalJson -InputObject ([ordered]@{ commit = $LockedCommit })
        $osv = Invoke-RSGhosttyStatusWebRequest -Source osv -Uri 'https://api.osv.dev/v1/query' -Method POST -Body $osvBody -AttemptedUtc $AttemptedUtc
        $attempts.Add($osv.Attempt)
        $vulnerabilities = if ($null -ne $osv.Value -and $null -ne $osv.Value.PSObject.Properties['vulns']) {
            @($osv.Value.vulns)
        }
        else {
            @()
        }
        if ($null -ne $osv.Value) {
            $advisories = @($vulnerabilities | ForEach-Object {
                    [ordered]@{
                        source = 'osv'
                        identifier = [string]$_.id
                        url = if (@($_.references).Count -gt 0 -and [string]$_.references[0].url -match '^https://') { [string]$_.references[0].url } else { "https://osv.dev/vulnerability/$($_.id)" }
                        disposition = 'affected'
                        retrievedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc
                    }
                })
        }
    }

    return [pscustomobject]@{
        Attempts = @($attempts)
        Value = [pscustomobject]@{
            substitutionRefused = $false
            commit = $resolvedCommit
            tree = $resolvedTree
            version = $version
            aheadCount = $aheadCount
            behindCount = $behindCount
            comparisonStatus = $comparisonStatus
            advisories = $advisories
        }
    }
}

function Get-RSGhosttyStatusAdvisoryObservation {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Lock,
        [Parameter(Mandatory = $true)][DateTime]$AttemptedUtc
    )

    $lockedCommit = [string]$Lock.upstream.commit
    $osvBody = ConvertTo-RSCanonicalJson -InputObject ([ordered]@{ commit = $lockedCommit })
    $osv = Invoke-RSGhosttyStatusWebRequest -Source osv -Uri 'https://api.osv.dev/v1/query' -Method POST -Body $osvBody -AttemptedUtc $AttemptedUtc
    $advisories = @()
    $vulnerabilities = if ($null -ne $osv.Value -and $null -ne $osv.Value.PSObject.Properties['vulns']) {
        @($osv.Value.vulns)
    }
    else {
        @()
    }
    if ($null -ne $osv.Value) {
        $advisories = @($vulnerabilities | ForEach-Object {
                [ordered]@{
                    source = 'osv'
                    identifier = [string]$_.id
                    url = if (@($_.references).Count -gt 0 -and [string]$_.references[0].url -match '^https://') { [string]$_.references[0].url } else { "https://osv.dev/vulnerability/$($_.id)" }
                    disposition = 'affected'
                    retrievedUtc = Get-RSGhosttyStatusUtcText $AttemptedUtc
                }
            })
    }
    return [pscustomobject]@{
        Attempts = @($osv.Attempt)
        Value = if ($null -eq $osv.Value) { $null } else {
            [pscustomobject]@{
                substitutionRefused = $false
                commit = $lockedCommit
                tree = [string]$Lock.upstream.tree
                version = [string]$Lock.upstream.version
                aheadCount = 0
                behindCount = 0
                comparisonStatus = 'identical'
                advisories = $advisories
            }
        }
    }
}

function New-RSGhosttyRuntimeStatusReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][ValidateSet('Locked', 'Online', 'Candidate')][string]$Mode,
        [string]$CandidateCommit,
        [string]$NetworkFixturePath,
        [switch]$AdvisoryOnly,
        [DateTime]$GeneratedUtc = [DateTime]::UtcNow,
        [string]$ExpectedSourceCommit
    )

    $repoRoot = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    if ($Mode -ceq 'Candidate' -and $CandidateCommit -cnotmatch '^[0-9a-f]{40}$') {
        throw 'Candidate mode requires -CandidateCommit as one explicit lowercase full 40-character commit.'
    }
    if ($Mode -cne 'Candidate' -and -not [string]::IsNullOrWhiteSpace($CandidateCommit)) {
        throw '-CandidateCommit is valid only in Candidate mode.'
    }
    if ($Mode -ceq 'Locked' -and -not [string]::IsNullOrWhiteSpace($NetworkFixturePath)) {
        throw 'Locked mode performs zero network observation and does not accept a network fixture.'
    }
    if ($AdvisoryOnly -and $Mode -cne 'Online') {
        throw '-AdvisoryOnly is valid only in Online mode.'
    }

    $lockPath = Join-Path $repoRoot 'External\TerminalEngine\GhosttyRuntimeLock.v1.json'
    $schemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeLock.schema.json'
    $lock = Read-RSGhosttyRuntimeLock -LiteralPath $lockPath -SchemaPath $schemaPath -RepoRoot $repoRoot
    $lockSha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $lockPath).Hash.ToLowerInvariant()
    $sourceCommit = Get-RSGhosttyStatusRepositoryCommit -RepoRoot $repoRoot
    if (-not [string]::IsNullOrWhiteSpace($ExpectedSourceCommit) -and $sourceCommit -cne $ExpectedSourceCommit) {
        throw "Current source commit '$sourceCommit' does not match the required release binding '$ExpectedSourceCommit'."
    }

    $generatedText = Get-RSGhosttyStatusUtcText $GeneratedUtc
    $networkAttempts = @()
    $observed = $null
    if ($Mode -ceq 'Locked') {
        $observed = [pscustomobject]@{
            substitutionRefused = $false
            commit = [string]$lock.upstream.commit
            tree = [string]$lock.upstream.tree
            version = [string]$lock.upstream.version
            aheadCount = 0
            behindCount = 0
            comparisonStatus = 'identical'
            advisories = @()
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($NetworkFixturePath)) {
        $fixture = Read-RSGhosttyStatusFixture -LiteralPath $NetworkFixturePath -AttemptedUtc $GeneratedUtc
        $networkAttempts = @($fixture.Attempt)
        if ($null -ne $fixture.Value) {
            if ($AdvisoryOnly) {
                $observed = [pscustomobject]@{
                    substitutionRefused = $false
                    commit = [string]$lock.upstream.commit
                    tree = [string]$lock.upstream.tree
                    version = [string]$lock.upstream.version
                    aheadCount = 0
                    behindCount = 0
                    comparisonStatus = 'identical'
                    advisories = @($fixture.Value.advisories)
                }
            }
            else {
                $selected = if ($Mode -ceq 'Candidate') { $fixture.Value.candidate } else { $fixture.Value.observation }
                if ($null -ne $selected) {
                    $observed = [pscustomobject]@{
                        substitutionRefused = $Mode -ceq 'Candidate' -and [string]$selected.commit -cne $CandidateCommit
                        commit = if ($Mode -ceq 'Candidate' -and [string]$selected.commit -cne $CandidateCommit) { $null } else { [string]$selected.commit }
                        tree = if ($Mode -ceq 'Candidate' -and [string]$selected.commit -cne $CandidateCommit) { $null } else { [string]$selected.tree }
                        version = if ($Mode -ceq 'Candidate' -and [string]$selected.commit -cne $CandidateCommit) { $null } else { [string]$selected.version }
                        aheadCount = if ($null -eq $selected.aheadCount) { $null } else { [int]$selected.aheadCount }
                        behindCount = if ($null -eq $selected.behindCount) { $null } else { [int]$selected.behindCount }
                        comparisonStatus = [string]$selected.comparisonStatus
                        advisories = @($fixture.Value.advisories)
                    }
                }
            }
        }
    }
    elseif ($AdvisoryOnly) {
        $advisoryObservation = Get-RSGhosttyStatusAdvisoryObservation -Lock $lock -AttemptedUtc $GeneratedUtc
        $networkAttempts = @($advisoryObservation.Attempts)
        $observed = $advisoryObservation.Value
    }
    else {
        $real = Get-RSGhosttyStatusRealObservation -Mode $Mode -LockedCommit ([string]$lock.upstream.commit) -CandidateCommit $CandidateCommit -AttemptedUtc $GeneratedUtc
        $networkAttempts = @($real.Attempts)
        $observed = $real.Value
    }

    $reasons = [Collections.Generic.List[string]]::new()
    $result = 'required-manual'
    $nextAction = 'manual-admission'
    $failedOutcomes = @($networkAttempts | Where-Object { [string]$_.outcome -in @('rate-limit', 'timeout', 'refused') })
    $malformedOutcomes = @($networkAttempts | Where-Object { [string]$_.outcome -ceq 'malformed' })
    if ($Mode -ceq 'Locked') {
        $result = 'current'
        $nextAction = 'none'
        $reasons.Add('canonical-lock-validated')
    }
    elseif ($failedOutcomes.Count -gt 0) {
        $result = 'unreachable'
        $nextAction = 'retry-observation'
        foreach ($outcome in @($failedOutcomes | ForEach-Object outcome | Sort-Object -Unique)) {
            $reasons.Add("network-$outcome")
        }
        if ($reasons.Count -eq 0) { $reasons.Add('network-refused') }
    }
    elseif ($malformedOutcomes.Count -gt 0) {
        $result = 'required-manual'
        $nextAction = 'manual-admission'
        $reasons.Add('observation-malformed')
    }
    elseif ($null -eq $observed) {
        $result = 'unreachable'
        $nextAction = 'retry-observation'
        $reasons.Add('network-refused')
    }
    elseif ([bool]$observed.substitutionRefused) {
        $result = 'required-manual'
        $nextAction = 'manual-admission'
        $reasons.Add('candidate-substitution-refused')
    }
    elseif ($null -eq $observed.commit -or $null -eq $observed.tree -or $null -eq $observed.version) {
        $result = 'required-manual'
        $nextAction = 'manual-admission'
        $reasons.Add('observation-malformed')
    }
    elseif ([int]$observed.behindCount -gt 0 -or [string]$observed.comparisonStatus -in @('behind', 'diverged')) {
        $result = 'required-manual'
        $nextAction = 'manual-admission'
        $reasons.Add('unrelated-history')
    }
    elseif (@($observed.advisories | Where-Object { [string]$_.disposition -in @('affected', 'unknown') }).Count -gt 0) {
        $result = 'security-review'
        $nextAction = 'review-security'
        $reasons.Add('advisory-review-required')
    }
    elseif ([string]$observed.commit -ceq [string]$lock.upstream.commit) {
        $result = 'current'
        $nextAction = 'none'
        $reasons.Add('observation-matches-lock')
    }
    elseif ([int]$observed.aheadCount -gt 0 -and [int]$observed.behindCount -eq 0) {
        $result = 'update-observed'
        $nextAction = if ($Mode -ceq 'Candidate') { 'freeze-candidate' } else { 'review-update' }
        $reasons.Add('descendant-update-observed')
    }
    else {
        $result = 'required-manual'
        $nextAction = 'manual-admission'
        $reasons.Add('ancestry-unresolved')
    }

    $observedAdvisories = if ($null -eq $observed) { @() } else { @($observed.advisories) }
    $advisories = @($observedAdvisories | ForEach-Object {
            [ordered]@{
                source = 'osv'
                identifier = [string]$_.identifier
                url = [string]$_.url
                disposition = [string]$_.disposition
                retrievedUtc = if ([string]::IsNullOrWhiteSpace([string]$_.retrievedUtc)) { $generatedText } else { [string]$_.retrievedUtc }
            }
        } | Sort-Object { $_.identifier })
    $reasonCodes = [string[]]@($reasons | Sort-Object -Unique -CaseSensitive)
    return [ordered]@{
        schemaVersion = 1
        generatedUtc = $generatedText
        mode = $Mode
        result = $result
        reasonCodes = $reasonCodes
        sourceBinding = [ordered]@{ repositoryCommit = $sourceCommit }
        canonicalLock = [ordered]@{
            path = 'External/TerminalEngine/GhosttyRuntimeLock.v1.json'
            sha256 = $lockSha256
            commit = [string]$lock.upstream.commit
            tree = [string]$lock.upstream.tree
            version = [string]$lock.upstream.version
            zigVersion = [string]$lock.toolchain.zigVersion
        }
        observation = [ordered]@{
            source = if ($Mode -ceq 'Locked' -or $AdvisoryOnly) { 'canonical-lock' } elseif ($Mode -ceq 'Online') { 'github-default-branch' } else { 'github-explicit-commit' }
            commit = if ($null -eq $observed) { $null } else { $observed.commit }
            tree = if ($null -eq $observed) { $null } else { $observed.tree }
            version = if ($null -eq $observed) { $null } else { $observed.version }
            aheadCount = if ($null -eq $observed) { $null } else { $observed.aheadCount }
            behindCount = if ($null -eq $observed) { $null } else { $observed.behindCount }
            observedUtc = $generatedText
            candidateRequestedCommit = if ($Mode -ceq 'Candidate') { $CandidateCommit } else { $null }
        }
        advisories = $advisories
        networkAttempts = @($networkAttempts)
        nextAction = $nextAction
    }
}

function Write-RSGhosttyRuntimeStatusReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][Collections.IDictionary]$Report,
        [string]$OutputPath
    )

    $repoRoot = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $buildRoot = Join-Path $repoRoot '.build'
    [void](New-Item -ItemType Directory -Path $buildRoot -Force)
    $generatedStamp = ([DateTime]::Parse([string]$Report.generatedUtc, [Globalization.CultureInfo]::InvariantCulture)).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $resolvedPath = if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        Join-Path $buildRoot "TerminalEngineStatus\$generatedStamp\status.json"
    }
    elseif ([IO.Path]::IsPathFullyQualified($OutputPath)) {
        [IO.Path]::GetFullPath($OutputPath)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $repoRoot $OutputPath))
    }
    $buildPrefix = $buildRoot.TrimEnd('\') + '\'
    if (-not $resolvedPath.StartsWith($buildPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Ghostty status output must remain below '$buildRoot': $resolvedPath"
    }
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $resolvedPath) -Force)
    $statusSchema = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeStatus.schema.json'
    Write-RSAtomicCanonicalJsonFile -LiteralPath $resolvedPath -InputObject $Report -SchemaPath $statusSchema -AllowedRoot $buildRoot
    return $resolvedPath
}

Export-ModuleMember -Function @(
    'New-RSGhosttyRuntimeStatusReport',
    'Write-RSGhosttyRuntimeStatusReport')
