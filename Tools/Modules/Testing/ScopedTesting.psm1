# Repository-owned scope planning. No consumer checkout, service or desktop dependency.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-ScopedGit {
    param([string] $Root, [string[]] $Arguments)
    $start = [Diagnostics.ProcessStartInfo]::new('git')
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($argument in @('-c','core.quotepath=false','-C',$Root) + $Arguments) { $start.ArgumentList.Add($argument) }
    $process = [Diagnostics.Process]::Start($start)
    try {
        $errors = $process.StandardError.ReadToEndAsync()
        $output = $process.StandardOutput.ReadToEnd()
        $process.WaitForExit()
        $errorText = $errors.GetAwaiter().GetResult()
        if ($process.ExitCode) { throw "Git scope discovery failed: $errorText" }
        return $output
    } finally { $process.Dispose() }
}

function Get-ScopedTrackedPaths {
    param([string] $Root)
    return @((Invoke-ScopedGit $Root @('ls-files','--cached','--others','--exclude-standard','-z')) -split '\x00' | Where-Object { $_ } | Sort-Object -Unique)
}

function Get-ScopedChangedPaths {
    param([string] $Root, [string] $BaseRef)
    $base = (Invoke-ScopedGit $Root @('merge-base',$BaseRef,'HEAD')).Trim()
    $committed = Invoke-ScopedGit $Root @('diff','--name-only','--no-renames','-z',$base,'HEAD','--')
    $staged = Invoke-ScopedGit $Root @('diff','--cached','--name-only','--no-renames','-z','--')
    $working = Invoke-ScopedGit $Root @('diff','--name-only','--no-renames','-z','--')
    $untracked = Invoke-ScopedGit $Root @('ls-files','--others','--exclude-standard','-z')
    return @(($committed + $staged + $working + $untracked) -split '\x00' | Where-Object { $_ } | Sort-Object -Unique)
}

function Test-ScopedPattern {
    param([string] $Path, [string] $Pattern)
    $expression = [regex]::Escape($Pattern).Replace('\*\*/','(?:.*/)?').Replace('\*\*','.*').Replace('\*','[^/]*').Replace('\?','[^/]')
    return [regex]::IsMatch($Path, '^' + $expression + '$', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Read-ScopedTestManifest {
    param([string] $Root)
    $manifest = Get-Content -LiteralPath (Join-Path $Root 'Tests/test-scopes.json') -Raw | ConvertFrom-Json
    if ($manifest.version -ne 1 -or @($manifest.scopes).Count -eq 0) { throw 'Invalid test scope manifest.' }
    $names = @($manifest.scopes | ForEach-Object { $_.name })
    if (@($names | Sort-Object -Unique).Count -ne $names.Count -or @($names | Where-Object { $_ -notmatch '^[A-Za-z][A-Za-z0-9.-]*$' }).Count) { throw 'Invalid or duplicate test scopes.' }
    foreach ($rule in $manifest.rules) {
        foreach ($name in $rule.scopes) { if ($name -ne '*' -and $name -notin $names) { throw "Rule selects unknown scope '$name'." } }
    }
    return $manifest
}

function Assert-ScopedTestNames {
    param([string] $Root)
    $declared = @(Get-Content -LiteralPath (Join-Path $Root 'Tests/native-test-files.json') -Raw | ConvertFrom-Json)
    if (@($declared | Sort-Object -Unique).Count -ne $declared.Count) { throw 'Duplicate native test inventory paths.' }
    foreach ($path in $declared) {
        if ($path -match '^(Specs|Measurements|legacy|External)/') { throw "Historical/external source cannot enter the active test inventory: $path" }
        if ($path -cnotmatch '(^|/)[^/]+\.Tests\.[^/]+\.(cpp|h)$' -or -not (Test-Path -LiteralPath (Join-Path $Root $path) -PathType Leaf)) { throw "Invalid/missing native test source: $path" }
    }
    $live = @(Get-ScopedTrackedPaths $Root | Where-Object {
        $_ -notmatch '^(Specs|Measurements|legacy|External)/' -and
        $_ -cne 'Tools/TerminalEngine/TerminalEngineGate0ContractTestAdapter.cpp' -and
        $_ -match '\.(cpp|h)$' -and ($_ -match '^Tests/|/SelfTest/|(^|/)[^/]*(Test|Mock|Fake)[^/]*\.(cpp|h)$') -and
        (Test-Path -LiteralPath (Join-Path $Root $_) -PathType Leaf)
    })
    foreach ($path in $live) { if ($path -notin $declared) { throw "Native test source is absent from Tests/native-test-files.json: $path" } }
    return $declared.Count
}

function Get-ScopedTestPlan {
    param([object] $Manifest, [AllowEmptyCollection()][string[]] $ChangedPaths, [string[]] $Scopes = @(), [switch] $Full)
    $names = @($Manifest.scopes | ForEach-Object { $_.name })
    $selected = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $reasons = [Collections.Generic.List[object]]::new()
    foreach ($scope in $Scopes) {
        if ($scope -notin $names) { throw "Unknown test scope '$scope'. Available: $($names -join ', ')" }
        [void]$selected.Add($scope)
        $reasons.Add([pscustomobject]@{path='(explicit)'; scopes=@($scope); reason='explicit scope'})
    }
    if ($Full) { foreach ($name in $names) { [void]$selected.Add($name) } }
    if (-not $Scopes.Count -and -not $Full) {
        foreach ($path in $ChangedPaths) {
            if ($path -match '(^/|(^|/)\.\.(/|$)|:|\\)') { throw "Expected a repository-relative Git path: $path" }
            if ($path -match '\.md$') {
                $targets = @($Manifest.scopes | Where-Object { -not $_.native } | ForEach-Object name)
                foreach ($name in $targets) { [void]$selected.Add($name) }
                $reasons.Add([pscustomobject]@{path=$path; scopes=$targets; reason='documentation/skill validation; no native test input'})
                continue
            }
            $matched = @($Manifest.rules | Where-Object { Test-ScopedPattern $path $_.pattern })
            if ($matched.Count) {
                foreach ($rule in $matched) {
                    $targets = if ('*' -in $rule.scopes) { $names } else { @($rule.scopes) }
                    foreach ($name in $targets) { [void]$selected.Add($name) }
                    $reasons.Add([pscustomobject]@{path=$path; scopes=$targets; reason=$rule.reason})
                }
            } elseif ($path -match '^(docs|Measurements|Changes)/') {
                $reasons.Add([pscustomobject]@{path=$path; scopes=@(); reason='documentation; no native test input'})
            } else {
                foreach ($name in $names) { [void]$selected.Add($name) }
                $reasons.Add([pscustomobject]@{path=$path; scopes=$names; reason='unmapped input; conservative full fallback'})
            }
        }
    }
    return [pscustomobject]@{scopes=@($names | Where-Object { $selected.Contains($_) }); reasons=@($reasons); full=($selected.Count -eq $names.Count)}
}

function Get-ScopedDigest {
    param([string] $Text)
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-','').ToLowerInvariant() }
    finally { $hash.Dispose() }
}

function Get-ScopedPrCandidateScopes {
    param([string] $Root, [object] $Manifest, [string] $Platform, [string] $Configuration,
        [AllowEmptyCollection()][string[]] $ChangedPaths)
    $profile = @($Manifest.prCoverage | Where-Object { $_.platform -eq $Platform -and $_.configuration -eq $Configuration })
    $covered = @($profile | ForEach-Object { $_.scopes } | Sort-Object -Unique)
    if ($Manifest.PSObject.Properties['prNativeScopeModule']) {
        Import-Module (Join-Path $Root $Manifest.prNativeScopeModule) -Force
        if (-not (Get-NativeScope -ChangedPaths $ChangedPaths).Native) {
            $covered = @($Manifest.scopes | Where-Object { -not $_.native -and $_.name -in $covered } | ForEach-Object name)
        }
    }
    return $covered
}

function Get-ScopedPrWorkflowDigest {
    <#
    .SYNOPSIS
    Binds reviewed PR coverage to its complete repository-owned execution surface.
    .DESCRIPTION
    Hashes workflows/local actions, tooling, build configuration, dependency pins and
    executable scripts anywhere in the checkout. This conservative superset includes
    recursively imported and dynamically selected tooling without guessing shell/YAML
    control flow. Missing local actions/workflows and reparse inputs fail closed.
    Product C++ implementation changes do not alter the gate's execution contract.
    .PARAMETER Root
    Repository root. Reads Git and source files; writes nothing and performs no fetch.
    .OUTPUTS
    Lowercase SHA-256 of canonical path/content digests. Excludes only the scope
    manifest's own prWorkflowDigest property to avoid a self-referential digest.
    #>
    param([Parameter(Mandatory=$true)][string] $Root)
    $allPaths = @(Get-ScopedTrackedPaths $Root)
    $actionRoots = [Collections.Generic.List[string]]::new()
    foreach ($yaml in @($allPaths | Where-Object { $_ -match '\.(yml|yaml)$' })) {
        $yamlPath = Join-Path $Root $yaml
        if (-not (Test-Path -LiteralPath $yamlPath -PathType Leaf)) { throw "Missing PR YAML input: $yaml" }
        foreach ($reference in [regex]::Matches([IO.File]::ReadAllText($yamlPath), '(?m)^\s*(?:-\s*)?uses:\s*["'']?(\./[^\s"''#]+)')) {
            $target = $reference.Groups[1].Value.Substring(2).Replace('\','/')
            if ($target -match '(^|/)\.\.(/|$)' -or $target.Contains('${{') -or [IO.Path]::IsPathRooted($target)) { throw "Unresolved local PR action/workflow: $target" }
            if (Test-Path -LiteralPath (Join-Path $Root $target) -PathType Container) { $actionRoots.Add($target.TrimEnd('/') + '/') }
        }
    }
    # Local actions may consume arbitrary assets beside their manifest, including Dockerfiles.
    $paths = @($allPaths | Where-Object {
        $candidatePath = $_
        @($actionRoots | Where-Object { $candidatePath.StartsWith($_, [StringComparison]::Ordinal) }).Count -gt 0 -or
        $_ -match '^(\.github/(workflows|actions)/|Tools/|Build/|Dependencies/)' -or
        $_ -match '\.(?:ps1|psm1|psd1|py|js|mjs|cjs|ts|sh|cmd|bat|yml|yaml|props|targets|vcxproj|proj|sln|slnx|vcxitems|cmake)$' -or
        $_ -match '^Tests/[^/]+\.json$' -or
        $_ -match '^Specs/(Testing|Build)/.*\.json$' -or
        $_ -match '^vcpkg(?:-tool|-configuration)?\.json$'
    } | Sort-Object -Unique)
    if ('.github/workflows/ci.yml' -notin $paths) { throw 'PR workflow is missing.' }
    $rows = [Collections.Generic.List[string]]::new()
    foreach ($relative in $paths) {
        $path = Join-Path $Root $relative
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "PR execution input is missing: $relative" }
        $component = Get-Item -LiteralPath $path -Force
        while ($component.FullName -ne [IO.Path]::GetFullPath($Root).TrimEnd('\','/')) {
            if (($component.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Reparse PR execution input: $relative" }
            $component = if ($component -is [IO.DirectoryInfo]) { $component.Parent } else { $component.Directory }
            if (-not $component) { throw "PR execution input escaped repository: $relative" }
        }
        $content = [IO.File]::ReadAllText($path) -replace "`r`n","`n"
        if ($relative -eq 'Tests/test-scopes.json') {
            $scope = $content | ConvertFrom-Json
            $scope.PSObject.Properties.Remove('prWorkflowDigest')
            $content = $scope | ConvertTo-Json -Depth 32 -Compress
        }
        if ($relative -match '\.(yml|yaml)$') {
            foreach ($reference in [regex]::Matches($content, '(?m)^\s*(?:-\s*)?uses:\s*["'']?(\./[^\s"''#]+)')) {
                $target = $reference.Groups[1].Value.Substring(2).Replace('\','/')
                if ($target -match '(^|/)\.\.(/|$)' -or $target.Contains('${{')) { throw "Unresolved local PR action/workflow: $target" }
                $localPath = Join-Path $Root $target
                if (Test-Path -LiteralPath $localPath -PathType Container) {
                    if (-not (Test-Path -LiteralPath (Join-Path $localPath 'action.yml') -PathType Leaf) -and
                        -not (Test-Path -LiteralPath (Join-Path $localPath 'action.yaml') -PathType Leaf)) { throw "Local action has no manifest: $target" }
                } elseif (-not (Test-Path -LiteralPath $localPath -PathType Leaf)) { throw "Local PR action/workflow is missing: $target" }
            }
        }
        $digest = if ($relative -match '\.(?:ps1|psm1|psd1|py|js|mjs|cjs|ts|sh|cmd|bat|yml|yaml|json|props|targets|vcxproj|proj|sln|slnx|vcxitems|cmake|patch|md|txt)$') { Get-ScopedDigest $content } else { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() }
        $rows.Add($relative + ':' + $digest)
    }
    return Get-ScopedDigest ($rows -join "`n")
}

function Get-ScopedPrCoverage {
    param([string] $Root, [object] $Manifest, [string] $Platform, [string] $Configuration)
    $profile = @($Manifest.prCoverage | Where-Object { $_.platform -eq $Platform -and $_.configuration -eq $Configuration })
    if (-not $profile.Count) { return @() }
    # The forthcoming PR executes its candidate workflow. Its reviewed digest must match; API problems keep work local.
    try {
        # GitHub receives committed bytes. Dirty or concurrently changing work cannot be delegated.
        $candidate = (Invoke-ScopedGit $Root @('rev-parse','HEAD')).Trim()
        if (Invoke-ScopedGit $Root @('status','--porcelain','--untracked-files=normal')) { return @() }
        $local = [IO.File]::ReadAllText((Join-Path $Root '.github/workflows/ci.yml')) -replace "`r`n","`n"
        if ((Get-ScopedPrWorkflowDigest $Root) -cne $Manifest.prWorkflowDigest -or $local -notmatch '(?m)^  pull_request:') { return @() }
        $workflow = & gh api "repos/$($Manifest.repository)/actions/workflows/ci.yml" 2>$null | ConvertFrom-Json
        if ($LASTEXITCODE -ne 0 -or $workflow.state -ne 'active') { return @() }
        $paths = @(Get-ScopedChangedPaths $Root ('origin/' + $Manifest.defaultBranch))
        $covered = @(Get-ScopedPrCandidateScopes $Root $Manifest $Platform $Configuration $paths)
        if ((Invoke-ScopedGit $Root @('rev-parse','HEAD')).Trim() -cne $candidate -or
            (Invoke-ScopedGit $Root @('status','--porcelain','--untracked-files=normal'))) { return @() }
        return $covered
    } catch [System.Management.Automation.RuntimeException] { return @() }
    catch [ArgumentException] { return @() }
}

function Test-RSScheduledTestsRequired {
    param([string]$EventName,[string]$Commit,[AllowNull()][object]$PreviousRun)
    if ($EventName -ne 'schedule' -or -not $PreviousRun -or -not $Commit) {return $true}
    if (-not $PreviousRun.PSObject.Properties['head_sha'] -or -not $PreviousRun.PSObject.Properties['conclusion']) {return $true}
    # Use the latest completed attempt, never an older success hidden behind a later failure or cancellation.
    return $PreviousRun.conclusion -ne 'success' -or $PreviousRun.head_sha -cne $Commit
}

Export-ModuleMember -Function Get-ScopedChangedPaths, Read-ScopedTestManifest, Assert-ScopedTestNames, Get-ScopedTestPlan,
    Get-ScopedDigest,
    Get-ScopedPrWorkflowDigest, Get-ScopedPrCoverage, Get-ScopedPrCandidateScopes,
    Test-RSScheduledTestsRequired
