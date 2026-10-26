<#
.SYNOPSIS
Runs the shared Changes behavior fixtures without a native build or external test framework.
.DESCRIPTION
Creates uniquely owned fixture directories beneath an existing caller-owned scratch root.
Never reads or changes the application's changelog. Removes only its own direct fixture children.
.PARAMETER ScratchRoot
Existing test scratch parent supplied by the repository's test harness.
.OUTPUTS
Console case results; throws if any assertion fails. Fixture files are removed on completion.
.NOTES
Requires PowerShell 7. No network, native code, application processes or desktop input.
.EXAMPLE
./Tools/Test-ChangelogWorkflow.ps1 -ScratchRoot <existing-test-scratch-directory>
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string] $ScratchRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $repository 'Tools/Modules/Documentation/Changelog.psm1')
$fold = Join-Path $repository 'Tools/Fold-Changelog.ps1'
$validate = Join-Path $repository 'Tools/Test-Changelog.ps1'
$parent = [IO.Path]::GetFullPath($ScratchRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (-not [IO.Directory]::Exists($parent)) { throw "Scratch root must already exist: $parent" }
# A repository adapter supplies its existing scratch root. Honor the stricter marked-sandbox policy where installed.
$sandboxModule = Join-Path $repository 'Tools/Modules/Testing/TestSandbox.psm1'
if ([IO.File]::Exists($sandboxModule)) {
    Import-Module $sandboxModule
    $testRoot = Get-RSTestSandboxRoot -RepoRoot $repository
    [void](Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path $parent -RequireDescendant)
}
$results = [Collections.Generic.List[string]]::new()
$failures = [Collections.Generic.List[string]]::new()

function Assert([bool] $Condition, [string] $Message) { if (-not $Condition) { throw $Message } }
function Refuses([scriptblock] $Body) {
    $refused = $false
    try { & $Body | Out-Null } catch { $refused = $true }
    Assert $refused 'Expected the operation to refuse'
}
function Put([string] $Root, [string] $Name, [string] $Text) {
    $path = Join-Path $Root $Name
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
    [IO.File]::WriteAllText($path, $Text, [Text.UTF8Encoding]::new($false))
}
function Snapshot([string] $Root) {
    return (@(Get-ChildItem -LiteralPath $Root -Recurse -File | Where-Object Name -ne '.changelog-fold.lock' | Sort-Object FullName | ForEach-Object {
        [IO.Path]::GetRelativePath($Root, $_.FullName) + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
    }) -join "`n")
}
function Invoke-FixtureProcess([Diagnostics.ProcessStartInfo] $Start,
    [ValidateRange(1,30000)][int] $TimeoutMilliseconds = 30000, [ref] $ProcessIdentity) {
    # This standalone cross-project suite owns this synthetic child; repository-specific native runners are not dependencies.
    $process = [Diagnostics.Process]::Start($Start)
    try {
        if ($ProcessIdentity) {
            $ProcessIdentity.Value = @{ Id = $process.Id; StartTicks = $process.StartTime.ToUniversalTime().Ticks }
        }
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        Assert ($process.WaitForExit($TimeoutMilliseconds)) 'Fixture child did not exit within its budget'
        $drain = [Threading.Tasks.Task]::WhenAll([Threading.Tasks.Task[]]@($stdout, $stderr))
        Assert ($drain.Wait(5000)) 'Fixture child pipes did not close within their drain budget'
        return [pscustomobject]@{ ExitCode = $process.ExitCode; Output = $stdout.GetAwaiter().GetResult() + $stderr.GetAwaiter().GetResult() }
    } finally {
        try {
            if (-not $process.HasExited) {
                $process.Kill($true)
                Assert ($process.WaitForExit(5000)) 'Owned fixture child did not terminate within its cleanup budget'
            }
        } finally { $process.Dispose() }
    }
}
function Case([string] $Name, [scriptblock] $Body) {
    $root = Join-Path $parent ('changelog-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($root)
    try {
        Put $root 'Changes/README.md' "# Changelog fragments`n"
        Put $root 'CHANGELOG.md' "# Changelog`n`n## Unreleased`n`n- Existing.`n`n## 1.0`n`n- Released.`n"
        & $Body $root
        $results.Add($Name)
        Write-Host "PASS changelog: $Name"
    } catch { $failures.Add("${Name}: $($_.Exception.Message)"); Write-Host "FAIL changelog: $($failures[-1])" }
    finally {
        $resolved = [IO.Path]::GetFullPath($root)
        if ([IO.Path]::GetDirectoryName($resolved) -cne $parent -or [IO.Path]::GetFileName($resolved) -notmatch '^changelog-[0-9a-f]{32}$') {
            throw "Refusing fixture cleanup outside its exact scratch parent: $resolved"
        }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

Case 'editing permits one valid dated Unicode change' {
    param($r)
    Put $r 'Changes/2026-10-21-unicode.md' "- Preserve café and 日本語.`n  Validation remains explicit.`n  - Pending native gates.`n"
    & $validate -Root $r
    Assert ((Get-ChangelogProblems $r).Count -eq 0) 'Valid change was rejected'
}
foreach ($bad in @(
    @{Name='notes.md'; Text='- Fine.'}, @{Name='2026-02-30-invalid.md'; Text='- Fine.'},
    @{Name='2026-10-21-Upper.md'; Text='- Fine.'}, @{Name='2026-10-21-empty.md'; Text=''},
    @{Name='2026-10-21-blank-bullet.md'; Text='- '}, @{Name='2026-10-21-paragraph.md'; Text='No bullet'},
    @{Name='2026-10-21-two.md'; Text="- One.`n- Two."}, @{Name='2026-10-21-margin.md'; Text="- One.`nAt margin."}
)) {
    Case "reject $($bad.Name) without mutation" {
        param($r)
        Put $r ('Changes/' + $bad.Name) $bad.Text
        $before = Snapshot $r
        Refuses { & $fold -Root $r }
        Refuses { & $validate -Root $r }
        Assert ((Snapshot $r) -ceq $before) 'Invalid input mutated files'
    }
}
Case 'single-fragment final gate is read-only and requires folding' {
    param($r)
    Put $r 'Changes/2026-10-21-single.md' '- Single.'
    $before = Snapshot $r
    Refuses { & $fold -Root $r -Check }
    Refuses { & $validate -Root $r -RequireFolded }
    Assert ((Snapshot $r) -ceq $before) 'Read-only gate mutated files'
    & $fold -Root $r
    & $fold -Root $r -Check
    Assert ([IO.File]::ReadAllText((Join-Path $r 'CHANGELOG.md')).Contains('- Single.')) 'Single entry was lost'
}
Case 'preview writes nothing' {
    param($r)
    Put $r 'Changes/2026-10-21-preview.md' '- Preview.'
    $before = Snapshot $r
    & $fold -Root $r -WhatIf
    Assert ((Snapshot $r) -ceq $before) 'WhatIf mutated files'
}
Case 'relative roots follow PowerShell location without touching the process directory' {
    param($r)
    $processRoot = Join-Path $r 'process-checkout'
    $requestedRoot = Join-Path $r 'requested-checkout'
    foreach ($checkout in @($processRoot, $requestedRoot)) {
        Put $checkout 'Changes/README.md' '# Fragments'
        Put $checkout 'CHANGELOG.md' "# Changelog`n`n## Unreleased`n`n- Existing.`n"
    }
    Put $requestedRoot 'Changes/2026-10-21-relative.md' '- Requested checkout only.'
    $processBefore = Snapshot $processRoot
    $requestedBefore = Snapshot $requestedRoot
    $previousDirectory = [Environment]::CurrentDirectory
    Push-Location -LiteralPath $requestedRoot
    try {
        [Environment]::CurrentDirectory = $processRoot
        Refuses { & $validate -Root . -RequireFolded }
        Refuses { & $fold -Root . -Check }
        & $fold -Root . -WhatIf
        Assert ((Snapshot $requestedRoot) -ceq $requestedBefore) 'Relative preview mutated the requested checkout'
        & $fold -Root .
        & $validate -Root . -RequireFolded
        Assert ([IO.File]::ReadAllText((Join-Path $requestedRoot 'CHANGELOG.md')).Contains('- Requested checkout only.')) 'Relative fold did not publish to the requested checkout'
        Assert ((Snapshot $processRoot) -ceq $processBefore) 'Relative command touched the process-directory checkout'
    } finally {
        [Environment]::CurrentDirectory = $previousDirectory
        Pop-Location
    }
}
Case 'newest first, stable topic order, preserve released history and README' {
    param($r)
    Put $r 'Changes/2026-10-20-old.md' '- Older.'
    Put $r 'Changes/2026-10-21-a.md' '- A.'
    Put $r 'Changes/2026-10-21-b.md' "- B.`n  Detail."
    & $fold -Root $r
    $text = [IO.File]::ReadAllText((Join-Path $r 'CHANGELOG.md'))
    Assert ($text -ceq "# Changelog`n`n## Unreleased`n`n- B.`n  Detail.`n- A.`n- Older.`n- Existing.`n`n## 1.0`n`n- Released.`n") 'Fold order/history differs'
    Assert (@(Get-ChildItem -LiteralPath (Join-Path $r 'Changes') -Force).Count -eq 1) 'Fold left fragments or removed README'
    $before = Snapshot $r
    & $fold -Root $r
    Assert ((Snapshot $r) -ceq $before) 'Empty repeat mutated files'
}
Case 'preserve BOM, CRLF and Unicode bytes' {
    param($r)
    $path = Join-Path $r 'CHANGELOG.md'
    [IO.File]::WriteAllText($path, "# Changelog`r`n`r`n## Unreleased`r`n`r`n", [Text.UTF8Encoding]::new($true))
    Put $r 'Changes/2026-10-21-encoding.md' "- café 日本語.`n  More.`n"
    & $fold -Root $r
    $bytes = [IO.File]::ReadAllBytes($path)
    Assert ($bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) 'BOM was lost'
    Assert ([IO.File]::ReadAllText($path) -ceq "# Changelog`r`n`r`n## Unreleased`r`n`r`n- café 日本語.`r`n  More.`r`n") 'Encoding or EOL changed'
}
Case 'retry after publication cleans fragments without duplicate entries' {
    param($r)
    Put $r 'Changes/2026-10-21-retry.md' "- Existing.`n"
    & $fold -Root $r
    $text = [IO.File]::ReadAllText((Join-Path $r 'CHANGELOG.md'))
    Assert ([regex]::Matches($text, '(?m)^- Existing\.$').Count -eq 1) 'Retry duplicated a published entry'
}
foreach ($separator in @(@{Name='space'; Text=' '}, @{Name='tab'; Text="`t"})) {
    Case "released heading with $($separator.Name) does not suppress a new entry" {
        param($r)
        Put $r 'CHANGELOG.md' ("# Changelog`n`n## Unreleased`n`n- Existing.`n`n##" + $separator.Text + "1.0`n`n- Released.`n")
        Put $r 'Changes/2026-10-21-again.md' '- Released.'
        & $fold -Root $r
        Assert ([regex]::Matches([IO.File]::ReadAllText((Join-Path $r 'CHANGELOG.md')), '(?m)^- Released\.$').Count -eq 2) 'Released history suppressed a new entry'
    }
}
Case 'missing README, invalid headings, unexpected files and directories fail' {
    param($r)
    Remove-Item -LiteralPath (Join-Path $r 'Changes/README.md')
    Assert ((Get-ChangelogProblems $r).Count -gt 0) 'Missing README passed'
    Put $r 'Changes/README.md' '# Restored'
    foreach ($text in @('# No heading', "## Unreleased`n`n## Unreleased`n")) {
        Put $r 'CHANGELOG.md' $text
        Refuses { & $fold -Root $r }
    }
    Put $r 'CHANGELOG.md' "# Changelog`n`n## Unreleased`n"
    Put $r 'Changes/notes.txt' 'Unsupported'
    Assert ((Get-ChangelogProblems $r).Count -gt 0) 'Unexpected file passed'
    Remove-Item -LiteralPath (Join-Path $r 'Changes/notes.txt')
    [void][IO.Directory]::CreateDirectory((Join-Path $r 'Changes/nested'))
    Assert ((Get-ChangelogProblems $r).Count -gt 0) 'Nested directory passed'
}
Case 'overlapping fold fails without touching inputs' {
    param($r)
    Put $r 'Changes/2026-10-21-lock.md' '- Locked.'
    $lease = [IO.File]::Open((Join-Path $r '.changelog-fold.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    try {
        $before = Snapshot $r
        Refuses { & $fold -Root $r }
        Assert ((Snapshot $r) -ceq $before) 'Contended fold changed input'
    } finally { $lease.Dispose() }
}
Case 'Unreleased heading at EOF is supported' {
    param($r)
    Put $r 'CHANGELOG.md' "# Changelog`n`n## Unreleased"
    Put $r 'Changes/2026-10-21-eof.md' '- EOF.'
    & $fold -Root $r
    Assert ([IO.File]::ReadAllText((Join-Path $r 'CHANGELOG.md')) -ceq "# Changelog`n`n## Unreleased`n`n- EOF.`n") 'EOF heading merged incorrectly'
}
Case 'public final runner rejects pending fragments before build or test discovery' {
    param($r)
    $scoped = [IO.File]::Exists((Join-Path $repository 'Test-Changes.ps1'))
    $runnerName = if ($scoped) { 'Test-Changes.ps1' } else { 'test.ps1' }
    Put $r $runnerName ([IO.File]::ReadAllText((Join-Path $repository $runnerName)))
    foreach ($relative in @('Tools/Fold-Changelog.ps1','Tools/Modules/Documentation/Changelog.psm1')) {
        Put $r $relative ([IO.File]::ReadAllText((Join-Path $repository $relative)))
    }
    Put $r 'Changes/2026-10-21-final-gate.md' '- Pending handoff.'
    $start = [Diagnostics.ProcessStartInfo]::new((Get-Process -Id $PID).Path)
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($argument in @('-NoProfile','-File',(Join-Path $r $runnerName))) { $start.ArgumentList.Add($argument) }
    if ($scoped) { $start.ArgumentList.Add('-Mode'); $start.ArgumentList.Add('PrePush') }
    $result = Invoke-FixtureProcess $start
    Assert ($result.ExitCode -ne 0 -and $result.Output.Contains('pending changelog fragment')) "Final runner did not refuse pending fragments: $($result.Output)"
}
Case 'a timed-out fixture child is terminated before fixture cleanup' {
    param($r)
    $start = [Diagnostics.ProcessStartInfo]::new((Get-Process -Id $PID).Path)
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($argument in @('-NoProfile','-Command','Start-Sleep -Seconds 60')) { $start.ArgumentList.Add($argument) }
    $identity = $null
    $message = ''
    try {
        try { Invoke-FixtureProcess $start -TimeoutMilliseconds 100 -ProcessIdentity ([ref]$identity) | Out-Null }
        catch { $message = $_.Exception.Message }
        Assert ($message.Contains('did not exit within its budget') -and $null -ne $identity) 'Timeout negative control was not reached'
        $remaining = Get-Process -Id $identity.Id -ErrorAction SilentlyContinue
        try { Assert ($null -eq $remaining -or $remaining.StartTime.ToUniversalTime().Ticks -ne $identity.StartTicks) 'Owned fixture child survived timeout cleanup' }
        finally { if ($remaining) { $remaining.Dispose() } }
    } finally {
        # Keep this negative control self-cleaning even if the implementation regresses; PID alone is not ownership proof.
        if ($identity) {
            $remaining = Get-Process -Id $identity.Id -ErrorAction SilentlyContinue
            if ($remaining) {
                try {
                    if ($remaining.StartTime.ToUniversalTime().Ticks -eq $identity.StartTicks) {
                        $remaining.Kill($true)
                        Assert ($remaining.WaitForExit(5000)) 'Negative-control cleanup did not finish'
                    }
                } finally { $remaining.Dispose() }
            }
        }
    }
}
Write-Host "Changelog behavior: $($results.Count) passed, $($failures.Count) failed."
if ($failures.Count) { throw ($failures -join "`n") }
$global:LASTEXITCODE = 0
