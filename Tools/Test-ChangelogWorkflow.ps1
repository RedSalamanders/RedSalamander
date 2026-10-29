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
Requires PowerShell 7 and Git for its local fixture repositories, which use a fixed identity and none of the
person's hooks or signing. No network, native code, application processes or desktop input.
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
$submit = Join-Path $repository 'Tools/Submit-ChangelogFold.ps1'
$tooling = @('Tools/Fold-Changelog.ps1', 'Tools/Test-Changelog.ps1', 'Tools/Modules/Documentation/Changelog.psm1')
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
    return (@(Get-ChildItem -LiteralPath $Root -Recurse -File | Where-Object { $_.Name -ne '.changelog-fold.lock' -and $_.FullName -notmatch '[\\/]\.git([\\/]|$)' } | Sort-Object FullName | ForEach-Object {
        [IO.Path]::GetRelativePath($Root, $_.FullName) + ':' + (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
    }) -join "`n")
}
function Invoke-FixtureGit([string] $Root, [string[]] $Arguments) {
    $output = @(& git -C $Root @Arguments 2>&1 | ForEach-Object { "$_" })
    Assert ($LASTEXITCODE -eq 0) "git $($Arguments -join ' ') failed in ${Root}: $($output -join ' ')"
    return , $output
}
function Get-FixtureGitLines([string] $Root, [string[]] $Arguments, [string] $Like = '*') {
    # Nonempty output lines as an array. Assigned first: piping the comma-wrapped array would make it one item.
    $lines = Invoke-FixtureGit $Root $Arguments
    return , [string[]]@($lines | Where-Object { $_ -and $_ -like $Like })
}
function New-FixtureRepository([string] $Path, [switch] $Bare) {
    # A throwaway repository: a fixed identity, exact bytes, and none of the person's global hooks or signing.
    [void][IO.Directory]::CreateDirectory($Path)
    [void](Invoke-FixtureGit $Path @(@('init', '-q') + $(if ($Bare) { @('--bare') } else { @() }) + @('-b', 'main')))
    foreach ($setting in @(@('user.name', 'Changelog fixture'), @('user.email', 'changelog@fixture.invalid'), @('commit.gpgsign', 'false'),
            @('tag.gpgsign', 'false'), @('core.autocrlf', 'false'), @('core.longpaths', 'true'), @('core.hooksPath', (Join-Path $Path '.no-hooks')))) {
        [void](Invoke-FixtureGit $Path (@('config') + $setting))
    }
}
function New-FixtureJunction([string] $Path, [string] $Target) {
    # Through the governed link creator where the repository has one, so its sandbox policy sees every test link.
    if (Get-Command New-RSTestSandboxLink -ErrorAction SilentlyContinue) { [void](New-RSTestSandboxLink -Path $Path -Target $Target -ItemType Junction) }
    else { [void](New-Item -ItemType Junction -Path $Path -Target $Target) }
}
function New-FoldFixture([string] $Root) {
    # A bare origin and a clone whose default branch holds two merged fragments and this suite's tooling.
    $origin = Join-Path $Root 'origin.git'
    $clone = Join-Path $Root 'clone'
    New-FixtureRepository $origin -Bare
    New-FixtureRepository $clone
    foreach ($relative in $tooling + @('Changes/README.md', 'CHANGELOG.md')) {
        $source = if ($relative.StartsWith('Tools/')) { Join-Path $repository $relative } else { Join-Path $Root $relative }
        Put $clone $relative ([IO.File]::ReadAllText($source))
    }
    Put $clone 'Changes/2026-10-20-first.md' '- First.'
    Put $clone 'Changes/2026-10-21-second.md' "- Second.`n  Detail."
    [void](Invoke-FixtureGit $clone @('add', '--all'))
    [void](Invoke-FixtureGit $clone @('commit', '-q', '-m', 'Merged changes'))
    # Relative, so the receiving repository path stays short however deep the scratch root is: Git for Windows refuses
    # a $GIT_DIR near 260 characters.
    [void](Invoke-FixtureGit $clone @('remote', 'add', 'origin', '../origin.git'))
    [void](Invoke-FixtureGit $clone @('push', '-q', 'origin', 'main'))
    $head = (Invoke-FixtureGit $clone @('rev-parse', 'HEAD'))[0].Trim()
    # The submission's temporary worktrees stay inside this fixture, never in the person's temporary folder. A short
    # name and long-path support keep a deep scratch root within Windows path limits.
    $worktrees = Join-Path $Root 'wt'
    [void][IO.Directory]::CreateDirectory($worktrees)
    return [pscustomobject]@{ Origin = $origin; Clone = $clone; Head = $head; FoldBranch = 'changelog/fold-' + $head.Substring(0, 12); WorktreeParent = $worktrees }
}
function Assert-NoFoldWorktree($Fixture) {
    $left = @(Get-ChildItem -LiteralPath $Fixture.WorktreeParent -Force | ForEach-Object { $_.Name })
    Assert ($left.Count -eq 0) "A temporary fold worktree was left behind: $($left -join ', ')"
}
function Save-FixtureBase([string] $Root) {
    # Commit everything as the default branch the comparisons and the runner read as origin/main.
    [void](Invoke-FixtureGit $Root @('add', '--all'))
    [void](Invoke-FixtureGit $Root @('commit', '-q', '-m', 'Fixture base'))
    [void](Invoke-FixtureGit $Root @('update-ref', 'refs/remotes/origin/main', 'HEAD'))
}
function Reset-FixtureBase([string] $Root) {
    [void](Invoke-FixtureGit $Root @('reset', '-q', '--hard', 'refs/remotes/origin/main'))
    [void](Invoke-FixtureGit $Root @('clean', '-q', '-f', '-d'))
}
function Edit-Fixture([string] $Root, [string] $Name, [string] $Old, [string] $New) {
    $path = Join-Path $Root $Name
    $text = [IO.File]::ReadAllText($path)
    Assert ($text.Contains($Old)) "Fixture $Name lacks '$Old'"
    [IO.File]::WriteAllText($path, $text.Replace($Old, $New), [Text.UTF8Encoding]::new($false))
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
Case 'CHANGELOG.md changes only by an exact fold of fragments the merge base holds' {
    param($r)
    New-FixtureRepository $r
    Put $r 'CHANGELOG.md' "# Changelog`n`n## Unreleased`n`n- Existing.`n- Second existing.`n`n## 1.0`n`n- Released.`n"
    Put $r 'Changes/2026-10-20-merged.md' "- Merged earlier.`n  With detail.`n"
    Save-FixtureBase $r
    # A feature change adds a fragment and leaves CHANGELOG.md alone.
    Put $r 'Changes/2026-10-21-feature.md' '- Feature.'
    Assert ((Get-ChangelogFoldProblems $r 'origin/main').Count -eq 0) 'An added fragment was refused'
    & $validate -Root $r -Base origin/main
    # The fold on the default branch: the merged fragment leaves, its entry arrives; line endings never count.
    Reset-FixtureBase $r
    & $fold -Root $r
    Assert ((Get-ChangelogFoldProblems $r 'origin/main').Count -eq 0) 'An exact fold was refused'
    $folded = Join-Path $r 'CHANGELOG.md'
    [IO.File]::WriteAllText($folded, ([IO.File]::ReadAllText($folded) -replace '\r?\n', "`r`n"))
    Assert ((Get-ChangelogFoldProblems $r 'origin/main').Count -eq 0) 'A line-ending-only difference was refused'
    # A renamed pending fragment keeps its text.
    Reset-FixtureBase $r
    Move-Item -LiteralPath (Join-Path $r 'Changes/2026-10-20-merged.md') -Destination (Join-Path $r 'Changes/2026-10-20-renamed.md')
    Assert ((Get-ChangelogFoldProblems $r 'origin/main').Count -eq 0) 'A renamed fragment was refused'
    foreach ($bad in @(
        @{ Name = 'a hand edit'; Expect = 'added by hand'; Change = { param($c) Edit-Fixture $c 'CHANGELOG.md' '- Existing.' "- Hand edit.`n- Existing." } },
        @{ Name = 'a fold of a fragment this change adds'; Expect = 'added by hand or folded from a fragment this change adds'; Change = { param($c) Put $c 'Changes/2026-10-21-new.md' '- New.'; & $fold -Root $c } },
        @{ Name = 'an edited entry'; Expect = 'was edited or removed'; Change = { param($c) Edit-Fixture $c 'CHANGELOG.md' '- Existing.' '- Existing, edited.' } },
        @{ Name = 'a released-history edit'; Expect = 'changed outside its Unreleased entries'; Change = { param($c) Edit-Fixture $c 'CHANGELOG.md' '- Released.' '- Released, edited.' } },
        @{ Name = 'an Unreleased paragraph'; Expect = 'changed outside its Unreleased entries'; Change = { param($c) Edit-Fixture $c 'CHANGELOG.md' "## Unreleased`n`n" "## Unreleased`n`nA note.`n`n" } },
        @{ Name = 'an unfolded removal'; Expect = 'was removed since origin/main without its entry'; Change = { param($c) Remove-Item -LiteralPath (Join-Path $c 'Changes/2026-10-20-merged.md') } },
        @{ Name = 'a duplicated folded entry'; Expect = 'is not the exact fold'; Change = { param($c) & $fold -Root $c; Edit-Fixture $c 'CHANGELOG.md' "- Merged earlier.`n  With detail.`n" "- Merged earlier.`n  With detail.`n- Merged earlier.`n  With detail.`n" } },
        @{ Name = 'reordered entries'; Expect = 'is not the exact fold'; Change = { param($c) Edit-Fixture $c 'CHANGELOG.md' "- Existing.`n- Second existing.`n" "- Second existing.`n- Existing.`n" } })) {
        Reset-FixtureBase $r
        & $bad.Change $r
        $problems = Get-ChangelogFoldProblems $r 'origin/main'
        Assert (@($problems | Where-Object { $_.Contains($bad.Expect) }).Count -eq 1) "$($bad.Name) was not refused as expected: $($problems -join ' | ')"
        Refuses { & $validate -Root $r -Base origin/main }
    }
    Reset-FixtureBase $r
    Refuses { & $validate -Root $r -Base refs/heads/no-such-base }
}
Case 'fold submission pushes one exact fold branch and leaves the caller checkout untouched' {
    param($r)
    $fixture = New-FoldFixture $r
    $origin = $fixture.Origin
    $clone = $fixture.Clone
    # A registered worktree whose folder is unavailable (an offline drive) must keep its registration.
    $offline = Join-Path $r 'offline-worktree'
    [void](Invoke-FixtureGit $clone @('worktree', 'add', '-q', '--detach', $offline, 'HEAD'))
    Remove-Item -LiteralPath $offline -Recurse -Force
    $registrations = Get-FixtureGitLines $clone @('worktree', 'list', '--porcelain') 'worktree *'
    Assert ($registrations.Count -eq 2) "The unavailable worktree was not registered: $($registrations -join ', ')"
    $before = Snapshot $clone
    $head = $fixture.Head
    $foldBranch = $fixture.FoldBranch
    $output = (& $submit -Root $clone -WorktreeParent $fixture.WorktreeParent -PushOnly -WhatIf 6>&1) -join "`n"
    Assert ($output.Contains(" in $(Join-Path $fixture.WorktreeParent 'changelog-fold-')")) "The fold worktree was not created in the sandbox folder it was given: $output"
    Assert ((Get-FixtureGitLines $origin @('branch', '--list', 'changelog/fold-*')).Count -eq 0) 'WhatIf pushed a fold branch'
    & $submit -Root $clone -WorktreeParent $fixture.WorktreeParent -PushOnly
    $changelog = (Invoke-FixtureGit $origin @('show', "${foldBranch}:CHANGELOG.md")) -join "`n"
    Assert ($changelog.Contains("## Unreleased`n`n- Second.`n  Detail.`n- First.`n- Existing.")) "The fold branch does not hold the exact fold: $changelog"
    $files = Get-FixtureGitLines $origin @('ls-tree', '--name-only', "${foldBranch}:Changes")
    Assert ($files.Count -eq 1 -and $files[0] -ceq 'README.md') "The fold branch kept fragments: $($files -join ', ')"
    Assert ((Snapshot $clone) -ceq $before) 'Submission changed the caller checkout files'
    Assert ((Invoke-FixtureGit $clone @('rev-parse', 'HEAD'))[0].Trim() -ceq $head) 'Submission moved the caller HEAD'
    Assert ((Get-FixtureGitLines $clone @('status', '--porcelain')).Count -eq 0) 'Submission left the caller checkout dirty'
    $after = Get-FixtureGitLines $clone @('worktree', 'list', '--porcelain') 'worktree *'
    Assert (($after -join "`n") -ceq ($registrations -join "`n")) "Submission left a temporary worktree or pruned another registration: $($after -join ', ')"
    & $submit -Root $clone -WorktreeParent $fixture.WorktreeParent -PushOnly
    $branches = Get-FixtureGitLines $origin @('branch', '--list', 'changelog/fold-*')
    Assert ($branches.Count -eq 1) "A repeated submission pushed a second fold branch: $($branches -join ', ')"
    # Once the fold is merged, the default branch has nothing left to fold.
    [void](Invoke-FixtureGit $origin @('update-ref', 'refs/heads/main', "refs/heads/$foldBranch"))
    $output = (& $submit -Root $clone -WorktreeParent $fixture.WorktreeParent -PushOnly 6>&1) -join "`n"
    Assert ($output.Contains('has no changelog fragments to fold')) "A folded default branch was folded again: $output"
    Assert-NoFoldWorktree $fixture
}
Case 'fold submission opens one pull request; WhatIf writes to neither the remote nor GitHub' {
    param($r)
    $fixture = New-FoldFixture $r
    # Stands in for gh: the submission resolves this function before any gh on PATH, and nothing reaches GitHub.
    $calls = [Collections.Generic.List[string]]::new()
    $openUrls = [Collections.Generic.List[string]]::new()
    function gh {
        $calls.Add(($args -join ' '))
        $global:LASTEXITCODE = 0
        if ($args[0] -ceq 'pr' -and $args[1] -ceq 'list') { $openUrls }
        elseif ($args[0] -ceq 'pr' -and $args[1] -ceq 'create') { 'https://github.com/fixture/example/pull/7' }
    }
    $creates = { @($calls | Where-Object { $_.StartsWith('pr create', [StringComparison]::Ordinal) }) }
    & $submit -Root $fixture.Clone -WorktreeParent $fixture.WorktreeParent -Repository fixture/example -WhatIf
    Assert (@(& $creates).Count -eq 0) 'WhatIf opened a pull request'
    Assert ((Get-FixtureGitLines $fixture.Origin @('branch', '--list', 'changelog/fold-*')).Count -eq 0) 'WhatIf pushed a fold branch'
    # A fold branch pushed earlier skips the fold itself, and WhatIf still opens nothing.
    & $submit -Root $fixture.Clone -WorktreeParent $fixture.WorktreeParent -PushOnly
    & $submit -Root $fixture.Clone -WorktreeParent $fixture.WorktreeParent -Repository fixture/example -WhatIf
    Assert (@(& $creates).Count -eq 0) 'WhatIf opened a pull request for a fold branch pushed earlier'
    & $submit -Root $fixture.Clone -WorktreeParent $fixture.WorktreeParent -Repository fixture/example
    $created = @(& $creates)
    Assert ($created.Count -eq 1 -and $created[0].Contains("--base main --head $($fixture.FoldBranch)") -and $created[0].EndsWith('--repo fixture/example')) "The fold pull request was not opened once against main: $($created -join ' | ')"
    # An open fold pull request is reported instead of opening a second, conflicting one.
    $openUrls.Add('https://github.com/fixture/example/pull/7')
    $output = (& $submit -Root $fixture.Clone -WorktreeParent $fixture.WorktreeParent -Repository fixture/example 6>&1) -join "`n"
    Assert (@(& $creates).Count -eq 1 -and $output.Contains('already open')) "A second fold pull request was opened: $output"
    # The open-PR query asks for every open pull request, not gh's default first 30.
    $lists = @($calls | Where-Object { $_.StartsWith('pr list', [StringComparison]::Ordinal) })
    Assert ($lists.Count -gt 0 -and @($lists | Where-Object { -not $_.Contains('--limit 1000') }).Count -eq 0) "An open-PR query was limited to gh's default page: $($lists -join ' | ')"
    Assert-NoFoldWorktree $fixture
}
Case 'a fold whose post-checkout hook fails reports it and leaves no worktree behind' {
    param($r)
    $fixture = New-FoldFixture $r
    # Git creates the worktree, then reports the hook's failure.
    $hooks = Join-Path $r 'failing-hooks'
    Put $hooks 'post-checkout' "#!/bin/sh`nexit 1`n"
    if (-not $IsWindows) { [IO.File]::SetUnixFileMode((Join-Path $hooks 'post-checkout'), [IO.UnixFileMode]'UserRead, UserWrite, UserExecute') }
    [void](Invoke-FixtureGit $fixture.Clone @('config', 'core.hooksPath', $hooks))
    $registrations = Get-FixtureGitLines $fixture.Clone @('worktree', 'list', '--porcelain') 'worktree *'
    $message = ''
    try { & $submit -Root $fixture.Clone -WorktreeParent $fixture.WorktreeParent -PushOnly } catch { $message = $_.Exception.Message }
    Assert ($message.Contains('git worktree add')) "The failing hook did not stop the fold with its error: $message"
    Assert-NoFoldWorktree $fixture
    $after = Get-FixtureGitLines $fixture.Clone @('worktree', 'list', '--porcelain') 'worktree *'
    Assert (($after -join "`n") -ceq ($registrations -join "`n")) "A failed fold left a worktree registration: $($after -join ', ')"
    Assert ((Get-FixtureGitLines $fixture.Origin @('branch', '--list', 'changelog/fold-*')).Count -eq 0) 'A failed fold pushed a branch'
}
Case 'invalid UTF-8 in CHANGELOG.md is refused without mutation' {
    param($r)
    [IO.File]::WriteAllBytes((Join-Path $r 'CHANGELOG.md'), [byte[]](@([Text.Encoding]::ASCII.GetBytes("# Changelog`n`n## Unreleased`n`n- Bad ")) + @(0xFF, 0x0A)))
    Put $r 'Changes/2026-10-21-utf8.md' '- Valid.'
    $before = Snapshot $r
    $message = ''
    try { & $fold -Root $r } catch { $message = $_.Exception.Message }
    Assert ($message.Contains('is not valid UTF-8')) "Invalid UTF-8 was not named: $message"
    Refuses { & $validate -Root $r }
    Assert ((Snapshot $r) -ceq $before) 'Invalid UTF-8 input was rewritten'
    # A UTF-16 byte order mark is not honored as another encoding: the file is still read as UTF-8 and refused.
    [IO.File]::WriteAllText((Join-Path $r 'CHANGELOG.md'), "# Changelog`n`n## Unreleased`n`n- Wide.`n", [Text.UnicodeEncoding]::new($false, $true))
    $before = Snapshot $r
    $message = ''
    try { & $fold -Root $r } catch { $message = $_.Exception.Message }
    Assert ($message.Contains('is not valid UTF-8')) "A UTF-16 changelog was not refused as UTF-8: $message"
    Assert ((Snapshot $r) -ceq $before) 'A UTF-16 changelog was rewritten'
}
Case 'links inside the checkout are refused; a junction above it is accepted' {
    param($r)
    if (-not $IsWindows) { Write-Host 'SKIP changelog junction controls: Windows junctions required.'; return }
    $holder = Join-Path $r 'holder'
    $inner = Join-Path $holder 'checkout'
    Put $inner 'Changes/README.md' '# Fragments'
    Put $inner 'CHANGELOG.md' "# Changelog`n`n## Unreleased`n"
    Put $inner 'Changes/2026-10-21-linked.md' '- Through a junction above the root.'
    $via = Join-Path $r 'via'
    New-FixtureJunction $via $holder
    & $fold -Root (Join-Path $via 'checkout')
    Assert ([IO.File]::ReadAllText((Join-Path $inner 'CHANGELOG.md')).Contains('- Through a junction above the root.')) 'A checkout below a junction was refused'
    $elsewhere = Join-Path $r 'elsewhere'
    Put $elsewhere '2026-10-21-outside.md' '- Outside.'
    Put $elsewhere 'README.md' '# Fragments'
    Remove-Item -LiteralPath (Join-Path $r 'Changes') -Recurse -Force
    New-FixtureJunction (Join-Path $r 'Changes') $elsewhere
    $before = Snapshot $elsewhere
    $message = ''
    try { & $fold -Root $r } catch { $message = $_.Exception.Message }
    Assert ($message.Contains('must not traverse reparse points') -and $message.Contains('Changes')) "A linked Changes folder was not refused: $message"
    Assert ((Snapshot $elsewhere) -ceq $before) 'A refused link target was changed'
}
Case 'an interrupted cleanup names the fragment and a rerun finishes it' {
    param($r)
    # Only Windows refuses to delete a file that another handle holds open without delete sharing.
    if (-not $IsWindows) { Write-Host 'SKIP changelog interrupted-cleanup control: Windows file sharing semantics required.'; return }
    $held = Join-Path $r 'Changes/2026-10-21-held.md'
    Put $r 'Changes/2026-10-21-held.md' '- Held.'
    # Another reader without delete sharing makes the removal fail after publication.
    $reader = [IO.File]::Open($held, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $message = ''
        try { & $fold -Root $r } catch { $message = $_.Exception.Message }
        Assert ($message.Contains('Changes/2026-10-21-held.md could not be removed') -and $message.Contains('Run Tools/Fold-Changelog.ps1 again')) "Interrupted cleanup gave no guidance: $message"
        Assert ([IO.File]::ReadAllText((Join-Path $r 'CHANGELOG.md')).Contains('- Held.')) 'The entry was not published before cleanup'
    } finally { $reader.Dispose() }
    & $fold -Root $r
    Assert (-not [IO.File]::Exists($held)) 'The rerun did not finish the cleanup'
    Assert ([regex]::Matches([IO.File]::ReadAllText((Join-Path $r 'CHANGELOG.md')), '(?m)^- Held\.$').Count -eq 1) 'The rerun duplicated the entry'
}
Case 'the public PrePush runner validates the changelog against origin/main before build or test discovery' {
    param($r)
    $scoped = [IO.File]::Exists((Join-Path $repository 'Test-Changes.ps1'))
    $runnerName = if ($scoped) { 'Test-Changes.ps1' } else { 'test.ps1' }
    New-FixtureRepository $r
    foreach ($relative in @($runnerName) + $tooling) { Put $r $relative ([IO.File]::ReadAllText((Join-Path $repository $relative))) }
    Save-FixtureBase $r
    $run = {
        $start = [Diagnostics.ProcessStartInfo]::new((Get-Process -Id $PID).Path)
        $start.UseShellExecute = $false
        $start.CreateNoWindow = $true
        $start.RedirectStandardOutput = $true
        $start.RedirectStandardError = $true
        foreach ($argument in @('-NoProfile','-File',(Join-Path $r $runnerName))) { $start.ArgumentList.Add($argument) }
        if ($scoped) { $start.ArgumentList.Add('-Mode'); $start.ArgumentList.Add('PrePush') }
        $outcome = Invoke-FixtureProcess $start
        # One line: a host without a console wraps a long error message, continuing it on " | " lines.
        $outcome | Add-Member -NotePropertyName Flat -NotePropertyValue ((($outcome.Output -replace '\r?\n\s*\|\s?', ' ') -replace '\s+', ' '))
        $outcome
    }
    Edit-Fixture $r 'CHANGELOG.md' '- Existing.' "- Hand edit.`n- Existing."
    $result = & $run
    Assert ($result.ExitCode -ne 0 -and $result.Flat.Contains('added by hand')) "The runner did not refuse a hand-edited changelog: $($result.Output)"
    Reset-FixtureBase $r
    Put $r 'Changes/2026-10-21-pending.md' '- Pending until the fold on main.'
    $result = & $run
    Assert ($result.Flat.Contains('Changelog valid: 1 pending fragment(s); CHANGELOG.md changes since origin/main are exact folds')) "The runner refused a pending fragment: $($result.Output)"
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
