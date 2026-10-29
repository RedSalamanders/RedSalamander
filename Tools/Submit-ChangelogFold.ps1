<#
.SYNOPSIS
Folds the Changes fragments merged to the default branch on a fold branch and opens its pull request.
.DESCRIPTION
Feature changes only add fragments; they are folded here, after they reach the default branch. Fetches that
branch, folds it in a temporary worktree outside this checkout, verifies the result is an exact fold, commits it,
pushes changelog/fold-<commit> and opens its pull request with gh. This checkout's branch, index and files never
change. Nothing is pushed when the branch has no fragments. A fold branch already pushed for the same commit is
reused, and an open fold pull request is reported instead of opening a second, conflicting one. The pull request
takes the repository's normal review and merge; merge it before folding again.
.PARAMETER Root
Repository root; defaults to the repository containing this command.
.PARAMETER Remote
Remote that holds the default branch; origin by default.
.PARAMETER Branch
Default branch to fold; main by default.
.PARAMETER Repository
GitHub repository (owner/name) of the pull request; read from the remote's URL by default.
.PARAMETER PushOnly
Pushes the fold branch without opening its pull request, for a remote gh cannot reach.
.OUTPUTS
Console summary naming the fold branch and pull request. WhatIf folds and commits only in the temporary
worktree, prints that commit, and neither pushes nor opens a pull request, also for a fold branch pushed earlier.
.NOTES
Requires PowerShell 7 and Git, and an authenticated gh unless PushOnly. Uses the network for git fetch/push and
the GitHub API. Commits with the repository's own Git identity, signing and hooks.
.EXAMPLE
./Tools/Submit-ChangelogFold.ps1 -WhatIf
.EXAMPLE
./Tools/Submit-ChangelogFold.ps1
#>
[CmdletBinding(SupportsShouldProcess)]
param([string] $Root = (Split-Path $PSScriptRoot -Parent), [string] $Remote = 'origin', [string] $Branch = 'main',
    [string] $Repository = '', [switch] $PushOnly)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Root)
function Invoke-FoldGit([string] $Directory, [string[]] $Arguments) {
    $output = @(& git -C $Directory @Arguments 2>&1 | ForEach-Object { "$_" })
    if ($LASTEXITCODE -ne 0) { throw "git $($Arguments -join ' ') failed ($LASTEXITCODE): $($output -join ' ')" }
    return , $output
}
function Invoke-FoldGh([string[]] $Arguments) {
    $output = @(& gh @Arguments --repo $Repository 2>&1 | ForEach-Object { "$_" })
    if ($LASTEXITCODE -ne 0) { throw "gh $($Arguments -join ' ') failed ($LASTEXITCODE): $($output -join ' ')" }
    return , $output
}
if (-not $PushOnly -and -not $Repository) {
    $url = (Invoke-FoldGit $Root @('remote', 'get-url', $Remote))[0].Trim()
    $match = [regex]::Match($url, 'github\.com[:/](?<repo>[^/]+/[^/]+?)(?:\.git)?/?$')
    if (-not $match.Success) { throw "$Remote ($url) is not a GitHub repository; pass -Repository, or use -PushOnly and open the pull request yourself." }
    $Repository = $match.Groups['repo'].Value
}
[void](Invoke-FoldGit $Root @('fetch', '--quiet', $Remote, "+refs/heads/${Branch}:refs/remotes/$Remote/$Branch"))
$commit = (Invoke-FoldGit $Root @('rev-parse', '--verify', "refs/remotes/$Remote/$Branch^{commit}"))[0].Trim()
$short = $commit.Substring(0, 12)
$foldBranch = "changelog/fold-$short"
if (-not $PushOnly) {
    # Assigned before filtering: piping the comma-wrapped array a helper returns would make it one item.
    $listed = Invoke-FoldGh @('pr', 'list', '--state', 'open', '--base', $Branch, '--json', 'url,headRefName',
        '--jq', '.[] | select(.headRefName | startswith("changelog/fold-")) | .url')
    $open = @($listed | Where-Object { $_ })
    if ($open.Count) { Write-Host "A fold pull request is already open; merge or close it, then run this again: $($open -join ', ')"; return }
}
$pushed = Invoke-FoldGit $Root @('ls-remote', '--heads', $Remote, "refs/heads/$foldBranch")
if (@($pushed | Where-Object { $_ }).Count) {
    Write-Host "$foldBranch is already pushed for $Branch at $short."
} else {
    # A short name: Git keeps this worktree's administration under .git/worktrees/<name>, within Windows path limits.
    $worktree = Join-Path ([IO.Path]::GetTempPath()) ('changelog-fold-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [void](Invoke-FoldGit $Root @('worktree', 'add', '--quiet', '--detach', $worktree, $commit))
    try {
        $changes = Join-Path $worktree 'Changes'
        $fragments = @(if ([IO.Directory]::Exists($changes)) { Get-ChildItem -LiteralPath $changes -Filter '*.md' -File | Where-Object Name -cne 'README.md' | Sort-Object Name -Descending })
        if (-not $fragments.Count) { Write-Host "$Remote/$Branch ($short) has no changelog fragments to fold."; return }
        # The default branch's own fold, then this command's exact-fold check against that branch. The temporary
        # worktree is folded even under WhatIf; only the push and the pull request are previews.
        & (Join-Path $worktree 'Tools/Fold-Changelog.ps1') -Root $worktree -WhatIf:$false
        & (Join-Path $worktree 'Tools/Fold-Changelog.ps1') -Root $worktree -Check
        & (Join-Path $PSScriptRoot 'Test-Changelog.ps1') -Root $worktree -Base $commit
        [void](Invoke-FoldGit $worktree @('add', '--all', '--', 'CHANGELOG.md', 'Changes'))
        $list = ($fragments | ForEach-Object { "- Changes/$($_.Name)" }) -join "`n"
        [void](Invoke-FoldGit $worktree @('commit', '--quiet', '-m', "Fold $($fragments.Count) changelog fragment(s)",
            '-m', "Folded from $Branch at $short by Tools/Submit-ChangelogFold.ps1:`n$list"))
        if (-not $PSCmdlet.ShouldProcess("$Remote $foldBranch", "Push the fold of $($fragments.Count) fragment(s) from $Branch at $short")) {
            Write-Host ((Invoke-FoldGit $worktree @('show', '--stat', '--format=%s%n%n%b', 'HEAD')) -join "`n")
            return
        }
        # Pushed from this checkout, where the remote (even a relative URL) resolves as configured; the worktree shares
        # its objects.
        $foldCommit = (Invoke-FoldGit $worktree @('rev-parse', 'HEAD'))[0].Trim()
        [void](Invoke-FoldGit $Root @('push', '--quiet', $Remote, "${foldCommit}:refs/heads/$foldBranch"))
        Write-Host "Pushed ${foldBranch}: $($fragments.Count) fragment(s) folded from $Branch at $short."
    } finally {
        # Only this command's own worktree is removed: a repository-wide prune would also drop the registration of
        # every worktree whose drive is offline right now. A failure here warns, so it never replaces the real result.
        $removal = @(& git -C $Root worktree remove --force $worktree 2>&1 | ForEach-Object { "$_" })
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "The temporary fold worktree $worktree was not unregistered ($($removal -join ' ')). Run: git worktree remove --force `"$worktree`""
        }
        if ([IO.Directory]::Exists($worktree)) {
            try { Remove-Item -LiteralPath $worktree -Recurse -Force }
            catch { Write-Warning "Delete the temporary fold worktree folder yourself: $worktree ($($_.Exception.Message))" }
        }
        $global:LASTEXITCODE = 0
    }
}
if ($PushOnly) { Write-Host "Open its pull request against $Branch."; return }
# A preview never writes to GitHub, also when the fold branch was already pushed and the fold itself was skipped.
if (-not $PSCmdlet.ShouldProcess("$Repository $foldBranch", "Open the fold pull request against $Branch")) { return }
$body = "Folds the changelog fragments merged to ``$Branch`` at $short into ``CHANGELOG.md``. Generated by ``Tools/Submit-ChangelogFold.ps1``; review and merge it like any other change."
$created = Invoke-FoldGh @('pr', 'create', '--base', $Branch, '--head', $foldBranch, '--title', 'Fold changelog fragments', '--body', $body)
Write-Host "Fold pull request: $($created[-1])"
