# Shared Changes workflow. Keep identical in DxUi, RedSalamander, RedXe, RedPrism and RedTilde.
Set-StrictMode -Version Latest

$script:FragmentNamePattern = '^(?<date>\d{4}-\d{2}-\d{2})-[a-z0-9]+(?:-[a-z0-9]+)*\.md$'
$script:UnreleasedHeadingPattern = '(?m)^## Unreleased[ \t]*(?:\r?\n(?:[ \t]*\r?\n)?|$)'

function Assert-ChangelogPath([string] $Path) {
    $current = [IO.Path]::GetFullPath($Path)
    while ($current) {
        if ((Test-Path -LiteralPath $current) -and ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw "Changelog paths must not traverse reparse points: $current"
        }
        $current = [IO.Path]::GetDirectoryName($current)
    }
}

function Get-ChangelogFragmentProblems([Parameter(Mandatory)][string] $Name, [AllowEmptyString()][string] $Text) {
    # A fragment is named <yyyy-mm-dd>-<topic>.md and holds one change: one top-level Markdown bullet, whose continuation
    # lines and sub-bullets are indented as in CHANGELOG.md.
    $problems = [Collections.Generic.List[string]]::new()
    $match = [regex]::Match($Name, $script:FragmentNamePattern)
    $parsed = [DateTime]::MinValue
    if (-not $match.Success) {
        $problems.Add("${Name}: name it <yyyy-mm-dd>-<topic>.md, the topic in lowercase letters, digits and hyphens")
    } elseif (-not [DateTime]::TryParseExact($match.Groups['date'].Value, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        $problems.Add("${Name}: $($match.Groups['date'].Value) is not a calendar date")
    }
    $lines = @($Text -split '\r?\n' | Where-Object { $_.Trim() })
    if ($lines.Count -eq 0 -or $lines[0] -notmatch '^- \S') {
        $problems.Add("${Name}: start with a Markdown bullet ('- ')")
    }
    $entries = @($lines | Where-Object { $_.StartsWith('- ', [StringComparison]::Ordinal) })
    if ($entries.Count -gt 1) { $problems.Add("${Name}: holds $($entries.Count) entries; record one change and indent its sub-bullets") }
    if (@($lines | Select-Object -Skip 1 | Where-Object { -not $_.StartsWith(' ', [StringComparison]::Ordinal) -and -not $_.StartsWith('- ', [StringComparison]::Ordinal) }).Count) {
        $problems.Add("${Name}: indent every line after the bullet")
    }
    return , $problems.ToArray()
}

function Get-ChangelogFragments([Parameter(Mandatory)][string] $Root) {
    # The fragments under Changes/, newest first: their dated names sort that way, and fragments of one day by topic.
    $directory = Join-Path $Root 'Changes'
    if (-not [IO.Directory]::Exists($directory)) { return , @() }
    Assert-ChangelogPath $directory
    $paths = [string[]]@([IO.Directory]::EnumerateFiles($directory, '*.md') | Where-Object { [IO.Path]::GetFileName($_) -cne 'README.md' })
    [Array]::Sort($paths, [StringComparer]::Ordinal)
    [Array]::Reverse($paths)
    return , @($paths | ForEach-Object {
        Assert-ChangelogPath $_
        [pscustomobject]@{ Name = [IO.Path]::GetFileName($_); Path = $_; Text = [IO.File]::ReadAllText($_, [Text.UTF8Encoding]::new($false, $true)) }
    })
}

function Get-ChangelogProblems([Parameter(Mandatory)][string] $Root, [switch] $RequireFolded) {
    $problems = [Collections.Generic.List[string]]::new()
    try {
        Assert-ChangelogPath $Root
        $readme = Join-Path $Root 'Changes/README.md'
        Assert-ChangelogPath $readme
        if (-not [IO.File]::Exists($readme)) { $problems.Add('Changes/README.md must explain changelog fragments') }
        $changelog = Join-Path $Root 'CHANGELOG.md'
        Assert-ChangelogPath $changelog
        if (-not [IO.File]::Exists($changelog) -or [regex]::Matches([IO.File]::ReadAllText($changelog), '(?m)^## Unreleased[ \t]*\r?$').Count -ne 1) {
            $problems.Add('CHANGELOG.md must keep exactly one "## Unreleased" heading, under which fragments are folded')
        }
        $directory = Join-Path $Root 'Changes'
        if ([IO.Directory]::Exists($directory)) {
            foreach ($entry in Get-ChildItem -LiteralPath $directory -Force) {
                Assert-ChangelogPath $entry.FullName
                if ($entry.PSIsContainer -or ($entry.Name -cne 'README.md' -and $entry.Extension -cne '.md')) {
                    $problems.Add("Changes/$($entry.Name): only README.md and dated Markdown fragments belong here")
                }
            }
        }
        $fragments = Get-ChangelogFragments $Root
        foreach ($fragment in $fragments) {
            foreach ($problem in (Get-ChangelogFragmentProblems $fragment.Name $fragment.Text)) { $problems.Add("Changelog fragment $problem") }
        }
        if ($RequireFolded -and $fragments.Count) { $problems.Add("$($fragments.Count) pending changelog fragment(s): run Tools/Fold-Changelog.ps1 before PrePush or final handoff, even for one fragment") }
    } catch { $problems.Add($_.Exception.Message) }
    return , $problems.ToArray()
}

function Merge-ChangelogFragments([Parameter(Mandatory)][string] $Changelog, [object[]] $Fragments) {
    # The changelog with the fragments' entries first under its Unreleased heading, in the order given, in the
    # changelog's own line endings.
    $heading = [regex]::Match($Changelog, $script:UnreleasedHeadingPattern)
    if (-not $heading.Success) { throw 'CHANGELOG.md has no "## Unreleased" heading to fold the fragments under' }
    if (-not $Fragments -or $Fragments.Count -eq 0) { return $Changelog }
    $newline = if ($Changelog.Contains("`r`n")) { "`r`n" } else { "`n" }
    $at = $heading.Index + $heading.Length
    $unreleased = ($Changelog.Substring($at) -split '(?m)^##[ \t]+', 2)[0]
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entry in [regex]::Matches($unreleased, '(?m)^- [^\r\n]*(?:\r?\n(?:[ \t][^\r\n]*|[ \t]*(?=\r?$)))*')) {
        [void]$seen.Add(($entry.Value.Trim() -replace '\r\n', "`n"))
    }
    $entries = foreach ($fragment in $Fragments) {
        $normalized = $fragment.Text.Trim() -replace '\r\n', "`n"
        # Retry after publication but before cleanup must not duplicate the published entry.
        if ($seen.Add($normalized)) { ($normalized -replace "`n", $newline) + $newline }
    }
    $separator = if ($at -gt 0 -and $Changelog[$at - 1] -ne "`n") { $newline + $newline } else { '' }
    return $Changelog.Substring(0, $at) + $separator + ($entries -join '') + $Changelog.Substring($at)
}

function Invoke-ChangelogFold([Parameter(Mandatory)][string] $Root) {
    $Root = [IO.Path]::GetFullPath($Root)
    $problems = Get-ChangelogProblems $Root
    if ($problems.Count) { throw "Fix the changelog first:`n$($problems -join "`n")" }
    $lockPath = Join-Path $Root '.changelog-fold.lock'
    Assert-ChangelogPath $lockPath
    $lease = [IO.FileStream]::new($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite,
        [IO.FileShare]::None, 1, [IO.FileOptions]::DeleteOnClose)
    $stage = $null
    try {
        $problems = Get-ChangelogProblems $Root
        if ($problems.Count) { throw "Fix the changelog first:`n$($problems -join "`n")" }
        $fragments = Get-ChangelogFragments $Root
        if (-not $fragments.Count) { return 0 }
        $changelog = Join-Path $Root 'CHANGELOG.md'
        $original = [IO.File]::ReadAllText($changelog)
        $bytes = [IO.File]::ReadAllBytes($changelog)
        $bom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        $merged = Merge-ChangelogFragments $original $fragments
        $stage = Join-Path $Root ('.changelog-' + [guid]::NewGuid().ToString('N') + '.tmp')
        [IO.File]::WriteAllText($stage, $merged, [Text.UTF8Encoding]::new($bom))
        foreach ($fragment in $fragments) {
            Assert-ChangelogPath $fragment.Path
            if ([IO.File]::ReadAllText($fragment.Path) -cne $fragment.Text) { throw "Fragment changed during folding: $($fragment.Name)" }
        }
        Assert-ChangelogPath $changelog
        if ([IO.File]::ReadAllText($changelog) -cne $original) { throw 'CHANGELOG.md changed during folding; retry after the other editor finishes' }
        # Publish before removing any input. An interrupted cleanup is safely retryable.
        [IO.File]::Move($stage, $changelog, $true)
        $stage = $null
        foreach ($fragment in $fragments) {
            Assert-ChangelogPath $fragment.Path
            if ([IO.File]::ReadAllText($fragment.Path) -cne $fragment.Text) { throw "Fragment changed after publication; kept for review: $($fragment.Name)" }
            [IO.File]::Delete($fragment.Path)
        }
        return $fragments.Count
    } finally {
        if ($stage -and [IO.File]::Exists($stage)) { [IO.File]::Delete($stage) }
        $lease.Dispose()
    }
}

Export-ModuleMember -Function Get-ChangelogFragmentProblems, Get-ChangelogFragments, Get-ChangelogProblems, Merge-ChangelogFragments, Invoke-ChangelogFold
