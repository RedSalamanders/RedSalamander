# Shared Changes workflow. Keep identical in DxUi, RedSalamander, RedXe, RedPrism, RedTilde and RedMonitor.
Set-StrictMode -Version Latest

$script:FragmentNamePattern = '^(?<date>\d{4}-\d{2}-\d{2})-[a-z0-9]+(?:-[a-z0-9]+)*\.md$'
$script:UnreleasedHeadingPattern = '(?m)^## Unreleased[ \t]*(?:\r?\n(?:[ \t]*\r?\n)?|$)'
$script:EntryPattern = '(?m)^- [^\r\n]*(?:\r?\n(?:[ \t][^\r\n]*|[ \t]*(?=\r?$)))*'
$script:StrictUtf8 = [Text.UTF8Encoding]::new($false, $true)

function Assert-ChangelogPath([string] $Path, [string] $Root) {
    # Links inside the checkout are refused, from the path up to and including the root, before anything is read or
    # written. A junction or synced folder above the root (where the user keeps the checkout) is not inspected.
    $rootPath = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    $current = [IO.Path]::GetFullPath($Path)
    while ($true) {
        if ((Test-Path -LiteralPath $current) -and ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw "Changelog paths must not traverse reparse points: $current"
        }
        if ([string]::Equals($current.TrimEnd('\', '/'), $rootPath, [StringComparison]::OrdinalIgnoreCase)) { return }
        $current = [IO.Path]::GetDirectoryName($current)
        if (-not $current) { throw "Changelog path is outside the repository ${Root}: $Path" }
    }
}

function Read-ChangelogText([string] $Path) {
    # Strict UTF-8 (a BOM is accepted): invalid bytes stop the command instead of becoming U+FFFD in a rewritten file.
    try { return [IO.File]::ReadAllText($Path, $script:StrictUtf8) }
    catch [Text.DecoderFallbackException] { throw "$Path is not valid UTF-8: $($_.Exception.Message)" }
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
    Assert-ChangelogPath $directory $Root
    $paths = [string[]]@([IO.Directory]::EnumerateFiles($directory, '*.md') | Where-Object { [IO.Path]::GetFileName($_) -cne 'README.md' })
    [Array]::Sort($paths, [StringComparer]::Ordinal)
    [Array]::Reverse($paths)
    return , @($paths | ForEach-Object {
        Assert-ChangelogPath $_ $Root
        [pscustomobject]@{ Name = [IO.Path]::GetFileName($_); Path = $_; Text = Read-ChangelogText $_ }
    })
}

function Get-ChangelogProblems([Parameter(Mandatory)][string] $Root, [switch] $RequireFolded) {
    $problems = [Collections.Generic.List[string]]::new()
    try {
        Assert-ChangelogPath $Root $Root
        $readme = Join-Path $Root 'Changes/README.md'
        Assert-ChangelogPath $readme $Root
        if (-not [IO.File]::Exists($readme)) { $problems.Add('Changes/README.md must explain changelog fragments') }
        $changelog = Join-Path $Root 'CHANGELOG.md'
        Assert-ChangelogPath $changelog $Root
        if (-not [IO.File]::Exists($changelog) -or [regex]::Matches((Read-ChangelogText $changelog), '(?m)^## Unreleased[ \t]*\r?$').Count -ne 1) {
            $problems.Add('CHANGELOG.md must keep exactly one "## Unreleased" heading, under which fragments are folded')
        }
        $directory = Join-Path $Root 'Changes'
        if ([IO.Directory]::Exists($directory)) {
            foreach ($entry in Get-ChildItem -LiteralPath $directory -Force) {
                Assert-ChangelogPath $entry.FullName $Root
                if ($entry.PSIsContainer -or ($entry.Name -cne 'README.md' -and $entry.Extension -cne '.md')) {
                    $problems.Add("Changes/$($entry.Name): only README.md and dated Markdown fragments belong here")
                }
            }
        }
        $fragments = Get-ChangelogFragments $Root
        foreach ($fragment in $fragments) {
            foreach ($problem in (Get-ChangelogFragmentProblems $fragment.Name $fragment.Text)) { $problems.Add("Changelog fragment $problem") }
        }
        if ($RequireFolded -and $fragments.Count) { $problems.Add("$($fragments.Count) pending changelog fragment(s) remain; they are folded on the default branch after merge (Tools/Submit-ChangelogFold.ps1)") }
    } catch { $problems.Add($_.Exception.Message) }
    return , $problems.ToArray()
}

function Split-Changelog([Parameter(Mandatory)][AllowEmptyString()][string] $Changelog) {
    # The text before the Unreleased heading, its entries in order (trimmed, LF), the rest of that section, and the
    # released history from the next "## " heading on. Line endings never count as a difference.
    $heading = [regex]::Match($Changelog, $script:UnreleasedHeadingPattern)
    if (-not $heading.Success) { throw 'CHANGELOG.md has no "## Unreleased" heading' }
    $at = $heading.Index + $heading.Length
    $next = [regex]::Match($Changelog.Substring($at), '(?m)^##[ \t]+')
    $end = if ($next.Success) { $at + $next.Index } else { $Changelog.Length }
    $section = $Changelog.Substring($at, $end - $at)
    return [pscustomobject]@{
        Before = $Changelog.Substring(0, $heading.Index) -replace '\r\n', "`n"
        Entries = [string[]]@([regex]::Matches($section, $script:EntryPattern) | ForEach-Object { $_.Value.Trim() -replace '\r\n', "`n" })
        Other = ([regex]::Replace($section, $script:EntryPattern, '') -replace '\s+', ' ').Trim()
        After = $Changelog.Substring($end) -replace '\r\n', "`n"
    }
}

function Read-ChangelogGit([Parameter(Mandatory)][string] $Root, [Parameter(Mandatory)][string[]] $Arguments, [switch] $Optional) {
    # Git output as stored, decoded as strict UTF-8 so no console code page alters it. -Optional returns $null when the
    # command fails (a missing object) instead of throwing.
    $start = [Diagnostics.ProcessStartInfo]::new('git')
    foreach ($argument in @('-C', $Root) + $Arguments) { $start.ArgumentList.Add($argument) }
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $process = [Diagnostics.Process]::Start($start)
    try {
        $buffer = [IO.MemoryStream]::new()
        $copy = $process.StandardOutput.BaseStream.CopyToAsync($buffer)
        $errors = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(60000)) { $process.Kill($true); throw "git $($Arguments -join ' ') did not finish within 60 seconds" }
        if (-not $copy.Wait(5000)) { throw "git $($Arguments -join ' ') output did not close" }
        if ($process.ExitCode -ne 0) {
            if ($Optional) { return $null }
            throw "git $($Arguments -join ' ') failed ($($process.ExitCode)): $($errors.GetAwaiter().GetResult().Trim())"
        }
        $bytes = $buffer.ToArray()
        $offset = if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { 3 } else { 0 }
        return $script:StrictUtf8.GetString($bytes, $offset, $bytes.Length - $offset)
    } finally { $process.Dispose() }
}

function Get-ChangelogFoldProblems([Parameter(Mandatory)][string] $Root, [Parameter(Mandatory)][string] $Base) {
    # Since this checkout's merge base with $Base, Changes/ may gain fragments, and CHANGELOG.md may change only by
    # folding fragments that the merge base holds and this checkout no longer has: the fold on the default branch.
    # Hand edits, entries folded from fragments this change adds, edited entries and unfolded removals are refused.
    $problems = [Collections.Generic.List[string]]::new()
    try {
        $Root = [IO.Path]::GetFullPath($Root)
        $mergeBase = (Read-ChangelogGit $Root @('merge-base', $Base, 'HEAD')).Trim()
        $baseText = Read-ChangelogGit $Root @('cat-file', 'blob', "${mergeBase}:CHANGELOG.md") -Optional
        # Before adoption there is no generated changelog to protect.
        if ($null -eq $baseText) { return , $problems.ToArray() }
        $changelog = Join-Path $Root 'CHANGELOG.md'
        Assert-ChangelogPath $changelog $Root
        $before = Split-Changelog $baseText
        $after = Split-Changelog (Read-ChangelogText $changelog)
        if ($before.Before -cne $after.Before -or $before.Other -cne $after.Other -or $before.After -cne $after.After) {
            $problems.Add("CHANGELOG.md changed outside its Unreleased entries since $Base; only folding fragments may change it")
        }
        $present = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $presentTexts = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        # Assigned first: the comma-wrapped array a function returns is one item to a foreach over the call itself.
        $fragments = Get-ChangelogFragments $Root
        foreach ($fragment in $fragments) {
            [void]$present.Add($fragment.Name)
            [void]$presentTexts.Add(($fragment.Text.Trim() -replace '\r\n', "`n"))
        }
        $folded = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
        foreach ($path in ((Read-ChangelogGit $Root @('ls-tree', '-z', '--name-only', $mergeBase, '--', 'Changes/')) -split "`0")) {
            $name = [IO.Path]::GetFileName($path)
            if (-not $name -or $name -ceq 'README.md' -or -not $name.EndsWith('.md', [StringComparison]::Ordinal) -or $present.Contains($name)) { continue }
            $folded[$name] = (Read-ChangelogGit $Root @('cat-file', 'blob', "${mergeBase}:$path")).Trim() -replace '\r\n', "`n"
        }
        $added = [Collections.Generic.List[string]]::new($after.Entries)
        foreach ($entry in $before.Entries) {
            if (-not $added.Remove($entry)) {
                $problems.Add("CHANGELOG.md: the Unreleased entry '$(($entry -split "`n")[0])' was edited or removed since $Base; entries change only through Changes/ fragments")
            }
        }
        $foldedTexts = [Collections.Generic.HashSet[string]]::new([string[]]@($folded.Values), [StringComparer]::Ordinal)
        foreach ($entry in $added) {
            if (-not $foldedTexts.Contains($entry)) {
                $problems.Add("CHANGELOG.md: '$(($entry -split "`n")[0])' was added by hand or folded from a fragment this change adds; keep it as a Changes/ fragment, which is folded on the default branch after merge")
            }
        }
        foreach ($name in $folded.Keys) {
            # A renamed pending fragment keeps its text under another name.
            if ($after.Entries -cnotcontains $folded[$name] -and -not $presentTexts.Contains($folded[$name])) {
                $problems.Add("Changes/$name was removed since $Base without its entry reaching CHANGELOG.md; restore it, since merged fragments are folded on the default branch")
            }
        }
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
    foreach ($entry in [regex]::Matches($unreleased, $script:EntryPattern)) {
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

function Write-ChangelogStage([string] $Path, [string] $Text, [bool] $Bom) {
    # Flushed to the disk before the rename publishes it, so a crash cannot publish a truncated changelog.
    $encoding = [Text.UTF8Encoding]::new($Bom)
    $bytes = [byte[]]($encoding.GetPreamble() + $encoding.GetBytes($Text))
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try { $stream.Write($bytes, 0, $bytes.Length); $stream.Flush($true) } finally { $stream.Dispose() }
}

function Invoke-ChangelogFold([Parameter(Mandatory)][string] $Root) {
    $Root = [IO.Path]::GetFullPath($Root)
    $problems = Get-ChangelogProblems $Root
    if ($problems.Count) { throw "Fix the changelog first:`n$($problems -join "`n")" }
    $lockPath = Join-Path $Root '.changelog-fold.lock'
    Assert-ChangelogPath $lockPath $Root
    $lease = [IO.FileStream]::new($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite,
        [IO.FileShare]::None, 1, [IO.FileOptions]::DeleteOnClose)
    $stage = $null
    try {
        $problems = Get-ChangelogProblems $Root
        if ($problems.Count) { throw "Fix the changelog first:`n$($problems -join "`n")" }
        $fragments = Get-ChangelogFragments $Root
        if (-not $fragments.Count) { return 0 }
        $changelog = Join-Path $Root 'CHANGELOG.md'
        $original = Read-ChangelogText $changelog
        $bytes = [IO.File]::ReadAllBytes($changelog)
        $bom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        $merged = Merge-ChangelogFragments $original $fragments
        $stage = Join-Path $Root ('.changelog-' + [guid]::NewGuid().ToString('N') + '.tmp')
        Assert-ChangelogPath $stage $Root
        Write-ChangelogStage $stage $merged $bom
        foreach ($fragment in $fragments) {
            Assert-ChangelogPath $fragment.Path $Root
            if ((Read-ChangelogText $fragment.Path) -cne $fragment.Text) { throw "Fragment changed during folding: $($fragment.Name)" }
        }
        Assert-ChangelogPath $changelog $Root
        if ((Read-ChangelogText $changelog) -cne $original) { throw 'CHANGELOG.md changed during folding; retry after the other editor finishes' }
        # Publish before removing any input. An interrupted cleanup is safely retryable.
        [IO.File]::Move($stage, $changelog, $true)
        $stage = $null
        foreach ($fragment in $fragments) {
            Assert-ChangelogPath $fragment.Path $Root
            if ((Read-ChangelogText $fragment.Path) -cne $fragment.Text) { throw "Fragment changed after publication; kept for review: $($fragment.Name)" }
            try { [IO.File]::Delete($fragment.Path) }
            catch {
                throw "CHANGELOG.md already holds the folded entries, but Changes/$($fragment.Name) could not be removed: $($_.Exception.Message) Run Tools/Fold-Changelog.ps1 again to finish; it does not duplicate published entries."
            }
        }
        return $fragments.Count
    } finally {
        if ($stage -and [IO.File]::Exists($stage)) { [IO.File]::Delete($stage) }
        $lease.Dispose()
    }
}

Export-ModuleMember -Function Get-ChangelogFragmentProblems, Get-ChangelogFragments, Get-ChangelogProblems, Get-ChangelogFoldProblems, Merge-ChangelogFragments, Invoke-ChangelogFold
