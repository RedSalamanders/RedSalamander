#Requires -Version 7.4
<#
.SYNOPSIS
Formats the C and C++ sources that CI formats, or verifies them with -Check.
.DESCRIPTION
The formatter is the one CI uses: Install-CiFormatter from this repository's .ci-engine installs the
clang-format build pinned in the engine's Policy/clang-format.json under .build/shared-format and checks
its SHA-256 and version. There is no other formatter lookup.
The files are the ones CI formats: tracked and new files that Git does not ignore, with one of the eight
C and C++ extensions, under a formatRoots entry of .ci/project.json and not matched by its formatExclude.
.PARAMETER Check
Verify without writing. Lists every file that needs formatting and fails if there is one.
#>
[CmdletBinding()]
param([switch]$Check)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$engine = Join-Path $root '.ci-engine'
$modules = Join-Path $engine 'Tools/Modules'
if (-not (Test-Path -LiteralPath (Join-Path $modules 'Ci.Format.psm1') -PathType Leaf)) {
    throw 'format.ps1 needs the shared CI engine in .ci-engine. Clone https://github.com/RedSalamanders/.github.git there and check out the engine.ref that .ci/project.json pins.'
}
Import-Module (Join-Path $modules 'Ci.Policy.psm1') -ErrorAction Stop
Import-Module (Join-Path $modules 'Ci.Format.psm1') -ErrorAction Stop

$project = Read-CiProject $root
$engineHead = (Invoke-CiGit $engine @('rev-parse', 'HEAD')).stdout.Trim()
if ($engineHead -cne $project.engine.ref) {
    throw ".ci-engine is at $engineHead, but .ci/project.json pins $($project.engine.ref). Run: git -C .ci-engine fetch origin; git -C .ci-engine checkout --detach $($project.engine.ref)"
}
$style = Join-Path $root '.clang-format'
if (-not (Test-Path -LiteralPath $style -PathType Leaf) -or [IO.File]::ReadAllText($style) -match 'InheritParentConfig') {
    throw 'Formatting requires the repository''s own root .clang-format, which must not inherit a style from outside the checkout.'
}

# The selection of Invoke-CiFormatting in .ci-engine/Tools/Modules/Ci.Format.psm1, applied to every file
# instead of only the files a pull request changes.
$extensions = @('.c', '.cc', '.cpp', '.cxx', '.h', '.hh', '.hpp', '.hxx')
$listed = (Invoke-CiGit $root @('-c', 'core.fsmonitor=false', 'ls-files', '--cached', '--others', '--exclude-standard', '-z')).stdout
$seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$sources = [Collections.Generic.List[string]]::new()
foreach ($path in $listed.Split([char]0, [StringSplitOptions]::RemoveEmptyEntries)) {
    if (-not $seen.Add($path)) { continue }
    if ([IO.Path]::GetExtension($path).ToLowerInvariant() -notin $extensions) { continue }
    $owned = @($project.formatRoots | Where-Object { $_ -eq '.' -or $path.StartsWith(($_.TrimEnd('/') + '/'), [StringComparison]::Ordinal) }).Count -gt 0
    if (-not $owned) { continue }
    if ($project.Contains('formatExclude') -and @($project.formatExclude | Where-Object { Test-CiGlob $path $_ }).Count -gt 0) { continue }
    if (-not (Test-Path -LiteralPath (Assert-CiContainedPath $root (Join-Path $root $path)) -PathType Leaf)) { continue }
    $sources.Add($path)
}
if (-not $sources.Count) { throw 'No C or C++ source file is under the formatRoots of .ci/project.json.' }

$formatter = Install-CiFormatter $root
$version = (Invoke-CiProcess $formatter @('--version') $root).stdout.Trim()
Write-Host "Formatter: $version ($formatter)"

# One clang-format process per processor, each given every n-th file from largest to smallest, so a pass
# takes about as long as the largest file. clang-format leaves an already formatted file untouched, so a
# clean pass rebuilds nothing.
$mode = if ($Check) { @('--style=file', '--Werror', '--dry-run') } else { @('--style=file', '--Werror', '-i') }
$characters = ($sources | Measure-Object -Property Length -Sum).Sum + 3 * $sources.Count
$count = [int][Math]::Min($sources.Count, [Math]::Max([Environment]::ProcessorCount, [Math]::Ceiling($characters / 12000)))
$batches = @(for ($index = 0; $index -lt $count; ++$index) { , [Collections.Generic.List[string]]::new() })
$ordered = @($sources | Sort-Object -Property { [IO.FileInfo]::new((Join-Path $root $_)).Length } -Descending)
for ($index = 0; $index -lt $ordered.Count; ++$index) { $batches[$index % $count].Add($ordered[$index]) }
$results = @($batches | ForEach-Object -ThrottleLimit ([Environment]::ProcessorCount) -Parallel {
        Set-Location -LiteralPath $using:root
        $executable = $using:formatter
        $arguments = @($using:mode) + '--' + $_
        $output = @(& $executable @arguments 2>&1 | ForEach-Object { $_.ToString() })
        [pscustomobject]@{ exitCode = $LASTEXITCODE; output = $output -join "`n" }
    })

$unformatted = [Collections.Generic.SortedSet[string]]::new([StringComparer]::Ordinal)
foreach ($result in $results) {
    if ($result.exitCode -eq 0) { continue }
    $violations = @([regex]::Matches($result.output, '(?m)^(.+?):\d+:\d+: error: code should be clang-formatted') | ForEach-Object { $_.Groups[1].Value })
    if (-not $Check -or -not $violations.Count) { throw "clang-format failed ($($result.exitCode)): $($result.output.Trim())" }
    foreach ($violation in $violations) { [void]$unformatted.Add($violation.Replace('\', '/')) }
}
if ($unformatted.Count) {
    foreach ($path in $unformatted) { Write-Host "Needs formatting: $path" }
    throw "$($unformatted.Count) of $($sources.Count) files need formatting. Run ./format.ps1 to format them."
}
Write-Host "$(if ($Check) { 'Checked' } else { 'Formatted' }) $($sources.Count) files."
