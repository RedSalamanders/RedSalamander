<#
.SYNOPSIS
    Counts lines of code across the RedSalamander solution and displays a categorised report.

.DESCRIPTION
    Uses cloc (Count Lines of Code) to measure every active source file in the repository,
    splitting results into three categories:

        Production   – shipping application and plugin code
        Test         – unit tests, integration tests, self-tests, performance tests
        DevEnv       – proof-of-concept projects, build scripts, installer tooling

    Only compiled / executed source files are counted (C/C++, PowerShell, MSBuild props).
    Data files, documentation, archived test runs, and third-party packages are excluded.

.PARAMETER Detailed
    When specified, prints per-project breakdowns inside each category.

.EXAMPLE
    .\Tools\Measure-SourceLines.ps1
    .\Tools\Measure-SourceLines.ps1 -Detailed
#>

[CmdletBinding()]
param(
    [switch]$Detailed
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Locate repository root ────────────────────────────────────────────────────

$repoRoot = (git -C $PSScriptRoot rev-parse --show-toplevel 2>$null)
if (-not $repoRoot) {
    $repoRoot = Split-Path $PSScriptRoot -Parent
}
$repoRoot = (Resolve-Path $repoRoot).Path

# ── Verify cloc is available ──────────────────────────────────────────────────

$clocCmd = Get-Command cloc -ErrorAction SilentlyContinue
if (-not $clocCmd) {
    Write-Error "cloc is not installed or not on PATH.  Install via:  winget install AlDanial.Cloc"
    exit 1
}
$clocVersion = & cloc --version 2>&1 | Select-Object -First 1

# ── Helpers ───────────────────────────────────────────────────────────────────

function Get-SourceFiles {
    param([string]$Root)

    # Only compiled / executed source code — no data, docs, or config markup.
    $extensions = @('.cpp', '.h', '.c', '.rc', '.inl',
                    '.ps1', '.psm1',
                    '.props', '.targets')

    $allFiles = @()
    $tracked   = git -C $Root ls-files --full-name 2>$null
    $untracked = git -C $Root ls-files --others --exclude-standard --full-name 2>$null
    if ($tracked)   { $allFiles += $tracked }
    if ($untracked) { $allFiles += $untracked }

    $allFiles | Sort-Object -Unique | Where-Object {
        $ext = [System.IO.Path]::GetExtension($_).ToLowerInvariant()
        $extensions -contains $ext
    } | Where-Object {
        $_ -notmatch '^(vcpkg_installed|\.build|\.vs|x64|ARM64|\.git|\.copilot|\.squad|Specs/TestRuns)[\\/]'
    }
}

function Classify-File {
    param([string]$RelativePath)

    $p = $RelativePath.Replace('\', '/')

    # ── Test code ─────────────────────────────────────────────────────────
    if ($p -match '^Tests/([^/]+)/') {
        return [PSCustomObject]@{ Category = 'Test'; Project = $Matches[1] }
    }
    if ($p -match '^RedSalamander/.*SelfTest') {
        return [PSCustomObject]@{ Category = 'Test'; Project = 'RedSalamander SelfTests' }
    }

    # ── Dev environment code ──────────────────────────────────────────────
    if ($p -match '^PoC/([^/]+)/') {
        return [PSCustomObject]@{ Category = 'DevEnv'; Project = "PoC/$($Matches[1])" }
    }
    if ($p -match '^Tools/') {
        return [PSCustomObject]@{ Category = 'DevEnv'; Project = 'Tools' }
    }
    if ($p -match '^Installer/') {
        return [PSCustomObject]@{ Category = 'DevEnv'; Project = 'Installer' }
    }
    if ($p -notmatch '/' -and $p -match '\.(ps1|props|targets)$') {
        return [PSCustomObject]@{ Category = 'DevEnv'; Project = 'Build Scripts' }
    }

    # ── Production code ───────────────────────────────────────────────────
    if ($p -match '^RedSalamander/')             { return [PSCustomObject]@{ Category = 'Production'; Project = 'RedSalamander' } }
    if ($p -match '^RedSalamanderMonitor/')       { return [PSCustomObject]@{ Category = 'Production'; Project = 'RedSalamanderMonitor' } }
    if ($p -match '^RedSalamanderSearchService/') { return [PSCustomObject]@{ Category = 'Production'; Project = 'SearchService' } }
    if ($p -match '^Common/Common/')              { return [PSCustomObject]@{ Category = 'Production'; Project = 'Common' } }
    if ($p -match '^Common/DxUi/')                { return [PSCustomObject]@{ Category = 'Production'; Project = 'DxUi' } }
    if ($p -match '^Common/PlugInterfaces/' -or $p -match '^Common/[^/]+\.(h|cpp)$') {
        return [PSCustomObject]@{ Category = 'Production'; Project = 'Common' }
    }
    if ($p -match '^Plugins/([^/]+)/') {
        return [PSCustomObject]@{ Category = 'Production'; Project = $Matches[1] }
    }

    return $null
}

# ── Discover and classify files ───────────────────────────────────────────────

Write-Host ""
Write-Host "  Scanning repository..." -ForegroundColor DarkGray
$allRelFiles = @(Get-SourceFiles -Root $repoRoot)

# Build lookup: relative path → { Category, Project }
$fileInfo    = @{}   # relPath → PSCustomObject
$allAbsPaths = [System.Collections.Generic.List[string]]::new()

foreach ($rel in $allRelFiles) {
    $info = Classify-File $rel
    if (-not $info) { continue }

    $absPath = Join-Path $repoRoot $rel.Replace('/', '\')
    if (-not (Test-Path $absPath)) { continue }

    $fileInfo[$absPath] = $info
    $allAbsPaths.Add($absPath)
}

Write-Host "  Found $($allAbsPaths.Count) source files.  Running cloc $clocVersion..." -ForegroundColor DarkGray

# ── Single cloc invocation ────────────────────────────────────────────────────

$listFile = [System.IO.Path]::GetTempFileName()
try {
    $allAbsPaths | Set-Content -Path $listFile -Encoding UTF8
    $rawCsv = & cloc --list-file="$listFile" --by-file --csv --quiet 2>$null
}
finally {
    Remove-Item $listFile -Force -ErrorAction SilentlyContinue
}

if (-not $rawCsv) {
    Write-Error "cloc returned no output."
    exit 1
}

# Parse per-file CSV rows
$clocRows = $rawCsv | ConvertFrom-Csv | Where-Object { $_.filename -and $_.filename -ne 'SUM' }

# Aggregate into Category → Project → totals
$catData = @{}  # Category → Project → @{ Files; Blank; Comment; Code }

foreach ($row in $clocRows) {
    $abs = $row.filename
    if (-not $fileInfo.ContainsKey($abs)) { continue }

    $cat  = $fileInfo[$abs].Category
    $proj = $fileInfo[$abs].Project

    if (-not $catData.ContainsKey($cat))            { $catData[$cat] = @{} }
    if (-not $catData[$cat].ContainsKey($proj))     { $catData[$cat][$proj] = @{ Files = 0; Blank = 0; Comment = 0; Code = 0 } }

    $catData[$cat][$proj].Files   += 1
    $catData[$cat][$proj].Blank   += [int]$row.blank
    $catData[$cat][$proj].Comment += [int]$row.comment
    $catData[$cat][$proj].Code    += [int]$row.code
}

# ── Display report ────────────────────────────────────────────────────────────

$categoryOrder  = @('Production', 'Test', 'DevEnv')
$categoryLabels = @{ Production = 'Production Code'; Test = 'Test Code'; DevEnv = 'Dev Environment' }
$categoryColors = @{ Production = 'Cyan'; Test = 'Yellow'; DevEnv = 'DarkGray' }

$grandTotals = @{ Files = 0; Blank = 0; Comment = 0; Code = 0 }
$catTotals   = @{}
foreach ($cat in $categoryOrder) { $catTotals[$cat] = @{ Files = 0; Blank = 0; Comment = 0; Code = 0 } }

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor White
Write-Host "  ║                     RedSalamander — Source Lines Report                             ║" -ForegroundColor White
Write-Host "  ╚══════════════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor White
Write-Host ""

foreach ($cat in $categoryOrder) {
    if (-not $catData.ContainsKey($cat)) { continue }

    $label = $categoryLabels[$cat]
    $color = $categoryColors[$cat]
    $projects = $catData[$cat]

    Write-Host "  ┌──────────────────────────────────────────────────────────────────────────────────────┐" -ForegroundColor $color
    Write-Host ("  │  {0,-84}│" -f $label.ToUpper()) -ForegroundColor $color
    Write-Host "  ├────────────────────────────────────┬─────────┬─────────┬───────────┬───────────────┤" -ForegroundColor $color
    Write-Host ("  │  {0,-34}│ {1,7} │ {2,7} │ {3,9} │ {4,13} │" -f "Project", "Files", "Blank", "Comment", "Code") -ForegroundColor $color
    Write-Host "  ├────────────────────────────────────┼─────────┼─────────┼───────────┼───────────────┤" -ForegroundColor $color

    $catFiles = 0; $catBlank = 0; $catComment = 0; $catCode = 0

    $sorted = $projects.GetEnumerator() | Sort-Object { -($_.Value.Code) }
    foreach ($entry in $sorted) {
        $projName = $entry.Key
        $t = $entry.Value

        $catFiles   += $t.Files
        $catBlank   += $t.Blank
        $catComment += $t.Comment
        $catCode    += $t.Code

        if ($Detailed -and $t.Files -gt 0) {
            $displayName = if ($projName.Length -gt 34) { $projName.Substring(0, 31) + "..." } else { $projName }
            Write-Host ("  │  {0,-34}│ {1,7:N0} │ {2,7:N0} │ {3,9:N0} │ {4,13:N0} │" -f $displayName, $t.Files, $t.Blank, $t.Comment, $t.Code) -ForegroundColor $color
        }
    }

    if ($Detailed -and $projects.Count -gt 1) {
        Write-Host "  ├────────────────────────────────────┼─────────┼─────────┼───────────┼───────────────┤" -ForegroundColor $color
    }

    Write-Host ("  │  {0,-34}│ {1,7:N0} │ {2,7:N0} │ {3,9:N0} │ {4,13:N0} │" -f "TOTAL", $catFiles, $catBlank, $catComment, $catCode) -ForegroundColor White
    Write-Host "  └────────────────────────────────────┴─────────┴─────────┴───────────┴───────────────┘" -ForegroundColor $color
    Write-Host ""

    $catTotals[$cat].Files   = $catFiles
    $catTotals[$cat].Blank   = $catBlank
    $catTotals[$cat].Comment = $catComment
    $catTotals[$cat].Code    = $catCode

    $grandTotals.Files   += $catFiles
    $grandTotals.Blank   += $catBlank
    $grandTotals.Comment += $catComment
    $grandTotals.Code    += $catCode
}

# ── Grand summary ─────────────────────────────────────────────────────────────

Write-Host "  ╔══════════════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor White
Write-Host "  ║  SUMMARY                                                                           ║" -ForegroundColor White
Write-Host "  ╠════════════════════════════════════╦═════════╦═════════╦═══════════╦═══════════════╣" -ForegroundColor White
Write-Host ("  ║  {0,-34}║ {1,7} ║ {2,7} ║ {3,9} ║ {4,13} ║" -f "Category", "Files", "Blank", "Comment", "Code") -ForegroundColor White
Write-Host "  ╠════════════════════════════════════╬═════════╬═════════╬═══════════╬═══════════════╣" -ForegroundColor White

foreach ($cat in $categoryOrder) {
    $label = $categoryLabels[$cat]
    $color = $categoryColors[$cat]
    $t = $catTotals[$cat]
    Write-Host ("  ║  {0,-34}║ {1,7:N0} ║ {2,7:N0} ║ {3,9:N0} ║ {4,13:N0} ║" -f $label, $t.Files, $t.Blank, $t.Comment, $t.Code) -ForegroundColor $color
}

Write-Host "  ╠════════════════════════════════════╬═════════╬═════════╬═══════════╬═══════════════╣" -ForegroundColor White
Write-Host ("  ║  {0,-34}║ {1,7:N0} ║ {2,7:N0} ║ {3,9:N0} ║ {4,13:N0} ║" -f "GRAND TOTAL", $grandTotals.Files, $grandTotals.Blank, $grandTotals.Comment, $grandTotals.Code) -ForegroundColor Green
Write-Host "  ╚════════════════════════════════════╩═════════╩═════════╩═══════════╩═══════════════╝" -ForegroundColor White
Write-Host ""

# ── Ratios ────────────────────────────────────────────────────────────────────

$prodCode = $catTotals['Production'].Code
$testCode = $catTotals['Test'].Code
$devCode  = $catTotals['DevEnv'].Code

if ($grandTotals.Code -gt 0) {
    $prodPct = [math]::Round(100.0 * $prodCode / $grandTotals.Code, 1)
    $testPct = [math]::Round(100.0 * $testCode / $grandTotals.Code, 1)
    $devPct  = [math]::Round(100.0 * $devCode  / $grandTotals.Code, 1)

    Write-Host "  Distribution:  Production $prodPct%  │  Test $testPct%  │  DevEnv $devPct%" -ForegroundColor DarkGray

    if ($prodCode -gt 0) {
        $ratio = [math]::Round($testCode / $prodCode, 2)
        Write-Host "  Test-to-Production ratio: ${ratio}:1" -ForegroundColor DarkGray
    }
}

Write-Host ""
