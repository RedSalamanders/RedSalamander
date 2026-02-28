<#
.SYNOPSIS
    Validates basic `Specs/` structure conventions.

.DESCRIPTION
    This script is intentionally simple and fast. It validates:

      - `Specs/` root contains no stray `*.md` files (except `Specs/README.md`).
      - No `________*` scratch files remain under `Specs/` (excluding `Specs/TestRuns/` artifacts).
      - RFC docs under `Specs/Plans/**` use the `RFC_*.md` naming convention.

    Exit code:
      - 0 on success
      - 1 if any violations are found

.EXAMPLE
    .\Tools\ValidateSpecs.ps1
#>

param()

$ErrorActionPreference = 'Stop'

function Resolve-RepoRoot() {
    $root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')
    return $root.Path
}

function Write-ErrorItem([string]$Title, [string[]]$Items) {
    Write-Host ''
    Write-Host $Title -ForegroundColor Red
    foreach ($i in $Items) {
        Write-Host ("  - {0}" -f $i) -ForegroundColor Red
    }
}

$repoRoot = Resolve-RepoRoot
$specsRoot = Join-Path $repoRoot 'Specs'
$testRunsRoot = Join-Path $specsRoot 'TestRuns'
$plansRoot = Join-Path $specsRoot 'Plans'

$hasError = $false

if (-not (Test-Path -LiteralPath $specsRoot)) {
    throw "Specs folder not found: $specsRoot"
}

# 1) No stray markdown at Specs root (except README.md)
$rootMd = Get-ChildItem -LiteralPath $specsRoot -File -Filter '*.md' | ForEach-Object { $_ }
$badRootMd = @($rootMd | Where-Object { $_.Name -ne 'README.md' } | ForEach-Object { $_.FullName })
if ($badRootMd.Count -gt 0) {
    $hasError = $true
    Write-ErrorItem "Unexpected markdown files at Specs root (only README.md is allowed):" $badRootMd
}

# 2) No "________*" scratch files in Specs (ignore TestRuns artifacts)
$scratch = Get-ChildItem -LiteralPath $specsRoot -Recurse -File -Filter '________*' | ForEach-Object { $_ }
$scratch = @($scratch | Where-Object { $_.FullName -notlike ($testRunsRoot + '*') } | ForEach-Object { $_.FullName })
if ($scratch.Count -gt 0) {
    $hasError = $true
    Write-ErrorItem "Scratch files should not exist under Specs/ (rename/move to Specs/Notes/):" $scratch
}

# 3) RFC naming under Specs/Plans/**
if (Test-Path -LiteralPath $plansRoot) {
    $planMd = Get-ChildItem -LiteralPath $plansRoot -Recurse -File -Filter '*.md' | ForEach-Object { $_ }
    $badRfc = @(
        $planMd |
            Where-Object { $_.Name -like '*RFC*' -and $_.Name -notlike 'RFC_*.md' } |
            ForEach-Object { $_.FullName }
    )
    if ($badRfc.Count -gt 0) {
        $hasError = $true
        Write-ErrorItem "RFC docs under Specs/Plans/** must be named RFC_*.md:" $badRfc
    }
}

if ($hasError) {
    Write-Host ''
    Write-Host 'ValidateSpecs: FAIL' -ForegroundColor Red
    exit 1
}

Write-Host 'ValidateSpecs: PASS' -ForegroundColor Green
exit 0

