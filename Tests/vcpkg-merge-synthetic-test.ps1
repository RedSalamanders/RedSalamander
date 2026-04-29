<#
.SYNOPSIS
    Fast synthetic validation for Merge-RSVcpkgTripletSafe lock semantics.

.DESCRIPTION
    Exercises the merge helper without running vcpkg. Covers:
      [1] Unchanged destination held with FileShare.Read (read-share lock):
          merge skips the file, no error. Models IDE/indexer header locks
          like include\wil\com.h.
      [2] Changed destination held with FileShare.Read (replace cannot happen):
          merge fails with a clear actionable error mentioning the file path
          and lock guidance.
      [3] Destination held with FileShare.None (exclusive / unreadable):
          merge fails clearly because equality cannot be verified. We do NOT
          silently skip files we cannot prove equal to the staged output.
      [4] New files copy and missing nested directories are created.
      [5] A sibling triplet directory under the same install root is not
          touched by a merge into a different triplet.

.NOTES
    The helper is dot-sourced from Tools\VcpkgInstallSafety.ps1; this script
    must not redefine Merge-RSVcpkgTripletSafe.
#>

[CmdletBinding()]
param(
    [switch]$SkipCleanup
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

. (Join-Path $repoRoot 'Tools\VcpkgInstallSafety.ps1')

function New-TestFile {
    param([string]$Path, [string]$Content)
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $Content)
}

function Open-Lock {
    param(
        [string]$Path,
        [System.IO.FileShare]$Share
    )
    return [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        $Share)
}

$testRoot = Join-Path $repoRoot '.build\merge-synthetic-test'
$failures = @()

function Reset-TestRoot {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
    New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
}

try {
    Write-Host '=== Merge synthetic tests ===' -ForegroundColor Cyan

    # ---------- Test 1: unchanged + read-share lock -> skip ----------
    Write-Host ''
    Write-Host '[1] unchanged destination, FileShare.Read lock -> skip' -ForegroundColor Yellow
    Reset-TestRoot
    $src = Join-Path $testRoot 'src\test-triplet'
    $dst = Join-Path $testRoot 'dst\test-triplet'
    New-TestFile -Path (Join-Path $src 'include\header.h') -Content '// v1'
    New-TestFile -Path (Join-Path $dst 'include\header.h') -Content '// v1'
    $lock = Open-Lock -Path (Join-Path $dst 'include\header.h') -Share ([System.IO.FileShare]::Read)
    try {
        $r = Merge-RSVcpkgTripletSafe -SourcePath $src -DestinationPath $dst -TripletName 'test-triplet'
        if ($r.FilesSkipped -eq 1 -and $r.FilesCopied -eq 0) {
            Write-Host '    PASS' -ForegroundColor Green
        } else {
            $failures += "Test 1: expected 1 skip / 0 copy, got skip=$($r.FilesSkipped) copy=$($r.FilesCopied)"
            Write-Host "    FAIL: $($failures[-1])" -ForegroundColor Red
        }
    } catch {
        $failures += "Test 1: merge threw on unchanged read-share-locked file: $($_.Exception.Message)"
        Write-Host "    FAIL: $($failures[-1])" -ForegroundColor Red
    } finally {
        $lock.Dispose()
    }

    # ---------- Test 2: changed + read-share lock -> clear failure ----------
    Write-Host ''
    Write-Host '[2] changed destination, FileShare.Read lock (cannot replace) -> fail clearly' -ForegroundColor Yellow
    Reset-TestRoot
    $src = Join-Path $testRoot 'src\test-triplet'
    $dst = Join-Path $testRoot 'dst\test-triplet'
    New-TestFile -Path (Join-Path $src 'include\header.h') -Content '// v2 NEW'
    New-TestFile -Path (Join-Path $dst 'include\header.h') -Content '// v1 OLD'
    $lockPath = (Join-Path $dst 'include\header.h')
    $lock = Open-Lock -Path $lockPath -Share ([System.IO.FileShare]::Read)
    $threw = $false; $msg = ''
    try {
        Merge-RSVcpkgTripletSafe -SourcePath $src -DestinationPath $dst -TripletName 'test-triplet' | Out-Null
    } catch {
        $threw = $true
        $msg = $_.Exception.Message
    } finally {
        $lock.Dispose()
    }
    if ($threw -and $msg -match [regex]::Escape($lockPath) -and $msg -match 'differs from staged output' -and $msg -match 'Close Visual Studio') {
        Write-Host '    PASS' -ForegroundColor Green
    } else {
        $failures += "Test 2: expected clear lock-replace failure naming '$lockPath'. threw=$threw msg=$msg"
        Write-Host "    FAIL: $($failures[-1])" -ForegroundColor Red
    }

    # ---------- Test 3: exclusive lock -> fail because equality cannot be verified ----------
    Write-Host ''
    Write-Host '[3] exclusive (FileShare.None) lock -> fail (cannot verify equality)' -ForegroundColor Yellow
    Reset-TestRoot
    $src = Join-Path $testRoot 'src\test-triplet'
    $dst = Join-Path $testRoot 'dst\test-triplet'
    New-TestFile -Path (Join-Path $src 'include\header.h') -Content '// v1'
    New-TestFile -Path (Join-Path $dst 'include\header.h') -Content '// v1'
    $lockPath = (Join-Path $dst 'include\header.h')
    $lock = Open-Lock -Path $lockPath -Share ([System.IO.FileShare]::None)
    $threw = $false; $msg = ''
    try {
        Merge-RSVcpkgTripletSafe -SourcePath $src -DestinationPath $dst -TripletName 'test-triplet' | Out-Null
    } catch {
        $threw = $true
        $msg = $_.Exception.Message
    } finally {
        $lock.Dispose()
    }
    if ($threw -and $msg -match [regex]::Escape($lockPath) -and $msg -match 'verify equality' -and $msg -match 'Refusing to skip') {
        Write-Host '    PASS' -ForegroundColor Green
    } else {
        $failures += "Test 3: expected clear 'cannot verify equality / Refusing to skip' failure naming '$lockPath'. threw=$threw msg=$msg"
        Write-Host "    FAIL: $($failures[-1])" -ForegroundColor Red
    }

    # ---------- Test 4: new file + nested dirs ----------
    Write-Host ''
    Write-Host '[4] new file copy, nested directories created' -ForegroundColor Yellow
    Reset-TestRoot
    $src = Join-Path $testRoot 'src\test-triplet'
    $dst = Join-Path $testRoot 'dst\test-triplet'
    New-TestFile -Path (Join-Path $src 'lib\library.lib') -Content 'BIN'
    New-TestFile -Path (Join-Path $src 'deep\nested\path\file.txt') -Content 'D'
    New-Item -ItemType Directory -Path $dst -Force | Out-Null
    $r = Merge-RSVcpkgTripletSafe -SourcePath $src -DestinationPath $dst -TripletName 'test-triplet'
    $okLib  = (Test-Path (Join-Path $dst 'lib\library.lib')) -and ([System.IO.File]::ReadAllText((Join-Path $dst 'lib\library.lib')) -eq 'BIN')
    $okDeep = (Test-Path (Join-Path $dst 'deep\nested\path\file.txt'))
    if ($okLib -and $okDeep -and $r.FilesCopied -eq 2) {
        Write-Host '    PASS' -ForegroundColor Green
    } else {
        $failures += "Test 4: expected 2 copies and both files present (libOk=$okLib deepOk=$okDeep copied=$($r.FilesCopied))"
        Write-Host "    FAIL: $($failures[-1])" -ForegroundColor Red
    }

    # ---------- Test 5: sibling triplet untouched ----------
    Write-Host ''
    Write-Host '[5] sibling triplet directory untouched by another triplet merge' -ForegroundColor Yellow
    Reset-TestRoot
    $installRoot = Join-Path $testRoot 'installed'
    $srcA = Join-Path $testRoot 'src\triplet-a'
    $dstA = Join-Path $installRoot 'triplet-a'
    $dstB = Join-Path $installRoot 'triplet-b'
    New-TestFile -Path (Join-Path $srcA 'include\a.h') -Content 'A-NEW'
    New-TestFile -Path (Join-Path $dstA 'include\a.h') -Content 'A-OLD'
    New-TestFile -Path (Join-Path $dstB 'include\b.h') -Content 'B-KEEP'
    $bMtimeBefore = (Get-Item (Join-Path $dstB 'include\b.h')).LastWriteTimeUtc
    $bHashBefore  = Get-RSVcpkgFileSha256Hex -Path (Join-Path $dstB 'include\b.h')
    Merge-RSVcpkgTripletSafe -SourcePath $srcA -DestinationPath $dstA -TripletName 'triplet-a' | Out-Null
    $bExists = Test-Path (Join-Path $dstB 'include\b.h')
    $bHashAfter = if ($bExists) { Get-RSVcpkgFileSha256Hex -Path (Join-Path $dstB 'include\b.h') } else { $null }
    if ($bExists -and $bHashBefore -eq $bHashAfter) {
        Write-Host '    PASS' -ForegroundColor Green
    } else {
        $failures += "Test 5: sibling triplet was modified (exists=$bExists hashEqual=$($bHashBefore -eq $bHashAfter))"
        Write-Host "    FAIL: $($failures[-1])" -ForegroundColor Red
    }

    Write-Host ''
    if ($failures.Count -eq 0) {
        Write-Host '=== All synthetic merge tests passed ===' -ForegroundColor Green
        exit 0
    } else {
        Write-Host "=== $($failures.Count) synthetic merge test(s) failed ===" -ForegroundColor Red
        $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        exit 1
    }
} finally {
    if (-not $SkipCleanup -and (Test-Path -LiteralPath $testRoot)) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
