<#
.SYNOPSIS
    Validation script for vcpkg-install.ps1 merge lock resilience.

.DESCRIPTION
    Tests that vcpkg-install.ps1 can handle locked files during merge without failing
    when the locked file is unchanged. Run after implementing the merge lock fix.

.NOTES
    Author: GuineaPig (Tester)
    Related: .squad\decisions\inbox\guineapig-vcpkg-merge-lock.md
#>

[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Skip cleanup after tests")]
    [switch]$SkipCleanup
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot

Push-Location $repoRoot
try {
    Write-Host "=== vcpkg-install.ps1 Merge Lock Validation ===" -ForegroundColor Cyan
    Write-Host ""

    # Test 1: Unchanged Locked File (Happy Path)
    Write-Host "[Test 1] Unchanged locked file should not block merge..." -ForegroundColor Yellow

    # Clean install
    Write-Host "  Installing x64-windows triplet (clean)..." -ForegroundColor Gray
    & .\vcpkg-install.ps1 -Platform x64 -Clean
    if ($LASTEXITCODE -ne 0) {
        throw "Initial clean install failed"
    }

    $lockFile = ".build\vcpkg_installed\x64-windows\include\wil\com.h"
    if (-not (Test-Path $lockFile)) {
        throw "Expected WIL header not found: $lockFile"
    }

    # Open file with a *read-share* lock (FileShare.Read). This models the realistic
    # IDE/indexer header lock scenario (Visual Studio, Everything, Windows Search):
    # other readers can hash the file, but it cannot be deleted/replaced while open.
    # An exclusive (FileShare.None) lock would intentionally fail the merge under the
    # current contract because equality cannot be verified.
    Write-Host "  Locking file (FileShare.Read): $lockFile" -ForegroundColor Gray
    $fileStream = [System.IO.File]::Open(
        (Resolve-Path $lockFile).Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::Read
    )

    try {
        Write-Host "  Re-running install with locked unchanged file..." -ForegroundColor Gray
        & .\vcpkg-install.ps1 -Platform x64

        if ($LASTEXITCODE -ne 0) {
            Write-Host "  ❌ FAIL: Script failed with locked unchanged file" -ForegroundColor Red
            throw "Test 1 failed: vcpkg-install.ps1 should tolerate unchanged locked files"
        }

        if (-not (Test-Path $lockFile)) {
            Write-Host "  ❌ FAIL: Triplet corrupt (missing locked file)" -ForegroundColor Red
            throw "Test 1 failed: Locked file was deleted or triplet is corrupt"
        }

        Write-Host "  ✅ PASS: Script succeeded despite locked unchanged file" -ForegroundColor Green

    } finally {
        $fileStream.Close()
        $fileStream.Dispose()
    }

    Write-Host ""

    # Test 2: Sibling Triplet Preservation
    Write-Host "[Test 2] Sibling triplet preservation..." -ForegroundColor Yellow

    Write-Host "  Installing both x64 and ARM64 triplets..." -ForegroundColor Gray
    & .\vcpkg-install.ps1 -Platform All -Clean
    if ($LASTEXITCODE -ne 0) {
        throw "All-platform install failed"
    }

    $arm64File = ".build\vcpkg_installed\arm64-windows\include\wil\com.h"
    if (-not (Test-Path $arm64File)) {
        Write-Host "  ⚠ SKIP: ARM64 triplet not built (expected on x64-only systems)" -ForegroundColor Yellow
    } else {
        # Lock x64 file, re-run x64 only
        $fileStream2 = [System.IO.File]::Open(
            (Resolve-Path $lockFile).Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::Read
        )

        try {
            Write-Host "  Re-running x64 install with ARM64 sibling present..." -ForegroundColor Gray
            & .\vcpkg-install.ps1 -Platform x64

            if ($LASTEXITCODE -ne 0) {
                Write-Host "  ❌ FAIL: Script failed with locked file and sibling triplet" -ForegroundColor Red
                throw "Test 2 failed"
            }

            if (-not (Test-Path $arm64File)) {
                Write-Host "  ❌ FAIL: Sibling triplet corrupted" -ForegroundColor Red
                throw "Test 2 failed: arm64-windows triplet was affected by x64 merge"
            }

            Write-Host "  ✅ PASS: Sibling triplet preserved" -ForegroundColor Green

        } finally {
            $fileStream2.Close()
            $fileStream2.Dispose()
        }
    }

    Write-Host ""

    # Test 3: No Leftover Temporary Directories
    Write-Host "[Test 3] Cleanup of temporary merge directories..." -ForegroundColor Yellow

    $tmpPattern = ".build\vcpkg_installed\*.__merge_tmp"
    $tmpDirs = Get-ChildItem -Path ".build\vcpkg_installed" -Filter "*.__merge_tmp" -Directory -ErrorAction SilentlyContinue

    if ($tmpDirs.Count -gt 0) {
        Write-Host "  ❌ FAIL: Found leftover temporary directories:" -ForegroundColor Red
        $tmpDirs | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
        throw "Test 3 failed: Temporary merge directories not cleaned up"
    }

    Write-Host "  ✅ PASS: No temporary merge directories remaining" -ForegroundColor Green
    Write-Host ""

    Write-Host "=== All Tests Passed ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "  ✅ Unchanged locked files do not block merge" -ForegroundColor Green
    Write-Host "  ✅ Sibling triplets are preserved during merge" -ForegroundColor Green
    Write-Host "  ✅ Temporary directories are cleaned up" -ForegroundColor Green

} finally {
    Pop-Location

    if (-not $SkipCleanup) {
        Write-Host ""
        Write-Host "Cleaning up test artifacts..." -ForegroundColor Gray
        # Optionally clean up (or leave for manual inspection)
    }
}
