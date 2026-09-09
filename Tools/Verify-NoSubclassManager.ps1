<#
.SYNOPSIS
Rejects forbidden comctl32 subclass-manager calls in migrated product code.

.DESCRIPTION
Scans Common, Plugins, and RedSalamander for SetWindowSubclass, DefSubclassProc, and RemoveWindowSubclass. It uses ripgrep when available and a PowerShell source-file scan otherwise.

.OUTPUTS
Writes matched source locations on failure or one success message. It does not emit pipeline objects.

.NOTES
Prerequisites: a repository checkout; ripgrep is optional. Side effects: changes the current location to the repository root and performs read-only source scans. Exit code is 1 when forbidden calls remain and nonzero when scanning fails. Primary consumer: .github/workflows/ci.yml and the DxUi migration contract.

.EXAMPLE
.\Tools\Verify-NoSubclassManager.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$patterns = @(
    'SetWindowSubclass('
    'DefSubclassProc('
    'RemoveWindowSubclass('
)

$searchRoots = @(
    'Common'
    'Plugins'
    'RedSalamander'
)

$matches = @()
$ripgrep = Get-Command 'rg' -ErrorAction SilentlyContinue
if ($ripgrep)
{
    foreach ($pattern in $patterns)
    {
        $result = & $ripgrep.Source --line-number --color never --fixed-strings $pattern @searchRoots 2>$null
        if ($LASTEXITCODE -eq 0 -and $result)
        {
            $matches += $result
        }
        elseif ($LASTEXITCODE -gt 1)
        {
            throw "ripgrep failed while checking pattern '$pattern'."
        }
    }
}
else
{
    $sourceFiles = Get-ChildItem -Path $searchRoots -Recurse -File -Include *.c,*.cc,*.cpp,*.cxx,*.h,*.hh,*.hpp,*.hxx,*.inl
    foreach ($pattern in $patterns)
    {
        $matches += @(
            $sourceFiles |
                Select-String -SimpleMatch -Pattern $pattern |
                ForEach-Object { "$($_.Path):$($_.LineNumber):$($_.Line)" }
        )
    }
}

if ($matches.Count -gt 0)
{
    Write-Host "Forbidden comctl subclass-manager usage remains in migrated code paths:" -ForegroundColor Red
    $matches | Sort-Object -Unique | ForEach-Object { Write-Host $_ }
    exit 1
}

Write-Host "Verified: no SetWindowSubclass / DefSubclassProc / RemoveWindowSubclass usage remains under Common, Plugins, or RedSalamander." -ForegroundColor Green
$global:LASTEXITCODE = 0
