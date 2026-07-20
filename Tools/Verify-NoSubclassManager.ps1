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
