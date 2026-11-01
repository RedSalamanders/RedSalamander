<#
.SYNOPSIS
    Formats tracked first-party C++ source with the repository-pinned clang-format.

.DESCRIPTION
    Selects tracked .cpp and .h files from RepositoryRoot while excluding vendor, specification,
    and agent-plan trees. Acquires and verifies the exact formatter pinned by
    Build/clang-format.json, then formats the selected files in place using .clang-format.

.PARAMETER RepositoryRoot
    Repository checkout whose tracked source files are selected. Defaults to this script's
    containing checkout.

.OUTPUTS
    Text status. The selected tracked source files may be rewritten; the pinned formatter
    cache is stored beneath .build/format.

.NOTES
    Requires PowerShell 7 and Git. A first run may download the pinned formatter wheel from
    PyPI; later runs use the verified repository-local cache. Does not format vendored files,
    archived specification evidence, or untracked files. Use -WhatIf to preview without
    acquiring the formatter or changing source files. Primary consumer: local source formatting.

.EXAMPLE
    .\format-all.ps1 -WhatIf

    Previews formatting of the tracked first-party source file set without installing the tool.

.EXAMPLE
    .\format-all.ps1

    Formats tracked first-party C++ files with the pinned formatter.
#>
#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [string]$RepositoryRoot = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = [IO.Path]::GetFullPath($RepositoryRoot)
$trackedPaths = @(& git -C $repositoryRoot ls-files --cached -- '*.cpp' '*.h')
if ($LASTEXITCODE -ne 0) {
    throw 'Could not enumerate tracked C++ source files with Git.'
}

$sourcePaths = @(
    $trackedPaths |
        Where-Object {
            $_ -and
            $_ -notmatch '^(?i:ThirdParty|Specs|advisor-plans|\.codex)/'
        } |
        Sort-Object -Unique |
        ForEach-Object {
            $path = Join-Path $repositoryRoot ($_ -replace '/', [IO.Path]::DirectorySeparatorChar)
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                $item = Get-Item -LiteralPath $path -Force
                if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                    throw "Refusing to format a reparse-point source file: $_"
                }
                $path
            }
        }
)
if ($sourcePaths.Count -eq 0) {
    Write-Host 'No tracked first-party C++ source files are available to format.'
    return
}

$pin = Get-Content -LiteralPath (Join-Path $repositoryRoot 'Build/clang-format.json') -Raw | ConvertFrom-Json
if ($pin.version -cnotmatch '^\d+\.\d+\.\d+$') {
    throw "The repository clang-format version is invalid: '$($pin.version)'."
}
$action = "Format $($sourcePaths.Count) tracked first-party C++ source files with clang-format $($pin.version)"
if (-not $PSCmdlet.ShouldProcess($repositoryRoot, $action)) {
    return
}

$formatterInstaller = Join-Path $repositoryRoot 'Tools/Install-ClangFormat.ps1'
$formatterPath = & $formatterInstaller
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace([string]$formatterPath)) {
    throw 'The pinned clang-format installer did not return a usable executable path.'
}
$formatterPath = [IO.Path]::GetFullPath([string]$formatterPath)
$versionOutput = @(& $formatterPath --version)
if ($LASTEXITCODE -ne 0 -or ($versionOutput -join "`n") -notmatch "^clang-format version $([regex]::Escape([string]$pin.version))(?:\s|$)") {
    throw "The formatter executable does not match pinned version $($pin.version): $formatterPath"
}

$failures = [Collections.Generic.List[string]]::new()
foreach ($sourcePath in $sourcePaths) {
    Write-Host "Formatting: $([IO.Path]::GetRelativePath($repositoryRoot, $sourcePath))"
    & $formatterPath -i --style=file $sourcePath
    if ($LASTEXITCODE -ne 0) {
        $failures.Add("$sourcePath (exit $LASTEXITCODE)")
    }
}
if ($failures.Count -gt 0) {
    throw "clang-format failed for $($failures.Count) file(s):`n$($failures -join "`n")"
}
Write-Host "Formatted $($sourcePaths.Count) tracked first-party C++ source files with clang-format $($pin.version)."
