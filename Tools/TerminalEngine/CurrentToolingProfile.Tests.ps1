Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$manifestPath = Join-Path $PSScriptRoot 'HistoricalTerminalProfiles.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$historicalPaths = [System.Collections.Generic.HashSet[string]]::new(
    [StringComparer]::OrdinalIgnoreCase)
foreach ($definition in @($manifest.historicalTests)) {
    [void]$historicalPaths.Add(([string]$definition.path).Replace('/', '\'))
}

$testsRoot = Join-Path $repoRoot 'Tools\Tests'
$currentTests = @(Get-ChildItem -LiteralPath $testsRoot -Filter '*.Tests.ps1' |
    Where-Object {
        $relativePath = [IO.Path]::GetRelativePath($repoRoot, $_.FullName)
        -not $historicalPaths.Contains($relativePath)
    } |
    Sort-Object Name)

if ($currentTests.Count -eq 0) {
    throw 'The current tooling Pester profile discovered no tests.'
}

foreach ($testFile in $currentTests) {
    . $testFile.FullName
}
