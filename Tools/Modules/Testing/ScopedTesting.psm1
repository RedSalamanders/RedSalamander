# Native test inventory and naming for the repository-owned test runner.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-ScopedGit {
    param([string] $Root, [string[]] $Arguments)
    $start = [Diagnostics.ProcessStartInfo]::new('git')
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.CreateNoWindow = $true
    $start.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
    foreach ($argument in @('-c','core.quotepath=false','-C',$Root) + $Arguments) { $start.ArgumentList.Add($argument) }
    $process = [Diagnostics.Process]::Start($start)
    try {
        $errors = $process.StandardError.ReadToEndAsync()
        $output = $process.StandardOutput.ReadToEnd()
        $process.WaitForExit()
        $errorText = $errors.GetAwaiter().GetResult()
        if ($process.ExitCode) { throw "Git scope discovery failed: $errorText" }
        return $output
    } finally { $process.Dispose() }
}

function Get-ScopedTrackedPaths {
    param([string] $Root)
    return @((Invoke-ScopedGit $Root @('ls-files','--cached','--others','--exclude-standard','-z')) -split '\x00' | Where-Object { $_ } | Sort-Object -Unique)
}

function Assert-ScopedTestNames {
    param([string] $Root)
    $declared = @(Get-Content -LiteralPath (Join-Path $Root 'Tests/native-test-files.json') -Raw | ConvertFrom-Json)
    if (@($declared | Sort-Object -Unique).Count -ne $declared.Count) { throw 'Duplicate native test inventory paths.' }
    foreach ($path in $declared) {
        if ($path -match '^(Specs|Measurements|legacy|External)/') { throw "Historical/external source cannot enter the active test inventory: $path" }
        if ($path -cnotmatch '(^|/)[^/]+\.Tests\.[^/]+\.(cpp|h)$' -or -not (Test-Path -LiteralPath (Join-Path $Root $path) -PathType Leaf)) { throw "Invalid/missing native test source: $path" }
    }
    $live = @(Get-ScopedTrackedPaths $Root | Where-Object {
        $_ -notmatch '^(Specs|Measurements|legacy|External)/' -and
        $_ -cne 'Tools/TerminalEngine/TerminalEngineGate0ContractTestAdapter.cpp' -and
        $_ -match '\.(cpp|h)$' -and ($_ -match '^Tests/|/SelfTest/|(^|/)[^/]*(Test|Mock|Fake)[^/]*\.(cpp|h)$') -and
        (Test-Path -LiteralPath (Join-Path $Root $_) -PathType Leaf)
    })
    foreach ($path in $live) { if ($path -notin $declared) { throw "Native test source is absent from Tests/native-test-files.json: $path" } }
    return $declared.Count
}

Export-ModuleMember -Function Assert-ScopedTestNames
