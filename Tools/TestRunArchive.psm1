Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function ConvertTo-RSRepoRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $fullPath = if ([IO.Path]::IsPathRooted($Path)) {
        [IO.Path]::GetFullPath($Path)
    } else {
        [IO.Path]::GetFullPath((Join-Path $root $Path))
    }
    if (-not $fullPath.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Archive path is outside the repository: $Path"
    }

    return $fullPath.Substring($root.Length + 1)
}

function Get-RSTestRunArchiveViolations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Paths,

        [long]$MaxFileBytes = 2MB,

        [long]$MaxRunBytes = 5MB
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $runKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $violations = [System.Collections.Generic.List[object]]::new()

    foreach ($path in $Paths | Sort-Object -Unique) {
        $relativePath = ConvertTo-RSRepoRelativePath -RepoRoot $root -Path $path
        $fullPath = Join-Path $root $relativePath
        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
            continue
        }

        $parts = $relativePath -split '[\\/]'
        if ($parts.Count -lt 5 -or $parts[0] -ne 'Specs' -or $parts[1] -ne 'TestRuns') {
            continue
        }

        $fileInfo = Get-Item -LiteralPath $fullPath
        if ($fileInfo.Length -gt $MaxFileBytes) {
            $violations.Add([pscustomobject]@{
                Kind = 'FileSize'
                Path = $relativePath
                Message = "Checked-in TestRuns file is $($fileInfo.Length) bytes; maximum is $MaxFileBytes bytes."
            })
        }

        $runKey = ($parts[0..4] -join '/')
        [void]$runKeys.Add($runKey)

        if ($fileInfo.Name -eq 'perf_metrics.jsonl') {
            $reader = [IO.File]::OpenText($fullPath)
            try {
                $firstLine = $reader.ReadLine()
            } finally {
                $reader.Dispose()
            }

            if (-not [string]::IsNullOrWhiteSpace($firstLine)) {
                try {
                    $firstMetric = $firstLine | ConvertFrom-Json
                } catch {
                    $violations.Add([pscustomobject]@{
                        Kind = 'PerfJson'
                        Path = $relativePath
                        Message = 'The first perf metric row is not valid JSON.'
                    })
                    continue
                }

                $profile = $parts[2]
                $machineHash = [string]$firstMetric.machineHash
                if ($profile -match '^[0-9a-fA-F]{8,40}$' -and -not [string]::IsNullOrWhiteSpace($machineHash) -and
                    -not [string]::Equals($profile, $machineHash, [StringComparison]::OrdinalIgnoreCase)) {
                    $violations.Add([pscustomobject]@{
                        Kind = 'MachineProfile'
                        Path = $relativePath
                        Message = "Archive profile '$profile' does not match embedded machineHash '$machineHash'."
                    })
                }
            }
        }
    }

    foreach ($runKey in $runKeys) {
        $runRoot = Join-Path $root ($runKey -replace '/', [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $runRoot -PathType Container)) {
            continue
        }

        [long]$runBytes = 0L
        foreach ($runFile in Get-ChildItem -LiteralPath $runRoot -File -Recurse -Force) {
            $runBytes += [long]$runFile.Length
        }

        if ($runBytes -gt $MaxRunBytes) {
            $violations.Add([pscustomobject]@{
                Kind = 'RunSize'
                Path = $runKey
                Message = "Checked-in TestRuns run is $runBytes bytes; maximum is $MaxRunBytes bytes."
            })
        }
    }

    return @($violations)
}

function Get-RSChangedTestRunArchivePaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $paths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $root = [IO.Path]::GetFullPath($RepoRoot)

    $hasOriginMaster = $false
    & git -C $root rev-parse --verify origin/master *> $null
    if ($LASTEXITCODE -eq 0) {
        $hasOriginMaster = $true
    }

    $commands = [System.Collections.Generic.List[object]]::new()
    if ($hasOriginMaster) {
        $commands.Add(@('diff', '--name-only', '--diff-filter=AM', 'origin/master...HEAD', '--', 'Specs/TestRuns'))
    }
    $commands.Add(@('diff', '--name-only', '--diff-filter=AM', '--', 'Specs/TestRuns'))
    $commands.Add(@('diff', '--cached', '--name-only', '--diff-filter=AM', '--', 'Specs/TestRuns'))

    foreach ($arguments in $commands) {
        $output = & git -C $root @arguments 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "git $($arguments -join ' ') failed while locating changed TestRuns artifacts."
        }
        foreach ($path in $output) {
            if (-not [string]::IsNullOrWhiteSpace($path)) {
                [void]$paths.Add($path.Trim())
            }
        }
    }

    return @($paths | Sort-Object)
}

Export-ModuleMember -Function Get-RSTestRunArchiveViolations, Get-RSChangedTestRunArchivePaths
