Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-RSExistingPath {
    <#
    .SYNOPSIS
        Resolves an existing path or raises the standard Resolve-Path error.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
}

function Read-RSJsonFile {
    <#
    .SYNOPSIS
        Reads a JSON file with the depth required by archived self-test results.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return (Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -Depth 64)
}

function Get-RSTestRunArtifactCandidates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Results', 'Trace')]
        [string]$ArtifactKind,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Auto', 'FileOps', 'CompareDirectories', 'Commands', 'SelfTest')]
        [string]$Suite
    )

    $namesBySuite = if ($ArtifactKind -eq 'Results') {
        @{
            FileOps = @('fileops_results.json')
            CompareDirectories = @('compare_results.json')
            Commands = @('commands_results.json')
            SelfTest = @('selftest_run_results.json', 'results.json')
            Auto = @(
                'fileops_results.json',
                'compare_results.json',
                'commands_results.json',
                'selftest_run_results.json',
                'results.json'
            )
        }
    }
    else {
        @{
            FileOps = @('fileops_trace.txt')
            CompareDirectories = @('compare_trace.txt')
            Commands = @('commands_trace.txt')
            SelfTest = @('selftest_run_trace.txt', 'trace.txt')
            Auto = @(
                'fileops_trace.txt',
                'compare_trace.txt',
                'commands_trace.txt',
                'selftest_run_trace.txt',
                'trace.txt'
            )
        }
    }

    return @($namesBySuite[$Suite] | ForEach-Object { Join-Path $RunRoot $_ })
}

function Find-RSTestRunArtifact {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Results', 'Trace')]
        [string]$ArtifactKind,

        [ValidateSet('Auto', 'FileOps', 'CompareDirectories', 'Commands', 'SelfTest')]
        [string]$Suite = 'Auto',

        [Parameter(Mandatory = $true)]
        [ValidateSet('ReturnNull', 'Throw')]
        [string]$MissingPolicy
    )

    $candidates = @(Get-RSTestRunArtifactCandidates `
        -RunRoot $RunRoot `
        -ArtifactKind $ArtifactKind `
        -Suite $Suite)
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    if ($MissingPolicy -eq 'Throw') {
        $description = if ($ArtifactKind -eq 'Results') { 'results JSON' } else { 'trace text' }
        throw "Could not find $description in run folder: $RunRoot"
    }

    return $null
}

function Find-RSTestRunResultsJson {
    <#
    .SYNOPSIS
        Finds the selected suite results artifact in a self-test run.

    .DESCRIPTION
        Uses the repository's stable suite candidate order. MissingPolicy is required
        so callers explicitly choose whether an absent artifact skips a run or fails.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunRoot,

        [ValidateSet('Auto', 'FileOps', 'CompareDirectories', 'Commands', 'SelfTest')]
        [string]$Suite = 'Auto',

        [Parameter(Mandatory = $true)]
        [ValidateSet('ReturnNull', 'Throw')]
        [string]$MissingPolicy
    )

    return Find-RSTestRunArtifact `
        -RunRoot $RunRoot `
        -ArtifactKind Results `
        -Suite $Suite `
        -MissingPolicy $MissingPolicy
}

function Find-RSTestRunTraceFile {
    <#
    .SYNOPSIS
        Finds the selected suite trace artifact in a self-test run.

    .DESCRIPTION
        Uses the repository's stable suite candidate order. MissingPolicy is required
        so callers explicitly choose whether an absent trace is optional or fatal.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunRoot,

        [ValidateSet('Auto', 'FileOps', 'CompareDirectories', 'Commands', 'SelfTest')]
        [string]$Suite = 'Auto',

        [Parameter(Mandatory = $true)]
        [ValidateSet('ReturnNull', 'Throw')]
        [string]$MissingPolicy
    )

    return Find-RSTestRunArtifact `
        -RunRoot $RunRoot `
        -ArtifactKind Trace `
        -Suite $Suite `
        -MissingPolicy $MissingPolicy
}

function Get-RSTestRunTotalCases {
    <#
    .SYNOPSIS
        Resolves the total case count from either a result object or summary counts.

    .DESCRIPTION
        The Results parameter set prefers a non-empty cases array and otherwise sums
        passed, failed, and skipped. The Counts parameter set sums those counts when
        nonzero and otherwise uses FallbackCaseCount, preserving analyzer behavior.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Results')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'Results')]
        [object]$Results,

        [Parameter(Mandatory = $true, ParameterSetName = 'Counts')]
        [int]$Passed,

        [Parameter(Mandatory = $true, ParameterSetName = 'Counts')]
        [int]$Failed,

        [Parameter(Mandatory = $true, ParameterSetName = 'Counts')]
        [int]$Skipped,

        [Parameter(Mandatory = $true, ParameterSetName = 'Counts')]
        [int]$FallbackCaseCount
    )

    if ($PSCmdlet.ParameterSetName -eq 'Counts') {
        $total = $Passed + $Failed + $Skipped
        if ($total -gt 0) {
            return $total
        }
        return $FallbackCaseCount
    }

    if ($Results.PSObject.Properties.Match('cases').Count -gt 0 -and
        $Results.cases -and
        $Results.cases.Count -gt 0) {
        return [int]$Results.cases.Count
    }

    $passedCount = if ($null -eq $Results.passed) { 0 } else { [int]$Results.passed }
    $failedCount = if ($null -eq $Results.failed) { 0 } else { [int]$Results.failed }
    $skippedCount = if ($null -eq $Results.skipped) { 0 } else { [int]$Results.skipped }
    return ($passedCount + $failedCount + $skippedCount)
}

function Format-RSPercent {
    <#
    .SYNOPSIS
        Formats an invariant-culture percentage or an empty value for zero totals.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Numerator,

        [Parameter(Mandatory = $true)]
        [int]$Denominator,

        [int]$Decimals = 1
    )

    if ($Denominator -le 0) {
        return ''
    }

    $percent = 100.0 * $Numerator / [double]$Denominator
    $format = '0.' + ('0' * $Decimals)
    return ($percent.ToString($format, [Globalization.CultureInfo]::InvariantCulture) + '%')
}

function Format-RSDeltaPercent {
    <#
    .SYNOPSIS
        Formats an invariant-culture signed percentage change.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [Nullable[int]]$OldValue,

        [AllowNull()]
        [Nullable[int]]$NewValue,

        [int]$Decimals = 1
    )

    if ($null -eq $OldValue -or $null -eq $NewValue -or $OldValue -le 0) {
        return $null
    }

    $delta = $NewValue - $OldValue
    if ($delta -eq 0) {
        return '0%'
    }

    $percent = 100.0 * $delta / [double]$OldValue
    $format = '0.' + ('0' * $Decimals)
    $text = $percent.ToString($format, [Globalization.CultureInfo]::InvariantCulture) + '%'
    if ($percent -gt 0) {
        return "+$text"
    }
    return $text
}

function Get-RSTestRunSummaryColor {
    <#
    .SYNOPSIS
        Selects the console color for a self-test run summary.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Failed,

        [Parameter(Mandatory = $true)]
        [int]$Skipped
    )

    if ($Failed -gt 0) {
        return 'Red'
    }
    if ($Skipped -gt 0) {
        return 'Yellow'
    }
    return 'Green'
}

Export-ModuleMember -Function @(
    'Resolve-RSExistingPath',
    'Read-RSJsonFile',
    'Find-RSTestRunResultsJson',
    'Find-RSTestRunTraceFile',
    'Get-RSTestRunTotalCases',
    'Format-RSPercent',
    'Format-RSDeltaPercent',
    'Get-RSTestRunSummaryColor'
)
