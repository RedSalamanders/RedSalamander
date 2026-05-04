Set-StrictMode -Version Latest

function Test-RSInteractiveTerminal {
    param(
        [bool]$IsOutputRedirected = [Console]::IsOutputRedirected,
        [bool]$IsErrorRedirected = [Console]::IsErrorRedirected,
        [bool]$HasRawUi = ($null -ne $Host -and $null -ne $Host.UI -and $null -ne $Host.UI.RawUI),
        [bool]$CanReadWindowTitle = $true,
        [hashtable]$Environment = @{}
    )

    if ($Environment.Count -eq 0) {
        foreach ($entry in Get-ChildItem Env:) {
            $Environment[$entry.Name] = $entry.Value
        }
    }

    if (-not $HasRawUi) {
        return $false
    }

    if (-not $CanReadWindowTitle) {
        return $false
    }

    $interactiveHostMarkers = @(
        'CODEX_SHELL',
        'WT_SESSION',
        'TERM_PROGRAM',
        'VSCODE_PID',
        'ConEmuPID',
        'ANSICON'
    )

    $hasInteractiveMarker = $false
    foreach ($marker in $interactiveHostMarkers) {
        if ($Environment.ContainsKey($marker) -and -not [string]::IsNullOrWhiteSpace([string]$Environment[$marker])) {
            $hasInteractiveMarker = $true
            break
        }
    }

    if (-not $IsOutputRedirected -and -not $IsErrorRedirected) {
        return $true
    }

    if ($hasInteractiveMarker) {
        return $true
    }

    return $false
}

function Test-RSDirectConsoleSupportedHost {
    param(
        [hashtable]$Environment = @{}
    )

    if ($Environment.Count -eq 0) {
        foreach ($entry in Get-ChildItem Env:) {
            $Environment[$entry.Name] = $entry.Value
        }
    }

    # These terminal hosts surface MSBuild progress reliably through the replay helper, not
    # through the direct child-console path. Keep them streaming so progress stays visible.
    if ($Environment.ContainsKey('CODEX_SHELL') -and -not [string]::IsNullOrWhiteSpace([string]$Environment['CODEX_SHELL'])) {
        return $false
    }

    if ($Environment.ContainsKey('WT_SESSION') -and -not [string]::IsNullOrWhiteSpace([string]$Environment['WT_SESSION'])) {
        return $false
    }

    return $true
}

function Get-RSMSBuildFileLoggerArguments {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogPath,

        [ValidateSet('minimal', 'normal', 'detailed', 'diagnostic')]
        [string]$Verbosity = 'minimal'
    )

    $resolvedLogPath = [System.IO.Path]::GetFullPath($LogPath)
    return @(
        '/fl',
        "/flp:Verbosity=$Verbosity;LogFile=$resolvedLogPath;Encoding=UTF-8"
    )
}

function Get-RSMSBuildInvocationPlan {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$UseInteractiveTerminal,

        [Parameter(Mandatory = $true)]
        [string]$LogPath,

        [hashtable]$Environment = @{}
    )

    $useDirectConsole = $UseInteractiveTerminal -and (Test-RSDirectConsoleSupportedHost -Environment $Environment)

    return [pscustomobject]@{
        UseDirectConsole    = $useDirectConsole
        AdditionalArguments = if ($useDirectConsole) {
            Get-RSMSBuildFileLoggerArguments -LogPath $LogPath -Verbosity 'minimal'
        }
        else {
            @()
        }
    }
}

function Test-RSMSBuildDiagnosticLine {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Line,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Warning', 'Error')]
        [string]$Kind
    )

    $diagnosticCodePattern = '[A-Z]+[A-Z0-9]*\d[A-Z0-9]*'
    if ($Kind -eq 'Warning') {
        return $Line -match "(?i):\s+warning\s+$diagnosticCodePattern\s*:"
    }

    return $Line -match "(?i):\s+(?:fatal\s+)?error\s+$diagnosticCodePattern\s*:"
}

function Get-RSMSBuildDiagnosticSummary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    $resolvedLogPath = [System.IO.Path]::GetFullPath($LogPath)
    $warningCount = 0
    $errorCount = 0

    foreach ($line in [System.IO.File]::ReadLines($resolvedLogPath)) {
        if (Test-RSMSBuildDiagnosticLine -Line $line -Kind 'Warning') {
            $warningCount++
        }
        if (Test-RSMSBuildDiagnosticLine -Line $line -Kind 'Error') {
            $errorCount++
        }
    }

    return [pscustomobject]@{
        LogPath      = $resolvedLogPath
        WarningCount = $warningCount
        ErrorCount   = $errorCount
    }
}

function Write-RSMSBuildDiagnosticSummary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )

    if (-not (Test-Path -LiteralPath $LogPath)) {
        Write-Host "Diagnostics: unavailable (log not found)" -ForegroundColor DarkYellow
        return
    }

    $summary = Get-RSMSBuildDiagnosticSummary -LogPath $LogPath
    $foregroundColor = if ($summary.ErrorCount -gt 0) {
        'Red'
    }
    elseif ($summary.WarningCount -gt 0) {
        'Yellow'
    }
    else {
        'Green'
    }

    Write-Host ("Diagnostics: {0} warning(s), {1} error(s)" -f $summary.WarningCount, $summary.ErrorCount) -ForegroundColor $foregroundColor
}

function Get-RSMSBuildLineForegroundColor {
    param(
        [AllowNull()]
        [string]$Line,

        [bool]$IsError = $false
    )

    if ($IsError) {
        return 'Red'
    }

    if ([string]::IsNullOrWhiteSpace($Line)) {
        return $null
    }

    if ($Line -match '(?i)(^|\s)(fatal\s+error|error\s+([A-Z]+)?\d+|MSBUILD\s*:\s*error)\b') {
        return 'Red'
    }

    if ($Line -match '(?i)(^|\s)warning\s+([A-Z]+)?\d+\b') {
        return 'Yellow'
    }

    if ($Line -match '(?i)\.(vcxproj|vcproj|sln)\s+->\s+') {
        return 'Green'
    }

    return $null
}

function Write-RSMSBuildStreamingLine {
    param(
        [AllowNull()]
        [string]$Line,

        [bool]$IsError = $false
    )

    $foregroundColor = Get-RSMSBuildLineForegroundColor -Line $Line -IsError $IsError
    if ([string]::IsNullOrWhiteSpace($foregroundColor)) {
        Write-Host $Line
    }
    else {
        Write-Host $Line -ForegroundColor $foregroundColor
    }
}
