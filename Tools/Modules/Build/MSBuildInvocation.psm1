# Private build module; import through its reviewed callers.
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

function Get-RSMSBuildDiagnosticKind {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Line
    )

    # MSBuild's canonical header permits a blank origin, a subcategory (for
    # example "Command line" or "fatal"), and an optional non-space code.
    # Anchor one header so diagnostic-looking text in its message is not counted
    # again. Prefer the origin-less form before consuming an origin delimiter.
    # The optional node prefix and drive colon are console/Windows path syntax.
    # Category tokens are locale-neutral even when the diagnostic text is not.
    $diagnosticMatch = [regex]::Match(
        $Line,
        '^\s*(?:\d+>)?\s*(?:(?:[a-z]:)?[^:\r\n]*:\s*)??(?:[^:\r\n]*?\s+)??(?<kind>warning|error)(?:\s+[^:\s]+)?\s*:',
        [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant)
    if ($diagnosticMatch.Success) {
        return $diagnosticMatch.Groups['kind'].Value.ToLowerInvariant()
    }

    return $null
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
        switch (Get-RSMSBuildDiagnosticKind -Line $line) {
            'warning' { $warningCount++ }
            'error' { $errorCount++ }
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

    switch (Get-RSMSBuildDiagnosticKind -Line $Line) {
        'error' { return 'Red' }
        'warning' { return 'Yellow' }
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

function Get-RSHostOrderedMSBuildRelativePaths {
    # Bin\MSBuild.exe is the 32-bit x86 MSBuild. The VC toolset derives the compiler host from the
    # MSBuild process architecture and downgrades arm64/x64 hosts it cannot see, so every discovery
    # strategy prefers the native 64-bit MSBuild of this host, then every other 64-bit executable
    # the host can run (ARM64 Windows emulates x64; x64 Windows cannot run arm64\), and only then
    # an x86 executable.
    $isArm64Host = [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64'
    [string[]]$candidates = if ($isArm64Host) {
        @('MSBuild\Current\Bin\arm64\MSBuild.exe', 'MSBuild\Current\Bin\amd64\MSBuild.exe')
    } else {
        @('MSBuild\Current\Bin\amd64\MSBuild.exe')
    }
    return $candidates + [string[]]@(
        'MSBuild\15.0\Bin\amd64\MSBuild.exe',
        'MSBuild\Current\Bin\MSBuild.exe',
        'MSBuild\15.0\Bin\MSBuild.exe'
    )
}

function Resolve-RSHostOrderedMSBuildPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    # A PATH hit from a developer prompt is usually the x86 MSBuild\Current\Bin\MSBuild.exe; map
    # it to the best host-ordered executable of the same installation when one exists.
    $match = [regex]::Match($Path, '^(?<root>.+?)\\MSBuild\\(?:Current|15\.0)\\Bin\\(?:(?:amd64|arm64)\\)?MSBuild\.exe$', 'IgnoreCase')
    if (-not $match.Success) { return $Path }
    foreach ($relative in Get-RSHostOrderedMSBuildRelativePaths) {
        $candidate = Join-Path $match.Groups['root'].Value $relative
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $Path
}

function Find-RSHostOrderedMSBuildInInstallation {
    param([Parameter(Mandatory = $true)][string]$InstallationPath)

    foreach ($relative in Get-RSHostOrderedMSBuildRelativePaths) {
        $candidate = Join-Path $InstallationPath $relative
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    return $null
}

function Get-RSMSBuildFileVersion {
    param([Parameter(Mandatory)][string]$Path)
    return [Diagnostics.FileVersionInfo]::GetVersionInfo($Path)
}

function Get-RSWorkflowSelectedMSBuild {
    # Workflow selection precedes discovery. An explicit selection is authoritative
    # and fails closed; a PATH-only compatibility candidate must pass the same
    # binary/layout checks after applying host ordering within its own installation.
    if ($env:GITHUB_ACTIONS -ne 'true') { return $null }
    $explicit = -not [string]::IsNullOrWhiteSpace($env:MSBUILD_EXE_PATH)
    if ($explicit) {
        $candidate = [IO.Path]::GetFullPath($env:MSBUILD_EXE_PATH)
    } else {
        $command = Get-Command msbuild.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $command -or -not $command.Source) { return $null }
        $candidate = Resolve-RSHostOrderedMSBuildPath -Path $command.Source
    }
    $layout = [regex]::Match($candidate, '^(?<root>.+?)\\MSBuild\\Current\\Bin\\(?:(?:amd64|arm64)\\)?MSBuild\.exe$', 'IgnoreCase')
    $version = $null
    if ($layout.Success -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        try { $version = Get-RSMSBuildFileVersion -Path $candidate }
        catch { if ($explicit) { throw "Cannot identify workflow-selected MSBuild: $candidate. $($_.Exception.Message)" } }
    }
    if (-not $layout.Success -or $null -eq $version -or $version.FileMajorPart -lt 18 -or
        $version.OriginalFilename -ine 'MSBuild.exe') {
        if ($explicit) { throw "Workflow-selected MSBuild is not an existing Visual Studio 2026 executable: $candidate" }
        return $null
    }
    $method = if ($explicit) { 'MSBUILD_EXE_PATH' } else { 'PATH' }
    return @{
        Path = $candidate
        InstallationPath = $layout.Groups['root'].Value
        Version = "MSBuild $($version.FileMajorPart) ($method)"
        Method = $method
    }
}

Export-ModuleMember -Function @(
    'Test-RSInteractiveTerminal',
    'Get-RSMSBuildInvocationPlan',
    'Get-RSMSBuildDiagnosticSummary',
    'Write-RSMSBuildDiagnosticSummary',
    'Get-RSMSBuildLineForegroundColor',
    'Write-RSMSBuildStreamingLine',
    'Get-RSHostOrderedMSBuildRelativePaths',
    'Resolve-RSHostOrderedMSBuildPath',
    'Find-RSHostOrderedMSBuildInInstallation',
    'Get-RSWorkflowSelectedMSBuild'
)
