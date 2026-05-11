Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Convert-RSWingetVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VersionText
    )

    $normalized = $VersionText.Trim() -replace '^[vV]', ''
    $parts = $normalized -split '\.'
    if ($parts.Count -lt 2) {
        return $null
    }

    try {
        return [version]::new([int]$parts[0], [int]$parts[1])
    }
    catch {
        return $null
    }
}

function Test-RSLegacyWingetSchemaHeaderWarning {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WingetVersion,

        [Parameter(Mandatory = $true)]
        [string[]]$OutputLines
    )

    $parsedVersion = Convert-RSWingetVersion -VersionText $WingetVersion
    if (-not $parsedVersion -or $parsedVersion -ge [version]'1.12') {
        return $false
    }

    $hasSucceededWithWarnings = $OutputLines | Where-Object {
        $_ -match '^Manifest validation succeeded with warnings\.$'
    } | Select-Object -First 1
    if (-not $hasSucceededWithWarnings) {
        return $false
    }

    $warningLines = @($OutputLines | Where-Object { $_ -match '^Manifest Warning:' })
    if ($warningLines.Count -eq 0) {
        return $false
    }

    foreach ($warning in $warningLines) {
        if ($warning -notmatch 'The schema header URL does not match the expected pattern') {
            return $false
        }

        if ($warning -notmatch 'winget-manifest\.(installer|defaultLocale|version)\.1\.12\.0\.schema\.json') {
            return $false
        }
    }

    return $true
}

function Invoke-RSWingetManifestValidation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,

        [string]$WingetCommand = 'winget'
    )

    if (-not (Test-Path $ManifestPath)) {
        throw "Winget manifest path not found: $ManifestPath"
    }

    if (-not (Get-Command -Name $WingetCommand -ErrorAction SilentlyContinue)) {
        throw "winget executable not found: $WingetCommand"
    }

    Write-Host "winget.exe version:"
    $wingetVersion = (& $WingetCommand --version).Trim()
    Write-Host $wingetVersion

    $hasNativeErrorPreference = Test-Path variable:PSNativeCommandUseErrorActionPreference
    $oldNativeErrorPreference = $null
    try {
        if ($hasNativeErrorPreference) {
            $oldNativeErrorPreference = $PSNativeCommandUseErrorActionPreference
            $PSNativeCommandUseErrorActionPreference = $false
        }

        $validationOutput = @(& $WingetCommand validate --manifest $ManifestPath 2>&1 | ForEach-Object { $_.ToString() })
        $validationExitCode = $LASTEXITCODE
    }
    finally {
        if ($hasNativeErrorPreference) {
            $PSNativeCommandUseErrorActionPreference = $oldNativeErrorPreference
        }
    }

    $validationOutput | ForEach-Object { Write-Host $_ }

    if ($validationExitCode -eq 0) {
        return
    }

    if (Test-RSLegacyWingetSchemaHeaderWarning -WingetVersion $wingetVersion -OutputLines $validationOutput) {
        Write-Warning "Ignoring winget.exe $wingetVersion schema-header warning for ManifestVersion 1.12.0. Newer winget validators accept these headers."
        return
    }

    throw "winget validate failed with exit code $validationExitCode."
}
