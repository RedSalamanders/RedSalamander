# Private packaging module; import through its reviewed callers.
Set-StrictMode -Version Latest

$runtimeDependencyModule = Join-Path $PSScriptRoot 'RuntimeDependencies.psm1'
$sanitizedEnvironmentModule = Join-Path $PSScriptRoot '..\Build\SanitizedEnvironment.psm1'
Import-Module $runtimeDependencyModule -ErrorAction Stop
Import-Module $sanitizedEnvironmentModule -ErrorAction Stop

function Invoke-RSPortableSmokeProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$ArgumentList = @()
    )

    $exitCode = Invoke-RSProcess `
        -FilePath $FilePath `
        -Arguments $ArgumentList `
        -WorkingDirectory (Get-Location).Path

    if ($exitCode -ne 0) {
        throw "Portable-package smoke process failed with exit code $exitCode`: $FilePath $($ArgumentList -join ' ')"
    }
}

function Test-RSPortableGhosttyRuntime {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$ExtractRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [string]$ReferenceRuntimeRoot,

        [switch]$NativeExecutionSkipped
    )

    $lockPath = Join-Path $RepoRoot 'External\TerminalEngine\GhosttyRuntimeLock.v1.json'
    $packageRuntimeRoot = Join-Path $ExtractRoot 'Plugins\TerminalRuntime'
    $packageDllPath = Join-Path $packageRuntimeRoot 'ghostty-vt.dll'
    $packageIdentityPath = Join-Path $packageRuntimeRoot 'terminal-runtime-identity.json'
    $referenceDllPath = Join-Path $ReferenceRuntimeRoot 'ghostty-vt.dll'
    $referenceIdentityPath = Join-Path $ReferenceRuntimeRoot 'terminal-runtime-identity.json'
    foreach ($requiredPath in @($lockPath, $packageDllPath, $packageIdentityPath, $referenceDllPath, $referenceIdentityPath)) {
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            throw "Portable Ghostty runtime evidence is missing: $requiredPath"
        }
    }

    $packageIdentitySha256 = (Get-FileHash -LiteralPath $packageIdentityPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $referenceIdentitySha256 = (Get-FileHash -LiteralPath $referenceIdentityPath -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($packageIdentitySha256 -cne $referenceIdentitySha256) {
        throw 'Portable Ghostty runtime identity differs from the receipt-bound shipping-profile identity.'
    }

    $identity = Get-Content -LiteralPath $packageIdentityPath -Raw | ConvertFrom-Json
    $lock = Get-Content -LiteralPath $lockPath -Raw | ConvertFrom-Json
    $expectedMachine = if ($Platform -eq 'ARM64') { 'ARM64' } else { 'AMD64' }
    $expectedImports = [string[]]@($lock.abi.importModules)
    $actualImports = [string[]]@($identity.importModules)
    if ([string]$identity.formatId -cne 'red-salamander-ghostty-runtime' -or
        [int]$identity.schemaVersion -ne 5 -or
        [string]$identity.ghosttyCommit -cne [string]$lock.upstream.commit -or
        [string]$identity.machine -cne $expectedMachine -or
        [int]$identity.exportCount -ne [int]$lock.abi.exportCount -or
        [string]$identity.exportSetSha256 -cne [string]$lock.abi.exportSetSha256 -or
        ($actualImports -join "`0") -cne ($expectedImports -join "`0")) {
        throw "Portable Ghostty runtime identity does not match the canonical $Platform lock contract."
    }

    $packageDllSha256 = (Get-FileHash -LiteralPath $packageDllPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $referenceDllSha256 = (Get-FileHash -LiteralPath $referenceDllPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $packageDllSize = (Get-Item -LiteralPath $packageDllPath).Length
    if ($packageDllSha256 -cne [string]$identity.dllSha256 -or
        $packageDllSha256 -cne $referenceDllSha256 -or
        $packageDllSize -ne [long]$identity.dllSizeBytes) {
        throw 'Portable Ghostty runtime DLL does not match its receipt-bound raw identity.'
    }

    foreach ($notice in @($lock.notices)) {
        $licenseName = Split-Path -Leaf ([string]$notice.outputRelativePath)
        $packageLicensePath = Join-Path $packageRuntimeRoot $licenseName
        $referenceLicensePath = Join-Path $ReferenceRuntimeRoot $licenseName
        foreach ($licensePath in @($packageLicensePath, $referenceLicensePath)) {
            if (-not (Test-Path -LiteralPath $licensePath -PathType Leaf)) {
                throw "Portable Ghostty notice is missing: $licensePath"
            }
        }
        $packageLicenseSha256 = (Get-FileHash -LiteralPath $packageLicensePath -Algorithm SHA256).Hash.ToLowerInvariant()
        $referenceLicenseSha256 = (Get-FileHash -LiteralPath $referenceLicensePath -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($packageLicenseSha256 -cne [string]$notice.normalizedTextSha256 -or
            $packageLicenseSha256 -cne $referenceLicenseSha256) {
            throw "Portable Ghostty notice '$licenseName' does not match the receipt-bound shipping-profile notice."
        }
    }

    return [pscustomobject]@{
        DllSha256 = $packageDllSha256
        IdentitySha256 = $packageIdentitySha256
        Machine = [string]$identity.machine
        RuntimeLoadSkipped = [bool]$NativeExecutionSkipped
    }
}

function Test-RSPortablePackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$ZipPath,

        [Parameter(Mandatory = $true)]
        [string]$BuildOutputDir,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Debug', 'Release')]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [string]$ScratchRoot = (Join-Path $env:TEMP 'RedSalamanderPackageSmoke'),

        [string]$ReferenceRuntimeRoot,

        [switch]$SkipExecution
    )

    if (-not (Test-Path -LiteralPath $ZipPath -PathType Leaf)) {
        throw "Portable package not found: $ZipPath"
    }
    $ZipPath = (Resolve-Path -LiteralPath $ZipPath -ErrorAction Stop).Path
    $BuildOutputDir = (Resolve-Path -LiteralPath $BuildOutputDir -ErrorAction Stop).Path

    $pluginContractSource = Join-Path $BuildOutputDir 'PluginContractTests.exe'
    if (-not $SkipExecution -and -not (Test-Path -LiteralPath $pluginContractSource -PathType Leaf)) {
        throw "Plugin contract harness not found: $pluginContractSource"
    }

    New-Item -ItemType Directory -Path $ScratchRoot -Force | Out-Null
    $extractRoot = Join-Path $ScratchRoot ([Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $extractRoot | Out-Null

    try {
        Expand-Archive -LiteralPath $ZipPath -DestinationPath $extractRoot

        $requiredPackageFiles = @(
            'RedLauncher.exe',
            'red.exe',
            'RedSalamander.exe',
            'RedSalamanderMonitor.exe',
            'Plugins\FileSystem.dll',
            'Plugins\FileSystem7z.dll',
            'Plugins\FileSystemCurl.dll',
            'Plugins\FileSystemDummy.dll',
            'Plugins\FileSystemGoogleDrive.dll',
            'Plugins\FileSystemMicrosoftDrive.dll',
            'Plugins\FileSystemS3.dll',
            'Plugins\ViewerText.dll',
            'Plugins\ViewerSqlite.dll',
            'Plugins\ViewerSpace.dll',
            'Plugins\ViewerImgRaw.dll',
            'Plugins\ViewerVLC.dll',
            'Plugins\ViewerPE.dll',
            'Plugins\ViewerWeb.dll',
            'Plugins\Terminal.dll'
        )
        $missingPackageFiles = @($requiredPackageFiles |
            ForEach-Object { Join-Path $extractRoot $_ } |
            Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
        if ($missingPackageFiles.Count -gt 0) {
            throw "Portable package is missing required files: $($missingPackageFiles -join ', ')"
        }

        $null = Assert-RSRuntimeDependenciesInOutput `
            -RepoRoot $RepoRoot `
            -BuildOutputDir $extractRoot `
            -Configuration $Configuration `
            -Platform $Platform

        if ([string]::IsNullOrWhiteSpace($ReferenceRuntimeRoot)) {
            $ReferenceRuntimeRoot = Join-Path $BuildOutputDir 'Plugins\TerminalRuntime'
        }
        $runtimeEvidence = Test-RSPortableGhosttyRuntime `
            -RepoRoot $RepoRoot `
            -ExtractRoot $extractRoot `
            -Platform $Platform `
            -ReferenceRuntimeRoot $ReferenceRuntimeRoot `
            -NativeExecutionSkipped:$SkipExecution

        if (-not $SkipExecution) {
            Push-Location $extractRoot
            try {
                Invoke-RSPortableSmokeProcess -FilePath (Join-Path $extractRoot 'RedSalamander.exe') -ArgumentList @('--help')

                $pluginContractPath = Join-Path $extractRoot 'PluginContractTests.exe'
                Copy-Item -LiteralPath $pluginContractSource -Destination $pluginContractPath
                Invoke-RSPortableSmokeProcess -FilePath $pluginContractPath -ArgumentList @('--package-smoke')
            } finally {
                Pop-Location
            }
        }

        return [pscustomobject]@{
            ExtractRoot = $extractRoot
            ExecutionSkipped = [bool]$SkipExecution
            RequiredFileCount = $requiredPackageFiles.Count
            GhosttyRuntime = $runtimeEvidence
        }
    } finally {
        if (Test-Path -LiteralPath $extractRoot) {
            Remove-Item -LiteralPath $extractRoot -Recurse -Force
        }
    }
}

Export-ModuleMember -Function Test-RSPortablePackage
