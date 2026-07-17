Set-StrictMode -Version Latest

$runtimeDependencyModule = Join-Path $PSScriptRoot 'RuntimeDependencies.psm1'
Import-Module $runtimeDependencyModule

function Invoke-RSPortableSmokeProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$ArgumentList = @()
    )

    $oldNativeErrorPreference = $PSNativeCommandUseErrorActionPreference
    try {
        $PSNativeCommandUseErrorActionPreference = $false
        & $FilePath @ArgumentList | Out-Host
        $exitCode = $LASTEXITCODE
    } finally {
        $PSNativeCommandUseErrorActionPreference = $oldNativeErrorPreference
    }

    if ($exitCode -ne 0) {
        throw "Portable-package smoke process failed with exit code $exitCode`: $FilePath $($ArgumentList -join ' ')"
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

        [switch]$SkipExecution
    )

    if (-not (Test-Path -LiteralPath $ZipPath -PathType Leaf)) {
        throw "Portable package not found: $ZipPath"
    }

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
            'Plugins\ViewerWeb.dll'
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
        }
    } finally {
        if (Test-Path -LiteralPath $extractRoot) {
            Remove-Item -LiteralPath $extractRoot -Recurse -Force
        }
    }
}

Export-ModuleMember -Function Test-RSPortablePackage
