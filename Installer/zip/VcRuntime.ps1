Set-StrictMode -Version Latest

$script:RSVcRuntimeRequiredDlls = @(
    'msvcp140.dll',
    'msvcp140_atomic_wait.dll',
    'vcruntime140.dll',
    'vcruntime140_1.dll'
)

function Get-RSVcRuntimeRedistDirectory {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [string[]]$SearchRoots
    )

    $arch = if ($Platform -eq 'ARM64') { 'arm64' } else { 'x64' }

    if (-not $SearchRoots -or $SearchRoots.Count -eq 0) {
        $SearchRoots = @(
            Join-Path $env:ProgramFiles 'Microsoft Visual Studio'
            Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio'
        ) | Where-Object { $_ -and (Test-Path $_) }
    }

    $candidates = foreach ($root in $SearchRoots) {
        if (-not (Test-Path $root)) {
            continue
        }

        Get-ChildItem -Path $root -Directory -Recurse -Filter 'Microsoft.VC*.CRT' -ErrorAction SilentlyContinue |
            Where-Object {
                $_.FullName -match '\\VC\\Redist\\MSVC\\[^\\]+\\' + [regex]::Escape($arch) + '\\Microsoft\.VC[^\\]+\.CRT$'
            } |
            ForEach-Object {
                $versionText = $_.Parent.Parent.Name
                $version = $null
                if (-not [version]::TryParse($versionText, [ref]$version)) {
                    $version = [version]'0.0'
                }

                [pscustomobject]@{
                    Directory = $_.FullName
                    Version = $version
                }
            }
    }

    $best = $candidates | Sort-Object -Property Version, Directory -Descending | Select-Object -First 1
    if (-not $best) {
        throw "MSVC runtime redistributable directory not found for $Platform. Install the Visual C++ redistributable components with Visual Studio Build Tools."
    }

    return $best.Directory
}

function Copy-RSVcRuntimeDependencies {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [Parameter(Mandatory)]
        [string]$DestinationDir,

        [string[]]$SearchRoots
    )

    if (-not (Test-Path $DestinationDir)) {
        New-Item -ItemType Directory -Path $DestinationDir -Force | Out-Null
    }

    $redistDir = Get-RSVcRuntimeRedistDirectory -Platform $Platform -SearchRoots $SearchRoots

    foreach ($dllName in $script:RSVcRuntimeRequiredDlls) {
        $dllPath = Join-Path $redistDir $dllName
        if (-not (Test-Path $dllPath)) {
            throw "Required MSVC runtime DLL '$dllName' was not found in $redistDir."
        }
    }

    $copied = @()
    foreach ($dll in Get-ChildItem -Path $redistDir -Filter '*.dll' -File) {
        Copy-Item -LiteralPath $dll.FullName -Destination $DestinationDir -Force
        $copied += $dll.Name
    }

    return $copied
}
