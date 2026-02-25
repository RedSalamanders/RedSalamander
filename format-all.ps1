function Get-ClangFormatVersion([string]$exePath) {
    if ([string]::IsNullOrWhiteSpace($exePath) -or -not (Test-Path $exePath)) {
        return $null
    }

    $output = $null
    try {
        $output = & $exePath --version 2>$null
    } catch {
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($output)) {
        return $null
    }

    if ($output -match '(?<ver>\d+\.\d+\.\d+)') {
        return [version]$Matches['ver']
    }

    return $null
}

function Get-LatestVisualStudioInstallPath() {
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere)) {
        return $null
    }

    $path = (& $vswhere -latest -prerelease -products * -property installationPath 2>$null | Select-Object -First 1)
    if (-not [string]::IsNullOrWhiteSpace($path)) {
        return $path.Trim()
    }

    $path = (& $vswhere -latest -products * -property installationPath 2>$null | Select-Object -First 1)
    if (-not [string]::IsNullOrWhiteSpace($path)) {
        return $path.Trim()
    }

    return $null
}

function Get-VisualStudioClangFormatPath([string]$vsInstallPath) {
    if ([string]::IsNullOrWhiteSpace($vsInstallPath)) {
        return $null
    }

    $candidates = @(
        "${vsInstallPath}\VC\Tools\Llvm\x64\bin\clang-format.exe",
        "${vsInstallPath}\VC\Tools\Llvm\bin\clang-format.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Resolve-ClangFormat() {
    $llvmProgramFiles = "${env:ProgramFiles}\LLVM\bin\clang-format.exe"
    $pfVersion = Get-ClangFormatVersion $llvmProgramFiles

    $vsPath = Get-VisualStudioClangFormatPath (Get-LatestVisualStudioInstallPath)
    $vsVersion = Get-ClangFormatVersion $vsPath

    if ($pfVersion -and $vsVersion) {
        if ($pfVersion -gt $vsVersion) {
            return @{
                Path = $llvmProgramFiles
                Version = $pfVersion
                Source = "Program Files LLVM"
            }
        }

        return @{
            Path = $vsPath
            Version = $vsVersion
            Source = "Visual Studio"
        }
    }

    if ($pfVersion) {
        return @{
            Path = $llvmProgramFiles
            Version = $pfVersion
            Source = "Program Files LLVM"
        }
    }

    if ($vsVersion) {
        return @{
            Path = $vsPath
            Version = $vsVersion
            Source = "Visual Studio"
        }
    }

    $fromPath = (Get-Command clang-format -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($fromPath -and (Test-Path $fromPath.Source)) {
        $pathVersion = Get-ClangFormatVersion $fromPath.Source
        return @{
            Path = $fromPath.Source
            Version = $pathVersion
            Source = "PATH"
        }
    }

    return $null
}

$clangFormat = Resolve-ClangFormat
if (-not $clangFormat) {
    throw "clang-format.exe not found. Install LLVM or the Visual Studio LLVM/Clang tools, or add clang-format to PATH."
}

$versionText = "unknown"
if ($clangFormat.Version) {
    $versionText = $clangFormat.Version.ToString()
}
Write-Host "Using clang-format ($($clangFormat.Source)): $($clangFormat.Path) (v$versionText)" -ForegroundColor Yellow

# Format all C++ files, excluding build/package directories
Get-ChildItem -Recurse -Include *.cpp,*.h |
    Where-Object {
        $_.FullName -notmatch '\\(packages|vcpkg_installed|\.vs|\.build|Debug|Release|x64|out|\.git)\\'
    } |
    ForEach-Object {
        Write-Host "Formatting: $($_.Name)" -ForegroundColor Cyan
        & $clangFormat.Path -i --style=file $_.FullName

        if ($LASTEXITCODE -eq 0) {
            Write-Host "  ✓ Success" -ForegroundColor Green
        } else {
            Write-Host "  ✗ Failed" -ForegroundColor Red
        }
    }

Write-Host "`nFormatting complete!" -ForegroundColor Yellow
Write-Host "Review changes with: git diff" -ForegroundColor Yellow
