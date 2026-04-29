# Windows Tool Discovery Pattern

**Scope:** PowerShell pattern for auto-discovering development tools (vcpkg, CMake, compilers, etc.) from common Windows installation locations including Visual Studio bundled tools, package managers (Chocolatey, Scoop), and standard directories.

## Overview

Windows development tools can be installed in many locations:
- **Visual Studio bundled** — Tools shipped with VS (detected via vswhere.exe)
- **Package managers** — Chocolatey, Scoop, winget
- **Common manual install locations** — C:\tools, C:\dev, %USERPROFILE%\source
- **Environment variables** — Tool-specific root paths (VCPKG_ROOT, CMAKE_ROOT)
- **PATH** — System or user PATH entries

**Priority Chain:** Explicit parameter → Repo-local → Tool-specific env var → PATH → System-wide discovery → Throw

## Pattern

### Function Structure

```powershell
function Resolve-ToolExePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,
        
        [Parameter(Mandatory = $false)]
        [string]$ExplicitToolExe
    )
    
    # 1. Explicit parameter (highest priority)
    if ($ExplicitToolExe) {
        if (-not (Test-Path $ExplicitToolExe -PathType Leaf)) {
            throw "Tool not found: $ExplicitToolExe"
        }
        Write-Verbose "Tool discovered at: $ExplicitToolExe (source: explicit parameter)"
        return (Resolve-Path $ExplicitToolExe).Path
    }
    
    # 2. Repo-local (e.g., vcpkg\vcpkg.exe in repo root)
    $repoLocal = Join-Path $RepoRoot "tool\tool.exe"
    if (Test-Path $repoLocal -PathType Leaf) {
        Write-Verbose "Tool discovered at: $repoLocal (source: repo-local)"
        return (Resolve-Path $repoLocal).Path
    }
    
    # 3. Tool-specific environment variable (e.g., VCPKG_ROOT, CMAKE_ROOT)
    if ($env:TOOL_ROOT) {
        $fromRoot = Join-Path $env:TOOL_ROOT "tool.exe"
        if (Test-Path $fromRoot -PathType Leaf) {
            Write-Verbose "Tool discovered at: $fromRoot (source: TOOL_ROOT environment variable)"
            return (Resolve-Path $fromRoot).Path
        }
    }
    
    # 4. PATH
    $cmd = Get-Command "tool.exe" -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) {
        Write-Verbose "Tool discovered at: $($cmd.Source) (source: PATH)"
        return $cmd.Source
    }
    
    # 5. Visual Studio bundled (via vswhere)
    $vsPath = Get-VisualStudioToolPath -ToolRelativePath "path\to\tool.exe" -Component "Microsoft.VisualStudio.Component.Name"
    if ($vsPath) {
        return $vsPath
    }
    
    # 6. Chocolatey
    $chocoPath1 = "C:\tools\tool\tool.exe"
    if (Test-Path $chocoPath1 -PathType Leaf) {
        Write-Verbose "Tool discovered at: $chocoPath1 (source: Chocolatey)"
        return (Resolve-Path $chocoPath1).Path
    }
    
    if ($env:ChocolateyInstall) {
        $chocoPath2 = Join-Path $env:ChocolateyInstall "lib\tool\tools\tool.exe"
        if (Test-Path $chocoPath2 -PathType Leaf) {
            Write-Verbose "Tool discovered at: $chocoPath2 (source: Chocolatey)"
            return (Resolve-Path $chocoPath2).Path
        }
    }
    
    # 7. Scoop
    $scoopPath1 = Join-Path $env:USERPROFILE "scoop\apps\tool\current\tool.exe"
    if (Test-Path $scoopPath1 -PathType Leaf) {
        Write-Verbose "Tool discovered at: $scoopPath1 (source: Scoop)"
        return (Resolve-Path $scoopPath1).Path
    }
    
    if ($env:SCOOP) {
        $scoopPath2 = Join-Path $env:SCOOP "apps\tool\current\tool.exe"
        if (Test-Path $scoopPath2 -PathType Leaf) {
            Write-Verbose "Tool discovered at: $scoopPath2 (source: Scoop)"
            return (Resolve-Path $scoopPath2).Path
        }
    }
    
    # 8. Common manual install locations
    $commonRoots = @(
        "C:\tool\tool.exe",
        "C:\dev\tool\tool.exe",
        (Join-Path $env:USERPROFILE "tool\tool.exe"),
        (Join-Path $env:USERPROFILE "source\tool\tool.exe")
    )
    
    foreach ($commonPath in $commonRoots) {
        if (Test-Path $commonPath -PathType Leaf) {
            Write-Verbose "Tool discovered at: $commonPath (source: common installation directory)"
            return (Resolve-Path $commonPath).Path
        }
    }
    
    # 9. Throw with helpful message
    throw "tool.exe not found. Install tool and add it to PATH, or set TOOL_ROOT, or pass -ToolExe. Auto-discovery checks: repo-local, TOOL_ROOT, PATH, Visual Studio bundled, Chocolatey, Scoop, and common installation directories."
}
```

### Visual Studio Tool Discovery via vswhere.exe

```powershell
function Get-VisualStudioToolPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolRelativePath,  # e.g., "VC\vcpkg\vcpkg.exe"
        
        [Parameter(Mandatory = $true)]
        [string]$Component  # e.g., "Microsoft.VisualStudio.Component.Vcpkg"
    )
    
    $vswherePath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswherePath -PathType Leaf)) {
        return $null  # vswhere not available, skip silently
    }
    
    try {
        $vsPath = & $vswherePath -latest -prerelease -products * -requires $Component -property installationPath 2>$null
        if ($vsPath) {
            $toolPath = Join-Path $vsPath $ToolRelativePath
            if (Test-Path $toolPath -PathType Leaf) {
                Write-Verbose "Tool discovered at: $toolPath (source: Visual Studio bundled)"
                return (Resolve-Path $toolPath).Path
            }
        }
    } catch {
        # vswhere exists but failed; skip silently
    }
    
    return $null
}
```

**vswhere.exe flags:**
- `-latest` — Return only the newest version
- `-prerelease` — Include preview/beta VS versions
- `-products *` — Include all VS products (Community, Professional, Enterprise)
- `-requires <component>` — Filter to VS installs with specified component
- `-property installationPath` — Return the root path (e.g., "C:\Program Files\Microsoft Visual Studio\2022\Community")

**Common VS Components:**
- `Microsoft.VisualStudio.Component.Vcpkg` — vcpkg package manager
- `Microsoft.VisualStudio.Component.VC.CMake.Project` — CMake tools
- `Microsoft.VisualStudio.Component.VC.Tools.x86.x64` — MSVC compiler

## Key Principles

1. **Explicit wins** — Always honor explicit parameter values first
2. **Repo-local second** — Check repo-local before system paths (predictable, versioned with repo)
3. **PATH beats system discovery** — If user added tool to PATH, that's their preferred version
4. **Silent fallthrough** — When env vars are unset or paths missing, continue silently to next check
5. **PathType Leaf** — Always use `Test-Path -PathType Leaf` for executable checks (ensures it's a file, not directory)
6. **Verbose diagnostics** — Emit `Write-Verbose` with source name when a path succeeds
7. **Helpful error** — Final throw should list all discovery methods attempted

## Chocolatey Discovery

Chocolatey has two common install patterns:

```powershell
# 1. Default global tools location
"C:\tools\<package>\<package>.exe"

# 2. ChocolateyInstall lib directory
"${env:ChocolateyInstall}\lib\<package>\tools\<package>.exe"
```

**Example (vcpkg):**
- `C:\tools\vcpkg\vcpkg.exe`
- `C:\ProgramData\chocolatey\lib\vcpkg\tools\vcpkg.exe` (if `$env:ChocolateyInstall = "C:\ProgramData\chocolatey"`)

## Scoop Discovery

Scoop uses versioned directories with `current` symlink:

```powershell
# 1. Default user scoop location
"${env:USERPROFILE}\scoop\apps\<app>\current\<app>.exe"

# 2. Custom SCOOP env var location
"${env:SCOOP}\apps\<app>\current\<app>.exe"
```

**Example (vcpkg):**
- `C:\Users\username\scoop\apps\vcpkg\current\vcpkg.exe`
- `D:\scoop\apps\vcpkg\current\vcpkg.exe` (if `$env:SCOOP = "D:\scoop"`)

## Common Manual Install Locations

Windows developers often clone/install to:
- `C:\<tool>\` — Root drive convention (Microsoft docs recommend C:\vcpkg for vcpkg)
- `C:\dev\<tool>\` — Developer tools directory
- `${env:USERPROFILE}\<tool>\` — User-local install (no admin rights needed)
- `${env:USERPROFILE}\source\<tool>\` — VS default "source" directory

## Usage in vcpkg-install.ps1

**Reference implementation:** Z:\src\RedSalamander\vcpkg-install.ps1 `Resolve-VcpkgExePath` function

**Priority chain:**
1. `-VcpkgExe` parameter
2. `vcpkg\vcpkg.exe` (repo-local)
3. `$env:VCPKG_ROOT`
4. PATH (`Get-Command vcpkg.exe`)
5. Visual Studio bundled (`vswhere.exe -requires Microsoft.VisualStudio.Component.Vcpkg`)
6. Chocolatey (`C:\tools\vcpkg\`, `${env:ChocolateyInstall}\lib\vcpkg\tools\`)
7. Scoop (`scoop\apps\vcpkg\current\`)
8. Common roots (`C:\vcpkg\`, `C:\dev\vcpkg\`, `%USERPROFILE%\vcpkg\`, `%USERPROFILE%\source\vcpkg\`)

## When to Use This Pattern

**Use for:**
- Development tools with multiple common installation methods (vcpkg, CMake, ninja, compilers)
- Tools that may be bundled with Visual Studio
- Tools commonly installed via Chocolatey or Scoop on Windows
- Scripts intended for broad Windows developer audience (not CI-only)

**Don't use for:**
- CI/automation scripts — Prefer explicit env vars or parameters for reproducibility
- Tools with single canonical install location
- Non-Windows platforms (Chocolatey/Scoop/vswhere are Windows-specific)

## Testing

```powershell
# Test with -Verbose to see discovery path
.\tool-script.ps1 -Verbose

# Test explicit override
.\tool-script.ps1 -ToolExe "C:\custom\path\tool.exe"

# Test error message when tool not found
# (rename tool.exe temporarily to verify helpful error)
```

## Related

- **vswhere.exe documentation:** https://github.com/microsoft/vswhere/wiki
- **Chocolatey install locations:** https://docs.chocolatey.org/en-us/choco/commands/install
- **Scoop app directory structure:** https://github.com/ScoopInstaller/Scoop/wiki/Apps
