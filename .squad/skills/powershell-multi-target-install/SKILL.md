# PowerShell Multi-Target Installation Pattern

**Skill:** PowerShell parameter patterns for installing multiple targets (platforms, configurations, triplets) with smart defaults.

**Use When:** Writing PowerShell scripts that support multiple build targets, platforms, or configurations where users want "install everything" by default but also need fine-grained control.

## Core Pattern

### Problem
Users want:
1. **Default behavior:** One command installs everything (all platforms, all configurations)
2. **Selective control:** Narrow to specific platform or configuration when needed
3. **Explicit override:** Bypass all logic and specify exact target

### Solution: Layered Parameters with Explicit Detection

```powershell
[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Target platform (x64, ARM64, or All). Default is All.")]
    [ValidateSet("x64", "ARM64", "All")]
    [string]$Platform = "All",

    [Parameter(HelpMessage = "Install variant (e.g., asan, static, etc.)")]
    [switch]$Variant,

    [Parameter(HelpMessage = "Explicit override for exact target")]
    [string]$ExplicitTarget = $null
)

$targetsToInstall = @()

if ($ExplicitTarget) {
    # Highest priority: explicit override
    $targetsToInstall = @($ExplicitTarget)
} elseif ($Platform -eq "All") {
    # Default: install all combinations
    $variantExplicit = $PSBoundParameters.ContainsKey('Variant')
    if ($variantExplicit) {
        if ($Variant) {
            $targetsToInstall = @("x64-variant", "arm64-variant")
        } else {
            $targetsToInstall = @("x64-standard", "arm64-standard")
        }
    } else {
        # Not specified: install everything
        $targetsToInstall = @("x64-standard", "arm64-standard", "x64-variant", "arm64-variant")
    }
} else {
    # Specific platform: one target
    $arch = if ($Platform -eq "ARM64") { "arm64" } else { "x64" }
    $suffix = if ($Variant) { "-variant" } else { "-standard" }
    $targetsToInstall = @("$arch$suffix")
}

# Install loop
foreach ($target in $targetsToInstall) {
    Write-Host ""
    Write-Host "=== Installing $target ===" -ForegroundColor Cyan
    
    # Invoke build/install command here
    & SomeInstaller.exe --target $target
    
    if ($LASTEXITCODE -ne 0) {
        throw "Installation failed for target $target with exit code $LASTEXITCODE"
    }
}
```

## Key Techniques

### 1. Explicit vs Default Detection

**Problem:** Distinguish "user didn't pass switch" from "user passed `-Switch:$false`"

**Solution:** Use `$PSBoundParameters.ContainsKey()`

```powershell
$variantExplicit = $PSBoundParameters.ContainsKey('Variant')
if ($variantExplicit) {
    # User explicitly passed -Variant or -Variant:$false
    if ($Variant) {
        # Install variant targets
    } else {
        # Install non-variant targets
    }
} else {
    # User didn't specify: install BOTH variant and non-variant
}
```

This enables the "install everything by default" behavior while still allowing users to narrow down.

### 2. Progressive Priority Resolution

```text
1. Explicit override (-Triplet, -Target) → install exactly that
2. Specific platform (-Platform x64/ARM64) → install one target
3. All platforms (-Platform All, default) → install multiple targets
   - With explicit variant flag → subset of targets
   - Without variant flag → all targets
```

Implement as nested if-elseif:

```powershell
if ($ExplicitOverride) {
    # Priority 1
} elseif ($Platform -eq "All") {
    # Priority 3 (default)
} else {
    # Priority 2
}
```

### 3. Multi-Target Loop Pattern

Build target array, then loop:

```powershell
$targetsToInstall = @()

# ... resolution logic fills $targetsToInstall ...

foreach ($target in $targetsToInstall) {
    Write-Host ""
    Write-Host "=== Installing $target ===" -ForegroundColor Cyan
    
    & SomeCommand --target $target
    
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed for target $target with exit code $LASTEXITCODE"
    }
}
```

**Benefits:**
- Single code path for install logic
- Per-target progress headers
- Fail-fast on error (aborts loop)
- Easy to add post-install validation loop

### 4. Per-Target Progress

Print clear headers before each install:

```powershell
Write-Host ""
Write-Host "=== Installing $target ===" -ForegroundColor Cyan
```

Users see exactly which target is running and where failures occur in multi-target scenarios.

### 5. Post-Install Validation Loop

Separate loop for validation after all installs succeed:

```powershell
foreach ($target in $targetsToInstall) {
    $checkPath = "path\to\$target\artifact"
    if (Test-Path $checkPath) {
        Write-Host "OK: $target installed at $checkPath" -ForegroundColor Green
    } else {
        Write-Host "Warning: $target artifact not found at $checkPath" -ForegroundColor Yellow
    }
}
```

Keeps install loop clean and allows batch validation.

## Real-World Example: vcpkg-install.ps1

```powershell
[CmdletBinding()]
param(
    [ValidateSet("x64", "ARM64", "All")]
    [string]$Platform = "All",
    
    [switch]$Asan,
    
    [string]$Triplet = $null
)

$tripletsToInstall = @()

if ($Triplet) {
    $tripletsToInstall = @($Triplet)
} elseif ($Platform -eq "All") {
    $asanExplicit = $PSBoundParameters.ContainsKey('Asan')
    if ($asanExplicit) {
        if ($Asan) {
            $tripletsToInstall = @("x64-windows-asan", "arm64-windows-asan")
        } else {
            $tripletsToInstall = @("x64-windows", "arm64-windows")
        }
    } else {
        $tripletsToInstall = @("x64-windows", "arm64-windows", "x64-windows-asan", "arm64-windows-asan")
    }
} else {
    $arch = if ($Platform -eq "ARM64") { "arm64" } else { "x64" }
    $suffix = if ($Asan) { "-asan" } else { "" }
    $tripletsToInstall = @("$arch-windows$suffix")
}

foreach ($triplet in $tripletsToInstall) {
    Write-Host ""
    Write-Host "=== Installing $triplet ===" -ForegroundColor Cyan
    
    & vcpkg.exe install --triplet $triplet --x-manifest-root . --x-install-root .build\vcpkg_installed
    
    if ($LASTEXITCODE -ne 0) {
        throw "vcpkg install failed for triplet $triplet with exit code $LASTEXITCODE"
    }
}
```

## Usage Examples

```powershell
# Install everything (default)
.\vcpkg-install.ps1

# Install specific platform (one target)
.\vcpkg-install.ps1 -Platform x64

# Install specific platform + variant
.\vcpkg-install.ps1 -Platform ARM64 -Asan

# Install all platforms, variant only
.\vcpkg-install.ps1 -Platform All -Asan

# Install all platforms, standard only
.\vcpkg-install.ps1 -Platform All -Asan:$false

# Explicit override (bypass all logic)
.\vcpkg-install.ps1 -Triplet x64-windows-static
```

## Anti-Patterns

### ❌ No Explicit Detection
```powershell
# BAD: Can't distinguish "not specified" from "-Variant:$false"
if ($Variant) {
    # variant targets
} else {
    # non-variant targets — what about "install both"?
}
```

### ❌ Nested Loops Instead of Target Array
```powershell
# BAD: Duplicates install logic
foreach ($platform in @("x64", "arm64")) {
    foreach ($variant in @("standard", "asan")) {
        & Install.exe --platform $platform --variant $variant
    }
}

# GOOD: Single loop over resolved targets
$targets = @("x64-standard", "arm64-standard", "x64-asan", "arm64-asan")
foreach ($target in $targets) {
    & Install.exe --target $target
}
```

### ❌ Per-Target Cleanup
```powershell
# BAD: Cleans before each target
foreach ($target in $targets) {
    Remove-Item -Path $installRoot -Recurse -Force
    & Install.exe --target $target
}

# GOOD: Clean once before loop
if ($Clean) {
    Remove-Item -Path $installRoot -Recurse -Force
}
foreach ($target in $targets) {
    & Install.exe --target $target
}
```

## When to Use This Pattern

✅ **Use when:**
- Script supports multiple platforms, architectures, or configurations
- Default behavior should be "install everything"
- Users need fine-grained control (specific platform, specific variant)
- Build targets can be installed independently (parallel or sequential)

❌ **Don't use when:**
- Only one target exists
- Targets have complex dependencies (use dependency graph instead)
- "Install everything" would be prohibitively slow (default to minimal instead)

## Related Patterns

- **Parameter Validation:** Use `[ValidateSet()]` for platform/variant enums
- **Error Propagation:** Check `$LASTEXITCODE` after external commands, throw on failure
- **Progress Reporting:** Print per-target headers with `-ForegroundColor Cyan`
- **Fail-Fast:** Abort loop on first error (don't continue installing after failure)

## See Also

- `vcpkg-install.ps1` — reference implementation
- PowerShell `$PSBoundParameters` documentation
- PowerShell `[ValidateSet()]` attribute
