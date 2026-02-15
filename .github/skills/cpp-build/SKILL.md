---
name: cpp-build
description: Build Red Salamander C++ project using PowerShell build script. Use when building, compiling, or rebuilding the solution or specific projects like RedSalamander, RedSalamanderMonitor, Common, or FileSystem.
metadata:
  author: DualTail
  version: "1.0"
---

# Building Red Salamander

## Quick Build Commands

```powershell
# Build entire solution (default)
.\build.ps1

# Build in Release configuration
.\build.ps1 -Configuration Release

# Build specific project
.\build.ps1 -ProjectName RedSalamander
.\build.ps1 -ProjectName RedSalamanderMonitor
.\build.ps1 -ProjectName Common

# Clean and rebuild
.\build.ps1 -Clean
.\build.ps1 -Rebuild
```

## Parameters

| Parameter | Values | Default |
|-----------|--------|---------|
| `-Configuration` | Debug, Release, ASan Debug | Debug |
| `-Platform` | x64, ARM64 | x64 |
| `-ProjectName` | RedSalamander, RedSalamanderMonitor, Common, FileSystem | All projects |
| `-Clean` | Switch | False |
| `-Rebuild` | Switch | False |

## Output Locations

- Debug: `.build\x64\Debug\*.exe, *.dll`
- Release: `.build\x64\Release\*.exe, *.dll`
- Debug (ARM64): `.build\ARM64\Debug\*.exe, *.dll`
- Release (ARM64): `.build\ARM64\Release\*.exe, *.dll`

## Build Order (Dependencies)

1. **Common** - Shared library (no dependencies)
2. **RedSalamanderMonitor** - Monitoring app (depends: Common)
3. **RedSalamander** - Main app (depends: Common, FileSystem, RedSalamanderMonitor)
4. **FileSystem** - Plugin (no dependencies)

## Visual Studio Build

1. Open `RedSalamander.sln` in Visual Studio 2022+
2. Select configuration (Debug/Release) and platform (x64)
3. Build → Build Solution (Ctrl+Shift+B)

## vcpkg Integration

- Uses vcpkg for package management
- Dependencies defined in `vcpkg.json`
- Keep vcpkg.json files up to date
- Pin versions for critical dependencies

## Dependencies

- **WIL** - Windows Implementation Library (RAII wrappers)
- **fmt** - Modern C++ formatting library
- **DirectX** - Graphics and multimedia APIs (D2D, D3D11, DXGI)

## Build Script Features

- Automatically locates MSBuild (VS 2022 or later)
- Builds entire solution when no ProjectName specified
- Shows build time and output file sizes
- Supports multi-processor builds (`/m`)
- Displays both executables when building full solution
