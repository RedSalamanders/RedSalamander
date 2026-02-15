# RedSalamander

A Windows-based C++ application featuring advanced text visualization, real-time debugging, and high-performance graphics rendering.

## Building the Project

### Prerequisites
- **Visual Studio 2026** (or later) with "Desktop development with C++" workload
- **vcpkg** for package management
- **Windows 11** (x64)

### Quick Start

#### Command Line Build

Use the `build.ps1` PowerShell script for easy building:

```powershell
# Build entire solution in Debug configuration (default)
.\build.ps1

# Build in Release configuration
.\build.ps1 -Configuration Release

# Build specific project only
.\build.ps1 -ProjectName RedSalamanderMonitor
.\build.ps1 -ProjectName RedSalamander
.\build.ps1 -ProjectName Common

# Clean build
.\build.ps1 -Clean

# Rebuild all
.\build.ps1 -Rebuild

# Combined options
.\build.ps1 -Configuration Release -ProjectName RedSalamanderMonitor -Rebuild
```

**Build Script Parameters:**
- `-Configuration` : `Debug` (default), `Release`, or `ASan Debug`
- `-Platform` : `x64` (default) or `ARM64`
- `-ProjectName` : Specific project name (builds entire solution if not specified)
- `-Clean` : Perform clean build
- `-Rebuild` : Rebuild all projects

#### Visual Studio Build

1. Open `RedSalamander.sln` in Visual Studio 2022
2. Select configuration (Debug/Release) and platform (x64)
3. Build → Build Solution (Ctrl+Shift+B)

### Solution Structure

The solution contains the following projects:
- **RedSalamander**: Main application
- **RedSalamanderMonitor**: Monitoring application with advanced text rendering
- **Common**: Shared library with utilities
- **FileSystem**: Plugin for virtual file system
- **FileSystemCurl**: Plugin providing FTP/SFTP/SCP/IMAP virtual file systems
- **PoC Projects**: ls1, ls2, ls3, ls4, FlipSequentialDiscard (proof-of-concept)

### Output

Built executables and libraries are located in:
```text
.build\x64\Debug\     (Debug builds)
.build\x64\Release\   (Release builds)
.build\ARM64\Debug\   (Debug builds)
.build\ARM64\Release\ (Release builds)
```

### MSIX Installer

Build Release + MSIX in one command:

```powershell
.\build.ps1 -Msix
```

Or build/package separately:

```powershell
.\build.ps1 -Configuration Release
msbuild Installer\msix\RedSalamanderInstaller.wapproj /p:Configuration=Release /p:Platform=x64
```

MSIX output is written to:
```text
.build\AppPackages\
```

See `Specs/InstallerMsixSpec.md` for signing and deployment details.

### MSI Installer

Build Release + MSI in one command (requires WiX Toolset v6+):

```powershell
.\build.ps1 -Msi
```

MSI output is written to:
```text
.build\AppPackages\
```

See `Specs/InstallerMsiSpec.md` for details.

## VCPKG Integration

### Enable User-Wide Integration

This makes MSBuild aware of vcpkg's installation path:

```shell
vcpkg.exe integrate install
```

This outputs:
```text
All MSBuild C++ projects can now #include any installed libraries. 
Linking will be handled automatically. Installing new libraries will 
make them instantly available.
```

### Installing Dependencies

Install all libraries from `vcpkg.json`:

```powershell
.\vcpkg-install.ps1

# ARM64:
.\vcpkg-install.ps1 -Platform ARM64
```

### Adding New Libraries

To add a new library:

```shell
vcpkg.exe add port <library-name>
```

This will download and build the library, making it available for use in your projects.

### Using vcpkg in Visual Studio

1. Open Visual Studio
2. Open the project you want to use vcpkg with
3. Ensure vcpkg integration is enabled:
   - Right-click on the project in Solution Explorer
   - Select "Properties"
   - Under "Configuration Properties", check if "Vcpkg" is listed
   - If listed, vcpkg integration is enabled

## Technology Stack

- **Language**: C++23
- **Build System**: MSBuild / Visual Studio 2022
- **Package Manager**: vcpkg
- **Graphics**: Direct2D, DirectWrite, Direct3D 11, DXGI
- **Platform**: Windows 10/11 (Unicode UTF-16)

## ETW Tracing (RedSalamanderMonitor)

[RedSalamanderMonitor](RedSalamanderMonitor/) is a real-time ETW (Event Tracing for Windows) viewer for RedSalamander events.
On some machines it may need extra privileges to start its ETW listener.

### One-time permission setup (avoid UAC prompts)

Run the helper script once to add your account to the local **Performance Log Users** group (the script will self-elevate):

```powershell
.\init-etw-trace.ps1
```

Then **sign out/in (or reboot)** so your access token picks up the new group membership, and launch:

- `.build\x64\Debug\RedSalamanderMonitor.exe` (Debug), or
- `.build\x64\Release\RedSalamanderMonitor.exe` (Release)

To undo the change later:

```powershell
.\init-etw-trace.ps1 -Remove
```

### Optional: capture an `.etl` trace file

If you need an external ETW session (e.g. Windows Performance Analyzer), use:

- `.\start-etw-trace.ps1`
- `.\stop-etw-trace.ps1`
- `.\clean-etw-trace.ps1`

Note: `RedSalamanderMonitor.exe` has its own built-in ETW listener and does not require an external session for normal use.

## Additional Documentation

- **Docs/**: User documentation (start at `Docs/README.md`)
- **AGENTS.md**: Comprehensive development guidelines for AI assistants and developers
- **.github/copilot-instructions.md**: GitHub Copilot specific guidelines
- **Specs/**: Detailed specifications for various components

## License

See `LICENSE.txt` for license information.

## User Documentations

See the `Docs/` folder for user manuals and guides:
- `Docs/RemoteFileSystems.md` (FTP/SFTP/SCP/IMAP)
	
