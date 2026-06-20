# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**RedSalamander** is a Windows-native file manager and monitoring application built in MSVC `stdcpplatest` mode. It features dual-pane file management, plugin-based virtual file systems, advanced text visualization with Direct2D/DirectWrite rendering, and real-time debugging via ETW (Event Tracing for Windows).

## Build Commands

```powershell
# Build entire solution (Debug)
.\build.ps1

# Build specific configuration
.\build.ps1 -Configuration Release

# Build specific project
.\build.ps1 -ProjectName RedSalamander
.\build.ps1 -ProjectName RedSalamanderMonitor

# Clean/Rebuild
.\build.ps1 -Clean
.\build.ps1 -Rebuild
```

Output: `.build\<Platform>\<Configuration>\` (e.g. `.build\x64\Debug\`, `.build\ARM64\Release\`)

Run the full local test suite with `.\Tools\Run-AllTests.ps1 -Suite Full` (or `-SkipBuild` to reuse an existing Debug build). See README "Self-tests" for details.

## Architecture

### Solution Structure (~21 core projects + localization satellites)

```text
RedSalamander/             # Main file manager application (incl. SelfTest\ suites)
RedSalamanderMonitor/      # ETW monitoring/debug tool with ColorTextView
RedSalamanderSearchService/# Background search/index Windows service (named pipe + SQLite)
RedConfigure/              # Standalone configuration tool
RedLauncher/               # Launcher app
Common/                    # Shared library (utilities, settings, DxUi framework)
  └── PlugInterfaces/      # COM-style plugin interfaces (IFileSystem, IViewer, ...)
Plugins/                   # All plugin DLLs
  ├── FileSystem, FileSystem7z, FileSystemCurl, FileSystemS3,
  │   FileSystemGoogleDrive, FileSystemMicrosoftDrive, FileSystemDummy
  └── ViewerText, ViewerSqlite, ViewerSpace, ViewerImgRaw,
      ViewerVLC, ViewerPE, ViewerWeb
Tests/                     # 8 standalone test projects (DxUiTests, PerformanceTests2, ...)
Tools/                     # PowerShell tooling (+ Pester tests in Tools/Tests)
Installer/                 # MSIX + MSI packaging
PoC/                       # Proof-of-concept projects
```

The solution additionally contains ~90 per-language localization satellite resource projects, so Solution Explorer shows far more than the core list.

### Project Dependencies

- **Common** → no dependencies (shared library)
- **RedSalamanderMonitor** → Common
- **RedSalamander** → Common + all plugins
- **Plugins** → independent DLLs using PlugInterfaces

### Key Components

| Component | Location | Purpose |
|-----------|----------|---------|
| FolderWindow | RedSalamander/ | Main window with dual-pane layout |
| FolderView | RedSalamander/ | File list rendering, selection, drag-drop (split into ~10 .cpp files) |
| ColorTextView | RedSalamanderMonitor/ | High-performance D2D text editor (~200KB implementation) |
| PlugInterfaces | Common/PlugInterfaces/ | COM-style interfaces for plugins |
| Helpers.h | Common/ | Core utilities, Debug logging, TraceLogging |
| SettingsStore | Common/ | Registry-based settings persistence |

### Plugin Architecture

Plugins use COM-style interfaces with a factory entry point:

```cpp
extern "C" HRESULT RedSalamanderCreate(REFIID riid, const FactoryOptions*, IHost*, const wchar_t* pluginId, void** ppv);
```

Key interfaces in `Common/PlugInterfaces/`:
- **IFileSystem** - Virtual file system operations (copy, move, delete, search)
- **IViewer** - File viewer with theming support
- **IHost** - Host callbacks for plugins

## Development Guidelines

**See AGENTS.md** for comprehensive guidelines and **.github/skills/** for detailed patterns.

### Critical Rules

1. **RAII is Mandatory** - All Windows resources MUST use WIL wrappers (`wil::unique_hicon`, `wil::unique_hdc`, etc.). Manual cleanup functions (`DestroyIcon`, `DeleteObject`, `EndPaint`) are **PROHIBITED**.

2. **Modern C++23** - Smart pointers, `std::format`, `std::optional` with `.value()` (never `*`), range-based loops, structured bindings.

3. **Error Handling** - Use `Debug::ErrorWithLastError()` for Win32 failures, `Debug::Error()` for unexpected failures, `Debug::Warning()` for recoverable issues. Don't log normal control-flow.

4. **No Hardcoded Strings** - UI strings go in `.rc` resources.

### Patterns to Avoid

- Raw `new`/`delete` → use smart pointers
- C-style casts → use `static_cast`, `reinterpret_cast`
- `goto` → use early returns + `wil::scope_exit`
- Raw Windows handles → use WIL RAII wrappers
- Blocking UI thread
- Global state and singletons

## Skills Reference

Detailed patterns in `.github/skills/`:

| Skill | Use When |
|-------|----------|
| **wil-raii** | Managing Windows handles (HICON, HDC, HWND, COM) |
| **cpp-build** | Building projects, understanding dependencies |
| **direct2d-rendering** | D2D/DirectWrite graphics code |
| **plugin-callbacks** | Implementing plugin callback patterns |
| **theming** | Working with theme colors and JSON5 themes |
| **error-handling** | Logging errors and warnings |
| **async-threading** | Multi-threaded operations |
| **localization** | RC resources and STRINGTABLE |

## Technology Stack

- **Language**: C++23, Unicode UTF-16
- **Build**: Visual Studio 2026 / toolset v145, MSBuild, vcpkg
- **Graphics**: Direct2D, DirectWrite, Direct3D 11, DXGI
- **Platform**: Windows 11 build 22000.2600, x64 and ARM64
- **Key Dependencies**: WIL, fmt, yyjson, libraw, libjpeg-turbo, 7zip, pe-parse

## Specifications

Detailed component specs in `Specs/` folder:
- `Specs/UI/UI_FolderView.md`, `Specs/UI/UI_FolderWindow.md`, `Specs/UI/UI_NavigationView.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Plugins/Plugins_ViewerPlugins.md`, `Specs/Plugins/Plugins_PluginAPI.md`
- `Specs/UI/UI_PreferencesDialog.md`, `Specs/SettingsStore.schema.json`
- `Specs/Core/Core_RedSalamanderMonitor.md`
