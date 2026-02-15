# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Red Salamander** is a Windows-native file manager and monitoring application written in C++23. It features dual-pane file management, plugin-based virtual file systems, advanced text visualization with Direct2D/DirectWrite rendering, and real-time debugging via ETW (Event Tracing for Windows).

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

## Architecture

### Solution Structure (18 projects)

```text
RedSalamander/          # Main file manager application
RedSalamanderMonitor/   # Monitoring/debug tool with ColorTextView
Common/                 # Shared library (utilities, settings, plugin interfaces)
  └── PlugInterfaces/   # COM-style plugin interfaces (IFileSystem, IViewer, etc.)

FileSystem/             # Standard Win32 file system plugin
FileSystem7z/           # 7-Zip archive plugin
FileSystemDummy/        # Test/stub plugin

ViewerText/             # Text/hex viewer plugin
ViewerSpace/            # Disk space visualization
ViewerImgRaw/           # Raw image viewer (libraw)
ViewerVLC/              # Video player (VLC backend)
ViewerPE/               # PE executable viewer

PoC/                    # Proof-of-concept projects
```

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
extern "C" HRESULT RedSalamanderCreate(REFIID riid, const FactoryOptions*, IHost*, void** ppv);
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
- **Build**: Visual Studio 2022+, MSBuild, vcpkg
- **Graphics**: Direct2D, DirectWrite, Direct3D 11, DXGI
- **Platform**: Windows 10/11 x64 only
- **Key Dependencies**: WIL, fmt, yyjson, libraw, libjpeg-turbo, 7zip, pe-parse

## Specifications

Detailed component specs in `Specs/` folder:
- FolderViewSpec.md, FolderWindowSpec.md, NavigationViewSpec.md
- PluginsVirtualFileSystem.md, PluginsViewer.md, PluginAPI.md
- PreferencesDialogSpec.md, SettingsStore.schema.json
- RedSalamanderMonitorSpec.md
