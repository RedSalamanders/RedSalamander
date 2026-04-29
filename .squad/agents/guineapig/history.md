# Project Context

- **Owner:** eric-jesover
- **Project:** RedSalamander — Windows-native C++23 file manager with Direct2D rendering, plugin architecture, and ETW monitoring
- **Stack:** C++23, Win32, Direct2D, DirectWrite, Direct3D 11, DXGI, vcpkg, MSBuild, Visual Studio 2026
- **Key files:** AGENTS.md, CLAUDE.md, Specs/, .github/skills/cpp-build/
- **Created:** 2026-03-21

## Core Context

Agent GuineaPig initialized as Tester. Responsible for testing, spec verification, and edge case discovery.

Key testing context:
- PerformanceTests1/ and PerformanceTests2/ — Existing performance test projects
- Specs/ — Source of truth for expected behavior (UI, Plugins, Core, Themes, Testing)
- Build: `.\build.ps1` (Debug), `.\build.ps1 -Configuration Release`
- Output: `.build\x64\Debug\`, `.build\x64\Release\`, `.build\ARM64\Debug\`, `.build\ARM64\Release\`

Key skills to read before work:
- .github/skills/cpp-build/SKILL.md
- .github/skills/compiler-warnings/SKILL.md

## Learnings

📌 Team initialized on 2026-03-21

📌 DxUiTests build: `build.ps1 -ProjectName DxUiTests` does NOT work — the ProjectName flag causes MSB4057 errors across all projects. Build directly with MSBuild: `MSBuild.exe PoC\DxUiTests\DxUiTests.vcxproj /p:Configuration=Debug /p:Platform=x64`

📌 DxUiTests run: `.\.build\x64\Debug\DxUiTests.exe` — exits with code 0 and prints "DxUiTests passed" on success; first failure exits with code 1.

📌 Device loss recovery path: `DebugSimulateDeviceLoss()` → `DiscardSizeDependentResources()` + `DiscardDeviceResources()` + `ResetSharedWindowHostGraphicsResources()` + `Invalidate()`. Brush cache (`unordered_map<uint32_t, ...>`) and fallback brush are cleared by `DiscardDeviceResources()`. D2D context/device/factory are reset. DWrite factory and text format caches survive (device-independent). Recovery happens lazily on next `Render()` → `EnsureDeviceResources()` → `RecreateBrushCache()`.

📌 Debug accessors available on WindowHost (`#ifdef _DEBUG`): `DebugGetRenderCount`, `DebugGetResizeCount`, `DebugGetResizeFailureCount`, `DebugSimulateDeviceLoss`, `DebugGetBrushCacheSize`, `DebugHasFallbackBrush`, `DebugHasD2DContext`, `DebugGetConfiguredTextFormatCount`, `DebugCreateAccessibilityProvider`.

📌 `AttachedHostWindow` creates a real HWND for tests needing WM_PAINT, swap chain, etc. Window is positioned off-screen (-32000,-32000) but `IsHostWindowEffectivelyVisible` returns true because it only checks `IsWindowVisible` + not iconic.

📌 **vcpkg-install.ps1 staged-merge pattern:** When validating vcpkg provisioning changes, test normal triplets FIRST (e.g., `.\vcpkg-install.ps1 -Platform x64`) before attempting ASAN triplets. The staged-merge architecture (isolated per-triplet staging under `.build\vcpkg_install_staging\<triplet>\`, then atomic merge to `.build\vcpkg_installed\<triplet>\`) ensures sibling triplet preservation and failure isolation. Always verify staging cleanup post-install (`(Get-ChildItem .build\vcpkg_install_staging\<triplet>).Count` should be 0 or nonexistent). For ASAN failures like OpenSSL LNK1158, verify normal triplets remain available after failure — this validates the safety net.

📌 **vcpkg Merge Lock Validation Session (2026-04-28):** Defined 6 binding acceptance criteria for vcpkg triplet merge lock tolerance. Created synthetic test suite (5/5 passing) validating: identical-file skip, differing-file replacement, unreadable-destination fail, lock-blocked-replacement fail, destination-extras preservation. Tests approved by Ripley as production validation baseline. See `.squad/orchestration-log/2026-04-28T19-49-41Z-GuineaPig-vcpkg-merge.md` and `.squad/log/2026-04-28T19-49-41Z-vcpkg-merge-lock-fix.md`.

📌 **ASan vcpkg Provisioning Acceptance Criteria (2026-04-28):** Defined 7 objective criteria for ASan triplet validation:
1. Triplets must NOT use `VCPKG_BUILD_TYPE=release` (MSBuild contract)
2. OpenSSL patch must apply successfully
3. OpenSSL must build successfully
4. No LNK1158 linker recursion errors
5. Both debug and release libraries exist
6. First-party ASan Debug builds link successfully
7. Patch matches OpenSSL source (or eliminated via alternative strategy)

These criteria are binding and inform all ASan provisioning work. Release-only triplet approach was explicitly rejected by Ripley because it violates the MSBuild contract (ASan Debug config hardwired to x64-windows-asan triplet; that triplet must support debug linking). See `.squad/decisions.md` sections 10–12 for full context and merger outcome.
