# Project Context

- **Owner:** eric-jesover
- **Project:** RedSalamander — Windows-native C++23 file manager with Direct2D rendering, plugin architecture, and ETW monitoring
- **Stack:** C++23, Win32, Direct2D, DirectWrite, Direct3D 11, DXGI, vcpkg, MSBuild, Visual Studio 2026
- **Key files:** AGENTS.md, CLAUDE.md, Specs/, .github/skills/
- **Created:** 2026-03-21

## Core Context

Agent Ripley initialized as Lead / Reviewer. Responsible for architecture, code review, and design decisions. Always check Specs/ and .github/skills/ before reviewing.

Key project components:
- RedSalamander/ — Main file manager application (FolderWindow, FolderView)
- RedSalamanderMonitor/ — Monitoring tool with ColorTextView (~200KB D2D text editor)
- Common/ — Shared utilities, PlugInterfaces (COM-style interfaces)
- Plugins/ — FileSystem, ViewerText, ViewerSpace, ViewerImgRaw, ViewerVLC, ViewerPE, FileSystem7z

### Historical Deep-Dives (2026-03-21 through 2026-07)
Conducted comprehensive reviews of three major subsystems with architectural findings documented in `.squad/decisions/inbox/`:

**DxUi Framework (2025-07-24):** 13 controls, WindowHost device management, shared resource generation tracking, PruneStaleInteractionState safety pattern. Found: text-editing duplication (ComboBox/TextField), SAFEARRAY manual cleanup. Strengths: exemplary AGENTS.md compliance, robust device-loss, clean API. Complementary D2D findings (Yoko): BeginDraw/EndDraw not scope_exit protected, accessibility TOCTOU race, Grid/Tree brush cache O(n) linear search, no dirty rectangle tracking. All COM resources correct wil::com_ptr usage.

**FileSystem Plugin (2025-07-25):** 10 files, ~4,500 lines, 8 interfaces. Critical issues: SharedFileOpsJobScheduler static singleton DLL unload race, DirectoryWatch callbacks not drained, WatchDirectory TOCTOU. Strengths: zero AGENTS.md violations, sophisticated threadpool watch + buffer pool, cooperative cancellation scheduler, correct extended path handling (\\?\), thorough reparse point handling, smart buffer trimming.

**Cloud Drive Plugins (2026-07-24):** GoogleDrive (1,861 LOC, read-only, libcurl) and MicrosoftDrive (4,882 LOC, full-featured, WinHTTP). Safety fixes implemented: Phase 2 (callback drain generation tracking + condition variable), Phase 3 (COM ownership wil::com_ptr). Pattern: double-buffer configuration strings, generation tracking for callback teardown. Cross-plugin: massive Utf16/Utf8/path/JSON duplication should be factored to Common/.

**Comprehensive WIP Plan:** `Specs/Plans/WIP/Code_DxUiAndFileSystemImprovementPlan.md` consolidates all findings. 14 sections, agent parallelization pattern (Stream A/B/C/D) established.

## Learnings

## Learnings

📌 **vcpkg Merge Lock Fix Session (2026-04-28):** GuineaPig defined 6 binding acceptance criteria for safe triplet merge. Sysadm attempted robocopy (rejected—insufficient for lock tolerance). MergeDoctor attempted hash-based with silent skip (rejected—violates clear-fail requirement). OpusMerge implemented approved solution: hash-based safe merge in `Tools\VcpkgInstallSafety.ps1` with explicit lock/readability failures. Tests: Pester 4/4, Synthetic 5/5. Ripley: APPROVED for production. See `.squad/log/2026-04-28T19-49-41Z-vcpkg-merge-lock-fix.md` for session details.

📌 **Team initialized on 2026-03-21

📌 **Preferences Dialog DxUi Migration — Complete (2025-03-22)**
- Led three-agent parallel review (Ripley review, Yoko per-page flags, Sysadm shell chrome)
- Approved forward migration strategy (complete DxUi, not revert to Win32)
- Verified all 6 critical bugs fixed: buttons visible, sizing correct, scrolling functional, page switching clean
- Identified 5 cleanup follow-ups for next pass
- Pattern learned: DxUi migration benefits from focused completion pass with feature flag removal

📌 **Phase 2 DxUi Migration — APPROVED WITH NOTES (c37f1b65)**
- Reviewed removal of all 6 `kEnable*DxSurface` flags and 14 stale HWND gates across 7 files
- Zero AGENTS.md regression guard violations (no raw new/delete, no sprintf_s, no catch(...), correct RAII)
- Verified data flow correctness: state.viewersSelectedExtensionText and state.pluginsSelectedPluginId serve as DxUi source-of-truth when legacy HWNDs are absent
- Verified DxUi callbacks (search edit, extension edit, viewer combo, grid selection, add/remove mapping) all guard HWND operations correctly without blocking DxUi-only paths
- Build passes clean (x64 Debug)
- Minor: stale comment in Panes.cpp line 31 ("Re-landed on the stabilized one-host page pattern.") should be removed
- Follow-up needed: Keyboard.cpp (Refresh gates on state.keyboardList) and Themes.cpp (UpdateEditorFromSelection gates on state.themesColorsList) have similar stale HWND gates not addressed in this commit

📌 **ASan vcpkg Provisioning Architecture Decision (2026-03-21 & 2026-04-28):** REJECTED release-only triplet workaround (`VCPKG_BUILD_TYPE=release`) as architecturally unsound. Key findings:
- MSBuild contract: `Directory.Build.props` lines 46–49 hardwire ASan Debug → x64-windows-asan triplet; that triplet must support debug linking
- Release-only triplet breaks linker phases (LNK2019 unresolved externals for yyjson, libcurl, sqlite3)
- Single Responsibility: A triplet owns one build type and must provide libs matching that type
- Constraint: Next agent must implement multi-config triplet (both debug + release libs) with OpenSSL conflict verification

Outcome: OpusVcpkg implemented the correct solution via portfile.cmake flag-stripping (regex removes bare `-INCREMENTAL` from `VCPKG_COMBINED_*_LINKER_FLAGS_DEBUG`, preserving `-INCREMENTAL:NO`). Root cause (triplet flag merging) is fixed. Remaining `LNK1158` failure is MSVC link.exe self-spawn issue (VS18 Insiders regression), not a config defect. Approved for merge with documented external blocker.
