> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# AppShell - verified findings (compact)

- **[crash/PLAUSIBLE]** `RedSalamander/CrashHandler.cpp:390` - Stack-overflow crashes are not handled: no SetThreadStackGuarantee, filter re-faults on exhausted stack
- **[data-loss/CONFIRMED]** `RedSalamander/RedSalamander.cpp:11124` - Save-on-exit is skipped on Windows logoff/shutdown (no WM_QUERYENDSESSION/WM_ENDSESSION handler) -> settings & window placement lost
- **[data-loss/CONFIRMED]** `RedSalamander/RedSalamander.cpp:11055` - No WM_QUERYENDSESSION/WM_ENDSESSION handling: settings, window placement and in-flight file operations are lost on logoff/shutdown
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/RedSalamander.cpp:463` - FindWindowW lookups by class name match windows of OTHER RedSalamander instances (no single-instance guard), enabling cross-process activation and a cross-process pointer send in the test hooks
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/RedSalamander.cpp:11055` - No WM_QUERYENDSESSION/WM_ENDSESSION handling: settings, window placement and session state are silently lost on Windows logoff/restart/shutdown
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/RedSalamander.cpp:7395` - Crash quarantine never disables any plugin: prompt is shown before the host window exists, so HostShowPrompt always fails and the YES branch is unreachable
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/CrashQuarantine.cpp:67` - Crash is mis-attributed to the last active filesystem plugin regardless of where the crash occurred, falsely disabling a healthy (often the built-in local) plugin
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/FileActionLauncher.cpp:614` - Virtual-filesystem / UNC focused path forwarded unvalidated to ShellExecuteEx local verb
- **[incorrect-behavior/CONFIRMED]** `RedSalamander/SplashScreen.cpp:486` - Splash WM_DPICHANGED / WM_SIZE handlers are dead code (intercepted by dxHost): stale window region and label layout after a DPI change
- **[incorrect-behavior/PLAUSIBLE]** `RedSalamander/CrashHandler.cpp:164` - Crash report uses CP_ACP for symbol/file names, mangling and possibly truncating non-ASCII paths
- **[incorrect-behavior/PLAUSIBLE]** `RedSalamander/CrashQuarantine.cpp:165` - Quarantine decision to disable a plugin is held only in in-memory g_settings and is lost if the session ends before the on-exit save
- **[leak/CONFIRMED]** `RedSalamander/RedSalamander.cpp:7558` - InitInstance / accelerator-load failure path leaks initialized plugin managers (background threads/COM/handles abandoned to process teardown)
- **[leak/CONFIRMED]** `RedSalamander/HostServices.cpp:1872` - Window-scoped alert overlay map leaks: entries never erased when plugin target windows are destroyed
- **[leak/PLAUSIBLE]** `RedSalamander/CrashHandler.cpp:524` - Crash marker dump path is passed to ShellExecute and minidump leaks full process memory/secrets
- **[leak/PLAUSIBLE]** `RedSalamander/HostServices.cpp:545` - TOCTOU between IsWindow() check and SendMessageW leaks marshaled payload / drops the call
- **[race/CONFIRMED]** `RedSalamander/HostServices.cpp:657` - Window-null fallthrough runs UI-thread secret bodies on a worker, mutating unsynchronized session-secret maps (data race / heap corruption)
- **[robustness/CONFIRMED]** `RedSalamander/CrashHandler.cpp:352` - Crash handler does heavy heap/lock/Sym work on the faulting thread, deadlocks or re-faults during real crashes
- **[robustness/CONFIRMED]** `RedSalamander/SplashScreen.cpp:603` - Splash worker thread reads owner HWND (g_owner) cross-thread during centering with only IsWindow() as a liveness guard
- **[robustness/PLAUSIBLE]** `RedSalamander/CrashHandler.cpp:390` - Unhandled-exception crash handler does heavy heap/Sym work and is not reentrancy-guarded; unreliable for stack-overflow / heap-corruption faults
- **[robustness/PLAUSIBLE]** `RedSalamander/CrashHandler.cpp:462` - Install() is not idempotent against concurrent first-callers; non-atomic SetXxxHandler sequence under race
- **[robustness/PLAUSIBLE]** `RedSalamander/CrashHandler.cpp:500` - Crash marker survives clean shutdown when the previous-crash UI never runs (self-test/headless and early-exit paths), causing stale quarantine offers on later launches
- **[robustness/PLAUSIBLE]** `RedLauncher/Main.cpp:355` - CreateProcessW reuses GetStartupInfoW output, forwarding the launcher's STARTUPINFO (lpReserved2/dwFlags) into the child unfiltered
- **[robustness/PLAUSIBLE]** `RedSalamander/HostServices.cpp:2100` - Plaintext connection secrets cached in plain std::wstring session maps that can reallocate without scrubbing
- **[security/CONFIRMED]** `RedLauncher/Main.cpp:358` - Launcher spawns RedSalamander.exe without setting lpCurrentDirectory, letting the child inherit an attacker-controlled CWD (DLL-planting hardening gap)
- **[security/CONFIRMED]** `RedSalamander/FileActionLauncher.cpp:616` - External program launched with attacker-controlled working directory enables binary/DLL planting
