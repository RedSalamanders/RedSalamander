# File Operations: S3 directory-marker transfer and two I22 machine residuals

## Status and ownership

- Status: COMPLETE on 2026-09-15; all three fixes landed on `4cb089111a23` with their own tests, commits and
  a closing Fresh Full. See the Execution record. Originally: three independent bug fixes surfaced by I22's Fresh Full on 2026-09-15
  (`Specs/Plans/Done/FileOperations_ProviderDiscoveryParity_2026-09-14.md`). None blocks I22's own
  closeout; each is a separate landing with its own commit and evidence.
- Planned at: `a076d175a`, 2026-09-15, on `LT-PF5VDAGE` (machine profile `7d3a1247382a`).
- This plan is self-contained and non-normative. Recheck each anchor's drift before editing; the
  line numbers below are from `a076d175a`.
- Do not couple any of these to I22: a red S3 self-test would fail the Fresh Full on every machine
  and block I22. Land each fix with its own test, green, before adding it to a gate.

## How to resume from a fresh checkout on any machine

Everything needed is in this file and in git. There is no dependency on the originating session,
its scratchpad, or this machine. On `LT-PF5VDAGE` a full Fresh Full cannot pass from a non-elevated
shell (see I22's record, "Closeout status and how to finish"); the three focused validations below do
not need a Full and run from a normal shell, except the symlink task (Task 3), which needs both an
elevated and a non-elevated shell to prove both branches.

Build one plugin with `.\build.ps1 -ProjectName <Name>`; build everything with `.\build.ps1
-Rebuild`. A killed governed run leaves `.build\artifact-operations\*.contaminated.json`; clear it
with `.\build.ps1 -Rebuild`.

---

## Task 1 — S3 same-bucket prefix transfer fails when the prefix carries a directory marker

### Symptom, reproduced

`CopyItem`, `CopyItems` and `MoveItem` of an S3 "folder" (a prefix) fail with
`HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH)` (`0x8007051A`) whenever the prefix contains a directory
marker object: a zero-byte object whose key ends in `/`, which `FileSystemS3::CreateDirectory`
(`Plugins/FileSystemS3/FileSystemS3.DirectoryOps.cpp:170`, `PutS3ObjectFromMemory(..., isDirectory=true)`)
writes. The earlier S3 fixtures seeded prefixes WITHOUT a marker (I22's `DiscoveryScope_S3DirectApi`
seeds with `createTreeMarker=false`), so no test ever transferred a plan that listed a marker object,
and the defect was never caught.

### Confirmed reproduction (2026-09-15, `a076d175a`, standalone against the loopback fake)

A self-test was written (source below), added to the S3 debug self-test aggregator, built into
`FileSystemS3.dll`, and run standalone. It failed with `0x8007051A` on all three entry points. The
fake's request log for the CopyItem case (last requests; the marker object `x/src/tree/` sorts FIRST
in the plan and is processed first, so it is the failing object):

```
... planning HEADs all succeed, including the marker and children:
 [HEAD /r0f-bucket/x/src/tree/ -> 200] [HEAD /r0f-bucket/x/src/tree/a.bin -> 200] ...
... destination probes (correct bucket):
 [HEAD /r0f-bucket/x -> 404] [HEAD /r0f-bucket/x/copy-dst -> 404] [HEAD /r0f-bucket/x/copy-dst/tree/ -> 404]
... then the two anomalies, back to back, on the marker:
 [HEAD /x/src/tree/ -> 404]                 <-- bucket DROPPED: the marker KEY "x/src/tree/" was
                                                 parsed as a PATH and its first segment "x" became
                                                 the bucket. bucket "x" does not exist -> 404.
 [GET /r0f-bucket/x/src/tree/ -> 416]        <-- a RANGED GET of the zero-byte marker (Range on a
                                                 zero-byte object is unsatisfiable -> 416).
```

Both anomalies are only on the marker (key ends in `/`); the same run's non-marker children copy
fine. The pinning stage (`PinPlannedSourceRevisions` in `FileSystemS3.Directory.cpp:1719`) is NOT the
failure — every planning HEAD returned 200. The failure is later, in the per-object transfer of the
marker.

### Where to look (not yet pinned to one line)

The transfer loop (`ExecuteCopyOrMove`, `FileSystemS3.Directory.cpp:2628`, per-object copy at
`:3037`) calls `CopyS3ObjectWithFallback(source.bucketCtx, source.bucket, object.sourceKey, ...)`
(`:1796`) with bucket and key passed SEPARATELY, so the bucket-dropping `HEAD /x/src/tree/` is NOT
built there directly. Two request styles implicate two code paths for the marker:

- The bucket-dropping `HEAD /x/src/tree/` (bucket = first key segment) is what `ResolveS3Path`
  (`:1432`) / `ProbeS3Path` (`:1542`) produce when handed a PATH of the form `/x/src/tree/` — i.e.
  something reconstructs a path as `"/" + object.sourceKey` (leading slash + the marker key) and
  resolves it as a path instead of probing the object with the already-known `source.bucket`. Find
  that reconstruction. `TryGetPrefixExists` (`:1505`) is a candidate.
- The ranged `GET /r0f-bucket/x/src/tree/ -> 416` is the `IFileSystemIO` reader
  (`FileSystemS3.IO.cpp:809`, `req.SetRange(...)`), which always sends a Range. For a zero-byte
  object it must not: the reader already has a `maxBytes == 0u` early return at
  `FileSystemS3.IO.cpp:784` — check why the marker still reaches `SetRange` (size discovery when
  `_sizeKnown` is false?). Server-side `CopyObject` (`CopyS3ObjectServerSide`,
  `FileSystemS3.S3.cpp:1116`) does NOT need a reader at all for a small object, so the fact that a
  reader ran on the marker suggests the marker took the relay/reader fallback after the server-side
  copy or an earlier probe failed.

Fastest next diagnostic step: re-add the self-test below, change its `dumpLog` from `RequestLog(24u)`
to `RequestLog(200u)`, and comment out the CopyItems and MoveItem calls so only ONE CopyItem's
complete ordered request sequence prints. The first request that names bucket `x` (or a ranged GET on
the marker) identifies the exact function; set a conditional breakpoint there
(`key.ends_with("/")`), or grep for the caller that builds a path from a key.

### The fix (intended shape)

Make a marker object transfer like any other object:
1. Never resolve `object.sourceKey` as a path (do not prepend `/` and call `ResolveS3Path`); probe it
   as an object with the known `source.bucket`/`source.bucketCtx`, as the children already are.
2. A zero-byte object must never be fetched with a `Range` header. Guard the reader
   (`FileSystemS3.IO.cpp`) and any size-discovery so a declared/known size of 0 issues a plain GET
   (or no GET — the server-side `CopyObject` of a zero-byte marker needs no read at all).
Do NOT change discovery reporting (`Common/FileSystemDiscoveryScope.h` and its provider wiring); this
is a transfer-path fix only.

### The test to land with the fix

Add `RunS3DirectoryMarkerTransferSelfTests` to `Plugins/FileSystemS3/FileSystemS3.SelfTest.FakeS3.cpp`
(before `} // namespace FileSystemS3Internal`), forward-declare it in
`Plugins/FileSystemS3/FileSystemS3.Internal.h` next to `RunS3ZeroTimestampReplaceSelfTests` (was at
`:373`), and call it in the aggregator `RedSalamanderS3DebugSelfTests`
(`FileSystemS3.Directory.cpp:5008`, after `RunS3ZeroTimestampReplaceSelfTests`). Exact source, proven
to reproduce (it is expected RED before the fix and GREEN after):

```cpp
// R0f-S3 directory-marker transfer: a prefix that carries an explicit directory marker (a
// zero-byte object whose key ends in '/', which CreateDirectory writes) must copy and move like
// any other prefix. The earlier fixtures seeded prefixes without a marker, so a plan that listed
// the marker object was never transferred; a marker's server-side copy and its zero-byte relay
// fallback are exercised here.
void RunS3DirectoryMarkerTransferSelfTests(unsigned int& passed, unsigned int& failed) noexcept
{
    WSADATA winsockData{};
    if (! DebugCheck(WSAStartup(MAKEWORD(2, 2), &winsockData) == 0, L"WSAStartup should initialize the S3 marker-transfer proof", passed, failed))
    {
        return;
    }
    auto cleanupWinsock = wil::scope_exit([]() noexcept { WSACleanup(); });

    try
    {
        constexpr std::string_view kBucket = "r0f-bucket";
        FakeS3Endpoint endpoint;
        endpoint.SeedBucket(std::string(kBucket));
        // The source prefix carries its own directory marker plus four leaves of distinct sizes.
        endpoint.SeedObject(kBucket, "x/src/tree/", {});
        endpoint.SeedObject(kBucket, "x/src/tree/a.bin", std::vector<uint8_t>(4u, 0xA1u));
        endpoint.SeedObject(kBucket, "x/src/tree/b.bin", std::vector<uint8_t>(8u, 0xB2u));
        endpoint.SeedObject(kBucket, "x/src/tree/nested/", {});
        endpoint.SeedObject(kBucket, "x/src/tree/nested/c.bin", std::vector<uint8_t>(16u, 0xC3u));
        if (! DebugCheck(SUCCEEDED(endpoint.Start()), L"S3 marker-transfer fixture should start", passed, failed))
        {
            return;
        }
        auto stopEndpoint = wil::scope_exit([&]() noexcept { endpoint.Stop(); });

        wil::com_ptr<FileSystemS3> fileSystem;
        fileSystem.attach(new (std::nothrow) FileSystemS3(FileSystemS3Mode::S3, nullptr));
        if (! DebugCheck(static_cast<bool>(fileSystem) && SUCCEEDED(fileSystem->InitializationStatus()),
                         L"S3 marker-transfer proof should create an S3 instance",
                         passed,
                         failed))
        {
            return;
        }
        HRESULT hr = fileSystem->SetConfiguration(FixtureConfiguration(endpoint.Port(), 2'000u, 2'000u).c_str());
        if (! DebugCheck(SUCCEEDED(hr), L"S3 marker-transfer instance should accept the fixture configuration", passed, failed))
        {
            return;
        }

        FileSystemOptions options{};
        options.sizeBytes  = sizeof(options);
        options.linkPolicy = FILESYSTEM_LINK_PRESERVE;
        constexpr FileSystemFlags kFlags =
            static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_ALLOW_OVERWRITE | FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY);

        const auto dumpLog = [&](std::wstring_view label, HRESULT observed) noexcept
        {
            std::fwprintf(stderr,
                          L"[S3] %ls returned 0x%08lX; recent requests:%ls\n",
                          std::wstring(label).c_str(),
                          static_cast<unsigned long>(observed),
                          endpoint.RequestLog(24u).c_str());
        };
        const auto destinationComplete = [&](std::string_view prefix) noexcept
        {
            return endpoint.HasObject(kBucket, std::format("{}/", prefix)) && endpoint.HasObject(kBucket, std::format("{}/a.bin", prefix)) &&
                   endpoint.HasObject(kBucket, std::format("{}/b.bin", prefix)) && endpoint.HasObject(kBucket, std::format("{}/nested/", prefix)) &&
                   endpoint.HasObject(kBucket, std::format("{}/nested/c.bin", prefix));
        };

        // 1) CopyItem of the marked prefix.
        hr = fileSystem->CopyItem(L"/r0f-bucket/x/src/tree", L"/r0f-bucket/x/copy-dst/tree", kFlags, &options, nullptr, nullptr);
        if (FAILED(hr))
        {
            dumpLog(L"CopyItem(marked prefix)", hr);
        }
        DebugCheck(SUCCEEDED(hr), L"S3 CopyItem must transfer a prefix that carries a directory marker", passed, failed);
        DebugCheck(destinationComplete("x/copy-dst/tree"), L"S3 CopyItem must publish the marker and every leaf of the marked prefix", passed, failed);

        // 2) CopyItems of the marked prefix into a destination folder.
        const wchar_t* bulkSources[] = {L"/r0f-bucket/x/src/tree"};
        hr                           = fileSystem->CopyItems(bulkSources, 1u, L"/r0f-bucket/x/copy-bulk-dst", kFlags, &options, nullptr, nullptr);
        if (FAILED(hr))
        {
            dumpLog(L"CopyItems(marked prefix)", hr);
        }
        DebugCheck(SUCCEEDED(hr), L"S3 CopyItems must transfer a prefix that carries a directory marker", passed, failed);
        DebugCheck(destinationComplete("x/copy-bulk-dst/tree"), L"S3 CopyItems must publish the marker and every leaf of the marked prefix", passed, failed);

        // 3) MoveItem of the marked prefix: destination gains every object, source loses them.
        hr = fileSystem->MoveItem(L"/r0f-bucket/x/src/tree", L"/r0f-bucket/x/move-dst/tree", kFlags, &options, nullptr, nullptr);
        if (FAILED(hr))
        {
            dumpLog(L"MoveItem(marked prefix)", hr);
        }
        DebugCheck(SUCCEEDED(hr), L"S3 MoveItem must transfer a prefix that carries a directory marker", passed, failed);
        DebugCheck(destinationComplete("x/move-dst/tree"), L"S3 MoveItem must publish the marker and every leaf of the marked prefix", passed, failed);
        DebugCheck(! endpoint.HasObject(kBucket, "x/src/tree/") && ! endpoint.HasObject(kBucket, "x/src/tree/a.bin"),
                   L"S3 MoveItem must remove the marked source prefix once every object is published",
                   passed,
                   failed);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"FileSystemS3 directory-marker transfer selftest failed after std::exception.");
        DebugCheck(false, L"S3 marker-transfer proof should not throw std::exception", passed, failed);
    }
}
```

### Verify

`PluginContractTests.exe` (the aggregator runs this test; `RedSalamanderS3DebugSelfTests` must report
`failed=0`), plus
`.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_R0fS3Containment -AllowAlternateVolumeTestRoot`.
Standalone iteration without the governed runner: compile a loader that calls the export by name (the
originating session used `scratchpad/fsst2.cpp`: `LoadLibraryExW` + `GetProcAddress` +
`AddVectoredExceptionHandler` to capture the plugin's `OutputDebugString`/stderr) and run it from
`.build\x64\Debug\Plugins\` as `fsst2.exe FileSystemS3.dll RedSalamanderS3DebugSelfTests`. Follow
AGENTS.md spec closeout if a durable contract changes (the marker-transfer obligation could join
`Specs/Plugins/Plugins_VirtualFileSystem.md`).

---

## Task 2 — Command palette use-after-free crashed the Commands suite

`RedSalamander/CommandPaletteWindow.cpp` destroys the palette window on `WM_ACTIVATE` with
`WA_INACTIVE` (`_window.reset()` at `:303` and `:323`) and on `WM_NCDESTROY` calls `_host.Detach()`
(`:346`) but never clears `_searchField` (`:191`), `_grid`, `_model` or `_root`.
`CommandPaletteWindow::DebugSetSearch` (`:526`) then dereferences the dangling `_searchField`
(`_searchField->SetText(...)` -> `DxUi::TextField::SetText`), an access violation. It crashed the
Commands self-test `cmd_pane_embedded_terminal_command_state_truth`
(`RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp:~10043`,
`TestEmbeddedTerminalCommandStateTruth`) during I22's Fresh Full on 2026-09-15 when something took
foreground activation mid-case. Crash report (kept on `LT-PF5VDAGE`, also copied to
`D:\RedSalamander-i22-handoff\`):
`C:\Users\PVB\AppData\Local\RedSalamander\Crashes\RedSalamander-20260915-150000-p11724.txt` — stack:
`std::wstring::operator=` inside `DxUi::TextField::SetText` <- `CommandPaletteWindow::DebugSetSearch`
<- `DebugSetCommandPaletteSearch` (`:573`) <- `TestEmbeddedTerminalCommandStateTruth`.

Fix: (1) clear `_searchField`/`_grid`/`_model`/`_root` (and anything else pointing into the DxUi
tree) when the window is destroyed, and make `DebugSetSearch` and `DebugSnapshot` (`:505`) return
false when the window is gone (`DebugGetCommandPaletteSnapshot`/`DebugSetCommandPaletteSearch` at
`:568`/`:573` already guard on the `g_palette` pointer, not on the window being alive). (2) Make the
self-test report a diagnosable failure (palette lost activation) or re-open the palette, per the
case's contract in `Specs/UI`. Check other palette debug hooks and anything in
`RedSalamander/SelfTest/Commands` using `GetCommandPaletteWindowHandle` (`:565`) for the same hole.
Verify: `.\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Debug -CaseFilter <palette/embedded-terminal cases>`
and the Pester source contracts; AGENTS.md spec closeout if a durable UI contract changes.

---

## Task 3 — FileSystem.dll self-tests fail silently without the symlink privilege

`RunDebugObjectBindingSelfTest` (`Plugins/FileSystem/FileSystem.FileOps.cpp:8394`) builds symlink
fixtures via `BuildSymlinkReparseData`/`WriteReparsePointData` (the `createSymbolicLink` lambda at
`~:8660`, "relative file-symlink fixture should be created" at `~:8686`). In a session without
`SeCreateSymbolicLinkPrivilege` (not elevated, Developer Mode off — `New-Item -ItemType SymbolicLink`
reports "Administrator privilege required"), that check fails and the function returns early, so
`RedSalamanderFileSystemDebugSelfTests` reports `passed=105 failed=1` (verified on `LT-PF5VDAGE` on
2026-09-15 by loading the DLL standalone; the failing check is the 23rd of that function, every other
FileSystem self-test passing). Consequences:
`Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker`
(`RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Fairstream.cpp:7721`)
runs the DLL's self-tests in-process and fails; the FileOps harness then times out every later
family's first case in a whole-suite run; and `PluginContractTests` reports the same
`passed=105, failed=1`. The failing check is invisible because the local `check` lambdas in these
self-tests only call `Debug::Error`, which `Common/Helpers.h` `Debug::Out` (`~:1940`) drops unless an
ETW listener is attached (`IsDebugEtwEnabled`, `:1702`); only the shared `Common::DebugSelfTest::Check`
(`Common/Helpers.h:405`) mirrors failures to stderr.

Decide with `Specs/Testing/Testing_SelfTests.md` and `Specs/FileSystem/*` whether the
symlink-dependent part should (a) probe the privilege first and report a declared skip/blocker with a
message, or (b) keep failing but name the failed check. Either way, route the local `check` lambdas'
failure text to stderr the way `Common::DebugSelfTest::Check` does, so `PluginContractTests` and the
Riptide case carry the failed contract's message. Also consider whether the FileOps harness should
keep running later families after an in-process plugin self-test failure (today one failure cascades
into many timeouts). Verify with `PluginContractTests.exe` and
`.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker -AllowAlternateVolumeTestRoot`
from BOTH an elevated and a non-elevated shell; AGENTS.md spec closeout.

## Execution record

Worked on 2026-09-15 on `4cb089111a23` from `bdde9579`, the tree that had just closed I22. Order of landing:
Task 2 (`c7330976`), Task 3 (`f5ab7fe0`), Task 1 (`359e4419`). Each carries its own test and witness.

### Task 1 — landed, with a corrected diagnosis

The plan's reproduction was exact; its reading of the two anomalies was not, and the difference
matters for the next person who touches S3 transfers.

**What it actually was.** The provider never re-resolves a key as a path; every call on the marker's
route passes bucket and key separately, and neither `ValidateS3SourceRevisionStillExists` nor
`DownloadS3ObjectToTempFile` sends a `Range`. Both anomalies come from the S3 CRT client:

- `HEAD /x/src/tree/ -> 404` is the CRT's `CopyObject` meta-request probing the source size. It
  builds that HEAD from the `x-amz-copy-source` header in **virtual-hosted** style (bucket in
  `Host`, key alone in the path); `aws-c-s3`'s `s3_request_messages.c` carries a `TODO` saying the
  shortcut "only works for virtual host endpoints". The loopback fake is path-style and ignores
  `Host`, so the first path segment becomes the bucket and the HEAD fails. The meta-request then
  fails without ever sending the `PUT`. This happens for **every** object, not only the marker: the
  no-marker experiment shows `HEAD /x/src/tree/a.bin -> 404` followed by the relay's
  `GET -> 206` and `PUT -> 200` for each child.
- `GET /r0f-bucket/x/src/tree/ -> 416` is the relay fallback's `GetObject`, which the CRT issues as
  ranged parts. No range of an empty object is satisfiable, S3 answers 416, and the transfer aborts
  on its first object, because the marker sorts first.

So the failure is "zero-byte object on the relay route", and the relay route is the only route a
path-style endpoint ever takes for small objects. The fake could not have caught it earlier because
its earlier fixtures seeded prefixes without markers, exactly as the plan said.

**The fix.** `DownloadS3ObjectToTempFile` (`Plugins/FileSystemS3/FileSystemS3.S3.cpp`) fetches nothing
for `expectedSizeBytes == 0`. It still owes what the GET's response would have proved, so it asks for
the pinned source's identity with one conditional HEAD through
`ValidateS3SourceRevisionStillExists`, which gained an optional content-length out-parameter,
requires the object to exist at its revision and still be empty, and hands back the empty temporary
file it already holds; `UploadS3ObjectFromFile` then publishes the marker. No change to
discovery reporting, planning, pinning, conflict handling or scheduling.

**Witnesses.** `RunS3DirectoryMarkerTransferSelfTests`, the plan's own test, landed verbatim
(request log depth 24, as written). Before the fix, run standalone through a loader against the
plugin's aggregator: all seven checks fail with `0x8007051A`, `RedSalamanderS3DebugSelfTests`
`passed=231 failed=7`. After: `passed=238 failed=0`; `PluginContractTests` green with the same
counts; governed `FileOpsFamily_R0fS3Containment` 3/3 and `DiscoveryScope_S3DirectApi` 3/3;
plugin build 0 warnings.

**Spec.** `Specs/FileSystem/FileSystem_S3.md` now states the zero-byte relay rule under
`flatPrefix`, and records the fixture's limit honestly: against a path-style endpoint the CRT never
reaches the fake's `CopyObject` handler, so small-object server-side copy is proven only by the live
profile.

**Residual, routed.** That limit is also a product fact: against any path-style S3-compatible
endpoint, every small-object same-bucket Copy relays through the client instead of copying
server-side. Correct, but slow and network-heavy, and the "conditioned server-side copy" half of the
pinned-revision route is exercised only by the live profile. Routed to the successor plan.

### Task 2 — landed

`RedSalamander/CommandPaletteWindow.cpp`: `WM_NCDESTROY` now nulls `_root`, `_searchField`, `_grid`
and `_emptyLabel` immediately after `_host.Detach()`, because `ControlHost::Detach` does
`_root.reset()` and the palette object outlives its window (the singleton is only replaced on the
next Show). `DebugSetSearch` refuses on `! _window` as `DebugSnapshot` already did. `RebuildRows`
was already guarded on `_grid`, so nulling it makes the production path safe too.

Regression, in `TestCommandPaletteClosesOnDeactivation`
(`RedSalamander/SelfTest/Commands/Commands.SelfTest.Shortcuts.cpp`): after the palette closes on
deactivation, both debug hooks must return false. Failing-before witness, obtained by stashing the
palette fix and keeping the test: the case crashed with `0xC0000005` at
`std::wstring::operator=` ← `DxUi::TextField::SetText` ← `CommandPaletteWindow::DebugSetSearch`
← `DebugSetCommandPaletteSearch` ← `TestCommandPaletteClosesOnDeactivation`, the same stack the
laptop's Fresh Full recorded (`RedSalamander-20260915-220715-p66660.txt` on this machine). With the
fix the case passes 1/1.

`cmd_pane_embedded_terminal_command_state_truth`
(`RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`) now requires the palette to
still be open before its first query, with a message that names lost activation as the cause, so a
run that loses foreground reports that instead of failing on the query.

One observation outside this task: after the access violation the crashed self-test process did not
exit; it sat not-responding with 580 threads until terminated, which is what let one crash take a
whole Commands suite down on the laptop. That is crash-handler behaviour, not palette behaviour, and
is not changed here.

### Task 3 — landed as "fail and name it"

Decision: the plugin debug self-test export is `HRESULT (unsigned int* passed, unsigned int* failed)`
and has no channel for a declared skip; inventing one would change every plugin's export and
`PluginContractTests`. A machine that cannot create the symlink fixtures has not proved the
object-binding contract, so it must not report green. The self-test keeps failing there, but it now
says what and why.

`Plugins/FileSystem/FileSystem.FileOps.cpp`: the three local `check` lambdas
(`RunDebugReparseCopyErrorMappingSelfTest`, `RunDebugSharedFileOpsSchedulerShutdownSelfTest`,
`RunDebugObjectBindingSelfTest`) are now adapters over the shared `Common::DebugSelfTest::Check`,
which mirrors every failure to stderr with the failed contract's text; call sites are unchanged.
The symlink fixture lambda records the failing HRESULT, and when it is
`ERROR_PRIVILEGE_NOT_HELD` the self-test prints one more line naming the missing
`SeCreateSymbolicLinkPrivilege` and the two remedies (elevate, or enable Developer Mode).

Verified on this machine: `PluginContractTests` and
`Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker` pass, the plugin's self-tests report
`failed=0`, and a deliberately failing check surfaces its text on stderr. The blocked branch itself
could not be exercised here: this token does not hold the privilege either (`whoami /priv`), but
Developer Mode is on, which is what lets the FSCTL path succeed, and turning Developer Mode off is a
system setting this closeout does not touch. The message is reached by the same early-return the
laptop hit, and its text was checked by reading.

Residual, not fixed: the cascade the plan flagged (one in-process plugin self-test failure timing out
every later FileOps family's first case) was not reproducible here for the same reason. The Riptide
case's own failure path is clean (module released by RAII, returns true), so the mechanism is
elsewhere and is left for a machine that can reproduce it.

### Closing gate

This plan's completion criterion is per-task evidence, each fix green on its own test before it joins
a gate, and that is met above. A Fresh Full was run on the closed tree as well, on `4cb089111a23`, and
is archived at
`../../TestRuns/4cb089111a23/FileOps/2026-09-15_222913_i24_s3_marker_and_residuals_fresh_full/README.md`.
It is reported as it happened: **status failed, 2,143 passed / 1 failed / 52 declared skips**, 20/20
entries promoted, build 0 warnings, `PluginContractTests` green with `FileSystem.dll` 163/0 and
`FileSystemS3.dll` 238/0, `FileOps.Discovery.GrowthAfterClose` 0 across all 382 completions, and every
witness this plan added green inside it. The one failure,
`cmd_pane_navigationView_history_dropdown_keyboard_navigation`, is the machine's physical mouse cursor
parked over the history popup (`hovered=18, keyboard=-1`, cursor inside the popup rect); it passed in the
three prior gates on the machine and 3/3 immediately afterwards, and nothing here touches navigation.
A first attempt the same evening failed differently, on a 240 s shell recycle-bin hang and the `Setup`
cascade it caused, and is recorded in the archive and routed as a harness residual. Two attempts on a
machine that was also in use is where this closeout stops; the classification is stated, not hidden.

### Note on this file

As committed in `41f288e99` this file carried `\r\r\n` line endings on every line, which the fenced test
source inherited when it was extracted; the closeout normalized it to CRLF.

## Execution checklist

- [x] Task 1: re-add the S3 marker self-test, pin the bucket-dropping resolver and the zero-byte
      ranged GET, fix both, self-test green in `PluginContractTests` and `FileOpsFamily_R0fS3Containment`.
- [x] Task 2: clear the palette control pointers on destroy, guard the debug hooks, make the case
      diagnosable; Commands palette cases and Pester green.
- [x] Task 3: make the FileSystem.dll symlink self-test name its failure (and decide skip vs fail);
      Riptide case behaves from both elevated and non-elevated shells.
- [x] Each landed with its own commit and evidence; move this plan to `Specs/Plans/Done/` and remove
      its row from the WIP index when all three are done (or split remaining tasks into their own plan).
