# Operation Curl Short-2xx Staged Upload Proof

**Status:** DONE - deterministic Debug/Release proof, mutation sensitivity, production-export verification, and archive complete
**Date:** 2026-07-21
**Owner row:** `FG-P0-3 proof depth` in `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`
**Origin evidence:** `Specs/Reviews/FS-Subsystem-Audit.md` and `Specs/Reviews/FS-Subsystem-Findings.md` (historical snapshots)

## Objective

Add deterministic, real Curl transport coverage for a remote endpoint that stores fewer bytes than requested while
returning a successful transfer response. Prove that the existing staged-upload verification rejects the short object,
deletes the staged remote object, preserves the MOVE source, leaves the final destination unmodified, and returns
`HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)`.

This is proof-depth work for an existing data-safety boundary. It must not replace the targeted remote size probe,
weaken fail-closed behavior, introduce a production-only test bypass, or add I/O to the ordinary production path.

## Last Finding and Current Spec/Code State

### Start checkpoint - 2026-07-21

- `CopyFileViaTemp(...)` uploads to a generated sibling staging path through `CurlUploadFromFile(...)`, then calls
  `CurlProbeRemoteFileSize(...)` before promotion. A known size mismatch deletes the staged object and returns
  `ERROR_PARTIAL_COPY`; a failed targeted probe falls back to the listing-based entry check and applies the same rule
  when that listing reports a known size.
- `CurlUploadFromFile(...)` currently returns `S_OK` when `curl_easy_perform(...)` returns `CURLE_OK`. The staged re-stat
  is therefore the independent guard against a server/proxy that reports protocol success after retaining only a
  prefix of the declared upload.
- `FileSystemCurlTests` has parser and fake IMAP command coverage but no fake FTP/SFTP/WebDAV transport capable of
  accepting a real libcurl upload and lying about the retained byte count. The exact short-success path is not
  deterministically RED-able today.
- The lasting integrity rules already exist in `Specs/FileSystem/FileSystem_FileOperations.md`; the aggregate owner
  remains open in `Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`. No production code has changed for
  this slice yet.
- `Specs/Reviews` is intentionally historical rather than WIP. Its new folder README documents that accepted findings
  must be routed to a dedicated WIP owner such as this plan instead of turning a mixed audit snapshot into a live queue.

### Design checkpoint - 2026-07-21

- The required shared-helper inventory is complete. `Specs/Core/Core_SharedHelpers.md`, `Common/`, and
  `Tests/TestSupport/` contain no reusable FTP/loopback protocol server. Microsoft Drive has a private OAuth redirect
  listener, but its single-request HTTP semantics and ownership do not match a multi-connection Curl FTP endpoint;
  duplicating or widening it would create the wrong dependency boundary.
- The selected seam is a Curl-local, `ENABLE_TESTS`-only fake FTP endpoint using WIL-owned sockets, ephemeral IPv4
  loopback ports, bounded `select`/socket timeouts, and an owned `std::jthread`. No production behavior or request shape
  changes are required.
- The selftest will create two endpoints with different ports and call the public FTP `MoveItem(...)` API using authority
  paths. Different ports force the real cross-connection path: source LIST/RETR, destination LIST/STOR/SIZE/DELE or
  RNFR/RNTO, followed by source DELE only after successful promotion.
- The destination endpoint has two deterministic modes: retain a strict prefix but reply `226` after consuming the full
  STOR stream, and retain the complete stream. This makes libcurl return `CURLE_OK` in both cases while the existing
  independent SIZE guard distinguishes them.
- Existing Curl selftests are `_DEBUG`-gated. This slice will align only those selftest declarations/definitions/exports
  with the repository-wide `ENABLE_TESTS` gate so test-enabled Release runs the same proof while production Release
  continues to omit the export.

### Historical resume checkpoint - 2026-07-21

Work is intentionally paused before any code implementation. Everything needed to resume is persisted here.

- Working branch: `codex/curl-short-2xx-proof`.
- Documentation already changed: root `README.md`, new `Specs/Reviews/README.md`, the WIP routing index, the aggregate
  Floodgate plan, and this dedicated plan. No Curl source/project file has been modified yet.
- First code milestone: add an `ENABLE_TESTS`-only Curl selftest translation unit containing the bounded fake FTP
  endpoint. Give it initial `/source.bin` bytes, record commands under a lock, retain either a strict upload prefix or
  the complete STOR body, and implement the minimum real libcurl command set (`USER`, `PASS`, `PWD`, `SYST`, `FEAT`,
  `TYPE`, `CWD`, `EPSV`/`PASV`, `LIST`, `SIZE`, `RETR`, `STOR`, `DELE`, `RNFR`, `RNTO`, `QUIT`).
- Ownership/teardown milestone: call `WSAStartup` once for the selftest scope with `wil::scope_exit` cleanup; own every
  listener/client/data socket with `wil::unique_socket`; bind `127.0.0.1:0`; use bounded `select` and socket timeouts;
  run each endpoint in an owned `std::jthread`; request stop and join before socket/WSA teardown. Catch only named
  exception types at the thread/selftest ABI boundaries and terminate on `std::bad_alloc`.
- Public-path milestone: create the FTP provider through its factory, configure short finite connection/operation
  timeouts and concurrency 1, then call `MoveItem(...)` between authority paths on the two different ephemeral ports.
  The short case must return `ERROR_PARTIAL_COPY`, leave source bytes unchanged, issue one successful staged SIZE,
  delete the generated staging object, and leave the final path absent. Switch only the destination retention mode for
  the control case; it must promote exact bytes and delete the source.
- Project/gating milestone: add the selftest source and `Ws2_32.lib` to `FileSystemCurl.vcxproj`; expose a narrow
  test-only factory helper if direct COM creation would otherwise require new raw ownership; change only Curl selftest
  guards/declarations/exports from `_DEBUG` to `ENABLE_TESTS`; invoke the new proof from
  `RedSalamanderCurlDebugSelfTests`.
- Validation milestone: run focused Debug first; prove mutation sensitivity by temporarily bypassing/inverting the
  staged-size comparison, record the expected failure, and restore it; run test-enabled Release; separately verify a
  production Release has no selftest export; run relevant PluginContract/source-contract/full gates; archive timings,
  bounded command counts, logs, and mutation evidence under `Specs/TestRuns/`.
- Closeout milestone: merge the lasting deterministic fault-injection requirement into the authoritative File
  Operations/testing specs, mark aggregate FG-P0-3 complete, reconcile `Specs/Plans/WIP/README.md`, and move this plan
  to `Specs/Plans/Done/` only after every Done-gate item below is evidenced.

### Final closeout - 2026-07-21

- Added `FileSystemCurl.SelfTest.Ftp.cpp`, an `ENABLE_TESTS`-only loopback FTP fixture with WIL-owned sockets,
  ephemeral ports, bounded waits, owned `std::jthread` teardown, and the minimum real Curl command surface. The
  selftest creates the FTP provider through a test-only factory helper and drives the public cross-connection
  `MoveItem(...)` path.
- The short endpoint consumed the complete 4,099-byte upload, retained 2,049 bytes, and replied `226`. The production
  staged-size guard returned `0x8007012B`, removed the staged sibling, retained the exact source, and preserved the
  existing final destination. The control retained/promoted 4,099 bytes, removed its rollback sibling, and deleted
  the source only after promotion.
- Debug and test-enabled Release `PluginContractTests` passed with Curl provider assertions 102/102. The standalone
  Curl harness passed in both configurations, and relevant source contracts passed 149/149. Captured bounds were
  20 ms / 15 destination commands for the short case and 30 ms / 27 commands for the control.
- Temporarily weakening `stagedProbeSize != fileSize` to `stagedProbeSize > fileSize` produced the required RED result:
  97 passed / 5 failed, `E_FAIL`, and harness exit 1. Restoring the guard returned 102/102 and exit 0.
- The production tests-disabled Release provider rebuilt with 0 warnings/errors, exported zero
  `RedSalamanderCurlDebugSelfTests` symbols, and passed `PluginContractTests` with the optional Curl selftest reported
  as an explicit skip. The prior full-solution Release recovery build's 10 warnings were pre-existing
  `ViewerPETests` unused-symbol warnings and did not involve FileSystemCurl.
- Authoritative behavior is merged into `Specs/FileSystem/FileSystem_FileOperations.md` and test policy into
  `Specs/Testing/Testing_SelfTests.md`. Compact results and metrics are archived under
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-21_220148_curl_short_success/`.

## Protected Scenario and Evidence Contract

Protected scenario: a cross-connection Curl MOVE whose upload receives a successful protocol result even though the
remote staging object contains only a strict prefix of the source. Distinct endpoint ports are required so FTP cannot
take its same-connection server-side rename path.

Required deterministic assertions:

1. The fake transport receives the real Curl upload path and intentionally retains fewer bytes than the declared
   source size while returning the protocol's success response.
2. The staged size probe observes the retained prefix size exactly once on the non-transient success path.
3. The operation returns `ERROR_PARTIAL_COPY`.
4. The original source path and bytes remain unchanged.
5. The generated staged remote object is deleted and the final destination is not created or overwritten.
6. A control case with a complete retained upload still promotes successfully and, for MOVE, removes the source.
7. The short-success test fails when the staged re-stat guard is deliberately bypassed or inverted.
8. Production Release exposes no debug-only selftest entry point, and ordinary production request shape is unchanged.

This task makes no runtime performance-improvement claim. Validation still records bounded fake-server command counts
and elapsed time so the new deterministic test cannot introduce an unbounded retry or timeout path. Because the
production hot path is intended to remain unchanged, no before/after application throughput comparison is required
unless implementation discovery finds that production logic must change.

## Implementation Boundaries

- Prefer an existing shared loopback/test-server helper if its ownership and protocol semantics match. Search
  `Specs/Core/Core_SharedHelpers.md`, `Common/`, and `Tests/TestSupport/` before adding a local server.
- If a Curl-specific fake endpoint is required, keep it in the Curl test dependency layer, bind only to loopback on an
  ephemeral port, use WIL RAII for sockets/handles, bound every accept/read/write wait, and make teardown deterministic.
- Exercise the actual plugin upload, probe, delete, and promote requests. A pure model test that calls a helper with a
  fabricated size is insufficient for FG-P0-3.
- Do not add credentials, external-network dependencies, sleeps without cancellation/deadline bounds, or a production
  feature flag that can suppress integrity verification.

## Progress Ledger

- [x] Updated the root README with the child-process TestSandbox isolation rule discovered during the previous closeout.
- [x] Added `Specs/Reviews/README.md` explaining why historical audit snapshots are not WIP plans and how findings route.
- [x] Created this dedicated WIP owner and recorded the current production/test state before implementation.
- [x] Inventory shared test helpers and the Curl test/project linkage; select the smallest deterministic real-transport seam.
- [x] Add the deterministic short-success case and preserve a RED mutation result/trace with the staged-size guard
      deliberately bypassed or inverted.
- [x] Implement the bounded fake transport or narrowly scoped test seam with RAII and explicit teardown.
- [x] Prove short-success rejection, source preservation, staged cleanup, final-path preservation, and control success.
- [x] Run Debug and test-enabled Release focused coverage plus relevant provider/source-contract gates.
- [x] Archive compact results, trace, command counts, and timing under `Specs/TestRuns/<MachineHash>/FileOps/<RunId>/`.
- [x] Merge lasting validation requirements into the authoritative File Operations/testing specs.
- [x] Mark FG-P0-3 complete, reconcile the WIP README, and move this plan to `Specs/Plans/Done/`.

## Done Gate

Move this plan to `Specs/Plans/Done/` only when the real Curl transport test is deterministic in Debug and test-enabled
Release; the required source/final/staged-object assertions and mutation sensitivity are proven; bounded timing/request
evidence is archived; production Release contains no test export; relevant source contracts and focused provider tests
are green; every unrelated broad-gate failure is classified with durable evidence; the authoritative specs own the
lasting contract; and the aggregate Floodgate row plus WIP routing index are reconciled.
