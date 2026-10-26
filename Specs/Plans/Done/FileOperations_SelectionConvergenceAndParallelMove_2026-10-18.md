# I42 Copy/Move selection convergence and folder parallelism

Status: DONE for the requested Copy/Move implementation and focused x64 qualification. Completed on 2026-10-18 against `a48e7d788` plus the source snapshots in the archived receipts, preserving unrelated worktree edits. This archive does not claim a green repository-wide PR or Full result.

## Checklist

- [x] Trace the two execution layers and fresh-discovery ownership; confirm the managed folder Move serial branch and Auto/manual provider inconsistency.
- [x] S0: add deterministic selection-shape regressions and retain unchanged-runtime baseline, including the configured real SMB resource.
- [x] S1: use one resolved Copy/Move concurrency policy for selected roots and directory descendants; enable proof-gated parallel managed folder Move with post-order cleanup.
- [x] S2: make fresh selected-file counting independent of transfer admission, bounded, cancelable and reliable; preserve directory traversal ownership and honest unknown/error states.
- [x] S3: validate Copy/Move, folder/selected files, nested/empty/large selections, cancel/skip/failure/retry/source preservation and destination contents; measure actual overlap and discovery closure.
- [x] S4: build and run affected Debug/Release validation and real disk-to-SMB matrix; archive source/profile-bound baseline/candidate evidence and limitations.
- [x] S5: update authoritative specs/docs, review changes, complete scoped closeout accounting and archive validation; move the completed scope to Done.

## User contract

The configured effective file budget (four in the requested scenarios) applies equally to one selected folder and its selected leaf files for Copy and managed Move. This is a file concurrency ceiling, not a promise that a short file or a provider with a narrower admitted limit remains simultaneously active. Explicit Manual one, provider/session limits, verification policy, cancellation and safety remain honored.

Counting must converge from fresh provider metadata, never from cached pane sizes. A selected-file list must not wait for preceding payload copies to finish before its later members are counted. Workers still bind/revalidate exact sources before mutation; counting confers no delete/publication authority. Unknown size, missing/changed source and enumeration failure must stay honest, and retries must not double count or reopen a frozen denominator.

Managed directory Move keeps exact source authority until every descendant has published/proved/cleaned up successfully. Complete directory listings before removing their children. Keep ancestor directories on skipped/failed/retained/canceled descendants. Release directory metadata and authorities only after all users finish; join workers before stack-owned state is destroyed. Do not simply remove the serial guard without extending parallel completion ownership.

## Implementation scope

- `RedSalamander/FolderWindow.FileOperations.State.cpp` and its private/task declarations: fresh selected-root metadata accounting, duplicate-report handling, bounded worker/producer lifetime, parallel directory completion and exact cleanup.
- `Plugins/FileSystem/FileSystem.FileOps.cpp`: Auto versus Manual concurrency resolution for recursive provider work.
- Existing File Operations discovery/SMB selftests and their runner registrations. Reuse sandbox, machine-resource, activation-blocker, hashing/content and progress observers. Extend governed noninteractive case admission only if required; read tooling governance first.
- Authoritative `Specs/FileSystem/FileSystem_FileOperations.md`, provider/popup/testing contracts and `docs/FileOperations.md`.

## Validation and evidence

Use `./Test-Changes.ps1 -Explain` and affected iteration. Build through `build.ps1`; launch all runtime tests through the governed runner beneath the existing marked `Z:\RedSalamander.Perf` sandbox. The enabled `path.smb.primary` resource is `\\rubygloom\RedSalamander-SelfTest\selftest`; use only its harness-owned child directories and cleanup. Never use user files as fixtures or terminate an independently launched application. The ordinary Release output is currently occupied by a user instance; use an unoccupied profile or isolated checkout for Release qualification if it remains occupied.

Baseline and candidate run the same deterministic fixture/scenario. Record configured/effective budget, maximum actual active files, fresh count/bytes, discovery closure wall time and payload bytes completed at closure, growth after close, provider directory enumeration counts, terminal outcomes and independently checked destination bytes/content. Require four actual overlapping nonempty files for each admitted four-file scenario and early stable discovery for selected leaves. Small-file/wide/nested cases retain bounded queue/memory behavior. Keep failed baseline evidence. Record skips and external prerequisite failures explicitly; a skipped SMB matrix is not real SMB qualification.

Required matrix: real disk-to-SMB Copy and managed Move x one folder and the equivalent selected leaves, plus native/host bridge Copy and local Copy controls for Auto with inactive Manual one and Manual four/one. Include known-empty files, nested/empty directories, large fresh selected-file counting, conflict skip/retry, cancellation and source cleanup safety. Same-root native Move remains a single identity-bound rename rather than a payload-copy workload. No desktop activation is needed for these storage assertions.

Archive compact evidence under `Specs/TestRuns/<MachineHash>/FileOps/2026-10-18_selection-convergence-and-parallel-move/`, linked to raw governed sandbox run IDs and receipts. Validate curated and inventory archives. Do not claim Full, ASan, ARM64 or ordinary-host visual qualification from focused x64 evidence.

## Implementation closeout

Auto now honors the host's nonzero resolved file budget instead of accidentally applying the inactive saved Manual value. Manual limits and provider/session limits remain effective. Managed directory Move uses the same bounded parallel file queue as Copy, with exact cleanup authority retained by the last-finisher directory completion. A complete directory listing precedes child cleanup; failed, skipped, retained or canceled descendants retain their ancestors. All jobs join before shared state is released.

Built-in local-disk selected files receive one fresh, cancelable metadata pass independent of transfer admission. No pane count/size cache is reused. The drive-root classification is cached only for scheduling eligibility; UNC, mapped-network and unknown roots retain their existing producer I/O policy. Directory roots retain one traversal owner. Fresh leaf reports are consumed once, size/kind changes withdraw certainty, and both Copy and Move publish byte certainty before notifying closed discovery. Durable contracts are merged into `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/UI/UI_FileOperationsPopup.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Testing/Testing_PerformanceValidation.md` and `docs/FileOperations.md`.

## Evidence and qualification boundary

The complete evidence index is [the same-machine FileOps archive](../../TestRuns/4cb089111a23/FileOps/2026-10-18_selection-convergence-and-parallel-move/README.md). It retains baseline/candidate metrics, exact build receipts, failed controls, raw sandbox paths, final post-closeout validation and archive validator results. Final receipts are recorded there rather than changing this frozen plan and invalidating the source snapshot again.

Debug local run `20261018T151723Z-72804-1e10d66f73d24e09b30064eb9bb7021c` and real-SMB run `20261018T144040Z-65212-10f906fe629842d683ccd684d6e01009` passed. Release local run `20261018T151915Z-72804-b47b30de57134680afaf42094f408c59` and real-SMB run `20261018T150546Z-64152-02235241de524e888d7cc640fcd78821` passed. Their exact receipts and subsequent local-drive/publication-guard runs are retained separately. All six SMB rows reached four actual file transactions; exact destination contents and source outcomes passed, with no skips. All five cleanup controls and the thirteen-file/six-directory nested Move passed. The NAS refuses exact folder-metadata rebind; the test requires an explicit unsupported-metadata diagnostic rather than claiming preserved timestamps.

Fresh counting of 4,097 local selected files closed in 15 ms in the retained Release treatment, before payload bytes, against a two-second acceptance budget. This is same-machine evidence for the bounded fixture, not a universal latency guarantee. Debug remediation, small-file cost and multiple-folder discovery regressions also passed. The exact-case no-activation admission, I42 Done-path admission and sibling rejection passed in the broader tooling run.

The broader Debug PR run passed twelve native/standalone entries and 951 Pester tests, but failed one preexisting frozen Terminal workflow checksum test. The clean-HEAD workflow contains dependency-bot action-version changes relative to its sealed checksum. This task changes neither the workflow nor its guard. PR remains failed and Full remains NOT_EVALUATED. PrePush explanation retains five interactive obligations; foreground, ARM64 and ASan are not qualified by this work. The user's independently running ordinary Release instance was left untouched; Release evidence uses the attached isolated checkout.

The separate legacy 350-file/two-selected-root SMB cost diagnostic timed out after 165 payloads and remains unqualified; the bounded nested case does not replace that gate. The sandbox audit retains 41 preexisting unowned historical directories. One intermediate timestamp-failure raw directory was removed by the next runner before preservation; the archive labels its surviving transcript/receipt-only evidence explicitly. Subsequent runs are copied to persistent evidence before launching the next runner. Earlier failures and the recovered PDB/RPC build failure remain recorded rather than represented as passes.

## Historical baseline and intermediate controls

Real SMB baseline `20261018T125900Z-29468-440d2b8db20b4c56948f42fabb0cb2dd` retained under the marked sandbox's `evidence/selection-convergence/` directory. With eight 4 MiB files and one empty file, native and bridge Copy reached four concurrent files, managed folder Move reached one and equivalent selected files reached four. All destination bytes and source outcomes matched. Selected-file discovery closed after all 32 MiB had already been copied: native Copy 9.969 s, bridge Copy 13.484 s, Move 14.062 s. Folder Move took 34.250 s versus selected-file Move 19.342 s. These are same-machine treatment baselines; discovery and overlap correctness are the primary acceptance predicates. The initial partial-byte-only bridge observer was invalid for single-chunk files; retain it separately, never cite its zero as a concurrency result. The baseline also exposed a fixture-only normalized-configuration backup comparison, corrected in the candidate.

Candidate targeted x64 Release compilation succeeded with zero warnings/errors (receipt `a18a60b88468eee8aa726225b3ebb05c71383b339feea51df5049aa8d30c327b`); later source-bound builds superseded that compilation-only result. Full baseline x64 Debug build succeeded with zero warnings/errors (receipt `dfd02b4474c9c3902b0773123a6168d87c1fdd5e00ad081688d30fe1ab26128f`).

Initial candidate Local run `20261018T134540Z-45164-3886259006ac48a19f50f723506ed8c0` passed all six concurrency rows and live selected-file totals; 4,097 fresh files counted in 39,770 us (discovery closure 47 ms), before payload bytes. Duplicate/changed-size controls and canceled source preservation passed. SMB run `20261018T134757Z-13240-85373ad0495e4d3ea3250e6adacd71a1` passed concurrency/content/source predicates but exposed missing Move byte-certainty publication to the popup; the coordinator now publishes it for both Copy and Move.

Corrected Debug SMB run `20261018T140042Z-102436-50a64d8401204965beb6d961cbae33d2` passed all six matrix rows, including live Move totals and four concurrent folder Move files (15.848 s treatment). Its cancellation fixture fired with four permits but only one confirmed pre-cleanup pause, while another file still published; a retained stage made the clean-destination predicate fail. The fixture now counts four actual pause waiters and holds all publications before injecting cleanup outcomes. Preserve the failed run; it is not a safety pass. The unrelated legacy 350-file/two-selected-root SMB diagnostic reached 165 payloads before its 300-second timeout, then canceled/drained; retain that diagnostic as an unqualified larger SMB cost scenario. I42 uses a bounded single-root nested witness with thirteen files and six directories, exact contents/timestamps and source removal; large fresh selection coverage remains 4,097 files locally.

Corrected x64 Debug receipt `985cbe9383f6b885c50d98e206247daabd8ddd0a681b84fce6f053a47a1fceaa` and x64 Release receipt `17a23818b42f51a18a4603d0392e7c01ed174ba9f6c3024922fe3abdad4a182f` built with zero warnings/errors. A preceding Release attempt failed C1090 (PDB RPC unavailable); its log/hash is retained separately. Later source-bound builds and passing focused runs are indexed above and in the evidence archive.
