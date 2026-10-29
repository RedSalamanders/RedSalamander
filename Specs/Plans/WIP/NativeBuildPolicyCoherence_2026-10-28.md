# I50 Native build policy, terminal results and short DxUi paths

Status: ACTIVE
Date: 2026-10-28
Owner: build maintainers

The user's agreed policy is /Wall with /WX disabled and zero warnings accepted;
Debug uses EditAndContinue, ASan Debug uses ProgramDatabase. Release enables
explicit COMDAT folding (/OPT:ICF); Debug and ASan Debug disable it.
The owning build contract is [the toolchain contract](../../Build/Build_Toolchain.md).
Five repositories implement the same policy without implicitly advancing DxUi pins.

## Checklist

- [x] Preserve the pre-change effective-profile audit and isolate changes from concurrent tooling work.
- [x] Restore RedTilde ARM64 through its normal entrypoint and pass the required policy/lease checks.
- [x] Enforce and negatively test the common compiler/debug/linker policy in every supported native profile and pinned archive.
- [x] Unify green zero-warning build results, diagnostic/log links and opt-in terminal executable hyperlinks, preserving plain text and existing build/run semantics.
- [x] Shorten source-commit and output-fingerprint folders to 16 hex characters with full identity validation and collision refusal; retain existing data without cleanup.
- [x] Update source discovery, direct helper builds, provenance, test fixtures and dependency identities together.
- [x] Publish zero-warning guidance and normative build/path/terminal contracts, with meaningful tooling tests.
- [ ] Execute affected tests, evaluated profiles, required native builds and applicable runtime qualification; retain every failed result and distinguish cross-builds from runtime tests.
- [ ] Resolve the reproduced historical C11 extra passive announcement from bounded per-session protocol evidence, retaining the current native Copy fixture's listener, contents, completion and deadline predicates until the origin is proved.
- [ ] Fold/check the changelog; validate formatting/specs/tooling and close this plan only after all required gates pass.

## Decisions and limits

Source pins and full SHA256 identities remain authoritative; folder prefixes are
locators. No source clone is reset for a prefix collision. Existing historical
output is retained in place. Build success is green only after successful
native exit, zero diagnostics and required attestation. Hyperlinks never launch
an application automatically and fall back to quoted launch commands.
The existing explicit RedSalamander Embedded compiler-debug recovery mode remains
an attested exception; ordinary Debug/ASan profiles use the shared defaults.

## Evidence

The [curated Curl diagnostics](../../TestRuns/Local-x64/CurlFixture/2026-10-29_101838/README.md)
retain complete original metric streams, result records and digest/size/row-count
bindings for the real fallback witness, 310-check Entry run and managed-host Copy
timeout. Explicit pre-stage archive validation passes all eight files. These are
diagnostics, not fresh Full application evidence or a paired timing result.

The expanded Entry export passes 310 checks; the native FileOps caller's fixed
expectation must advance from 298 to 310 with those additions. Preserve Copy's
142-check expectation. A cancelled old-source queued test conservatively marked
the x64/ASan artifacts interrupted; the normal guard refuses reuse until a full
solution `build.ps1 -Rebuild` in exactly that profile clears the marker. Keep the
guard and marker intact and use that governed recovery before final runtime work.

The controlled `ae990475a` ASan Entry export passes 304/0. Its exact owned refused
port 60143 appears in both the libcurl failed-connect row and the same-session
`USER -> PASS -> PWD -> CWD probe -> EPSV -> PASV` retirement tail, followed by one
successful LIST. Candidate accounting now conserves accepted transfers and only
these classified fallback announcements; six additional controls preserve rejection
of unexplained retirements, missing/invented accepts and excess creations. The
original large-corpus deadline and content/completion predicates are unchanged.
Disable Nagle only on the fake FTP server's accepted control sockets to remove
implicit small-reply pacing; production networking and explicit fixture stalls are
unchanged. Fresh native/Full qualification and matched timing evidence remain open.

The original C11 route also records a cursor data-connect failure at 07:26:12.165 UTC,
inside its 07:22:52--07:26:39 window. Add a controlled real-libcurl EPSV refusal/PASV
fallback witness with exclusive port ownership and exact command/transfer counts.
This must pass before attributing a pending announcement to fallback; the large
Copy fixture's existing predicates remain unchanged. The second `1c1da3629`
diagnostic still exhausts the unchanged Wide ASan deadline after 4017/4097
publications under concurrent build/test activity. Both this failure and the
preceding 142/0 diagnostic are preserved; neither erases historical C11.

The diagnostic-only `edb2af23` direct Curl run exhausts the existing six-minute
ASan Wide/CopyItem/concurrency1 deadline after 3952/4097 publications, without
reproducing C11's pending passive retirement. Its listing-build metrics show
75.6 seconds in repeated destination serialization. Replace per-row general
formatting with fixed grammar and decimal conversion, preserving exact byte
oracles and adding the maximum-size boundary. This does not establish or waive
the separate C11 origin; native/focused/Full validation remains open.

Baseline audit: E:/DxUi.Perf/flag-coherence-20261028-aea0f188.
Current task evidence: E:/DxUi.Perf/build-policy-evidence-20261028.
RedTilde ARM64 restore and required policy/lease logs are under that task's
RedTilde-ARM64-restore folder; both exited zero before implementation.
Candidate dependency: `1f385ee013189e967ad5298f2311801a066de2c7` (API4). Its six native builds, all three ordinary x64 profiles and all three focus profiles passed; focus runs recorded zero capability skips and verified restoration. Task logs `cs28-*-qualification3.log`, `cs28-*-full2.log` and `cs28-*-focus6.log` retain those separate gates. Product validation still requires its normal restore, matching headers/archive, six builds and Fresh Full execution.

The unfiltered Fresh Full Debug Commands run on9a73 completed963 cases with one failure in
`cmd_pane_fileops_retained_source_and_mixed_journeys`. Its original compound recovery assertion had no per-field
diagnostics; the failed trace is preserved under this project's `Commands-retained-source-failure1` evidence.
Source review establishes that a live popup may show worker-finished before the completed summary is published.
The fixture now waits, within its existing deadline, for the exact summary before reading the completed popup
projection, then retains all original clipboard/identity/disposition/action assertions and traces their values.
No production behavior changes and no original runtime cause is claimed from the passing-source analysis alone.
Focused Debug/Release/ASan and final unfiltered Full qualification remain required; the current older-source Full
continues collecting its remaining entries before any build/output replacement.

The existing post-finished/pre-summary worker pause point now forces this boundary in the same real Move journey.
The test requires a finished live snapshot with no exact summary and no recovery action while the worker is held,
releases the existing pause through RAII on every exit, then waits for actual publication and preserves every
original identity assertion. This is a deterministic control of the premature-read condition, not a timing sleep
or a synthetic completed-summary substitute. The original uninstrumented failure's exact field remains unknown.

The same older-source Full run reported two tooling failures: the long-path restore fixture's fixed padding
produced only247 characters after source locators shortened, and the tooling README lacked the new path module.
The fixture now sizes its padding from the actual16-character locator and still requires a metadata path above260
characters; the README module table again covers the path helper. Final Full validation remains pending.
The first focused retry then reached the new policy identity and exposed missing policy files in this synthetic
restore root. The fixture now supplies the two actual first-party policy inputs and verifies their identity fields
alongside all original compiler-host/environment/API-mismatch checks; the failed retry is retained.
The repaired dependency and tooling-governance suites each pass12 cases with zero failures or skips in a new marked
sandbox run. This qualifies those tooling repairs; the native recovery case and final Full matrix remain pending.

The older-source Full FileOps run also failed its mandatory R4-A19 independent-volume control because the caller
omitted `-AllowAlternateVolumeTestRoot`. The selected primary sandbox already has an enabled resource naming
the existing marked C:/RedSalamander.Perf, on a different fixed NTFS volume. No new drive-root initialization is
needed. The scoped entrypoint now exposes and forwards the canonical opt-in without changing its disabled default.
Behavior tests execute the real wrapper for Full/focused FileOps with and without the flag. Final native and Full
qualification will use the normal explicit opt-in and the existing exact-root registry restriction.
The focused scoped-testing suite passes all9 cases with zero failures/skips, including four actual child-process
wrapper combinations. Its first attempt reached all four forwarding assertions but failed fixture cleanup by
calling a private helper directly; that failed receipt remains, and the corrected test uses its owning module
without expanding the production export surface. The canonical PR execution digest is refreshed and verified.


## Local qualification checkpoint, 2026-10-29

The current five-repository audit passes 504 native profiles and 73,502 evaluated setting groups. All four consumers
pass their 24 required pinned-archive profiles; the additional Tilde lease profiles and negative controls pass.
Exact source identities, raw metadata and field differences are retained under the task's final-audit-20261029
evidence. Ordinary Debug and ASan formats remain separate from the recorded RedSalamander Full Embedded recovery mode.

Source 6acf passes six native builds and all three focused recovery profiles with independently verified held/publication/action facts. Debug, ASan and the normal Release retry are active. The first concurrent Release Full attempt is retained as BLOCKED_ATTESTATION: another profile owned the shared-dependency scope. The retry was admitted after both other profiles completed their front-loaded artifact-mutating entries and re-attestation; runtime readers may overlap under the existing desktop lease.

The ASan Commands run found an additional first-frame test failure: Find help had synchronous layout/backdrop state but no dispatched paint when inspected. The test used one Commands pump turn, which deliberately yields after 10 ms. The candidate test now waits at most two scaled seconds for the actual first paint with the shared bounded helper, records elapsed time/message count on failure, and retains every visibility, opacity, scrim and chrome assertion. It neither forces a paint nor injects input. The original failed trace and source/executable hashes are preserved under the task's `Sal-6acf-ASan-Find-help-first-frame-failure1` evidence. Candidate native/focused/Full qualification remains pending; this timing diagnosis is not a claim that the old product renderer was defective.

The sealed terminal audit also found `tooling.pester` failed in all three profiles despite 1043 passed, zero failed and three skipped tests. The archive-inventory skip is already an allowed history-dependent contract. The two unallowed skips were the installer mirror test's stale full-commit source path and the missing `.build/vcpkg-tool` prerequisite. The mirror test now uses the shared source locator and checks the complete Git HEAD before import; normal DxUi restore and the exact pinned tool checkout provide the required real inputs. The strict skip policy is unchanged.

ASan's build-toolchain fixture also collided with another profile's still-active MSBuild property query in general Pester. Its guard correctly refused the overlap; the first front-loaded artifact-writer admission was insufficient because general tooling can evaluate MSBuild too. Preserve this failed receipt and run the next Full profiles serially. This is an orchestration repair, not a reason to weaken process or shared-dependency guards.

The repaired mirror/version-database suite executes all 29 cases with zero failures or skips after normal DxUi restore and exact pinned tool-source materialization. The isolated shared-runtime Clean suite passes all ten cases with explicit x64/ASan Debug context. Tool inventory validation and specification inventory pass, with zero blocking spec findings. A broader direct-Pester diagnostic in the unbuilt integration checkout is retained as failed setup: six reparse/link fixtures correctly refused its ambient TEMP/TestDrive outside the marked sandbox, and ProductUiTests was not built there. Final complete tooling must use the canonical runner's contained environment and attested native products; no guard or skip-policy change is justified by that diagnostic.
