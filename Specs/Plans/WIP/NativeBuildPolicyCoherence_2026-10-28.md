# Native build policy, terminal results and short DxUi paths

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
- [ ] Enforce and negatively test the common compiler/debug/linker policy in every supported native profile and pinned archive.
- [x] Unify green zero-warning build results, diagnostic/log links and opt-in terminal executable hyperlinks, preserving plain text and existing build/run semantics.
- [x] Shorten source-commit and output-fingerprint folders to 16 hex characters with full identity validation and collision refusal; retain existing data without cleanup.
- [x] Update source discovery, direct helper builds, provenance, test fixtures and dependency identities together.
- [x] Publish zero-warning guidance and normative build/path/terminal contracts, with meaningful tooling tests.
- [ ] Execute affected tests, evaluated profiles, required native builds and applicable runtime qualification; retain every failed result and distinguish cross-builds from runtime tests.
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
