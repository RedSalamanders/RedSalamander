# Native component evidence — draft Debug subsection

Status: **Debug diagnostic evidence prepared; Release and archive validation pending.**
Source: `e83dd755f44cef4f720d3f439f5fbd06383a569a`; snapshot and exact build receipt are bound in
`evidence-bindings.json`. No native metric's blank commit field was filled in. This packet joins the
outer exact filtered invocation, matching after/final workspace identities, receipt, executable path,
owned native process/exit record, metric timestamps and byte hashes. Sealed prelaunch files rotated
into `previous_run` explicitly bind the receipt ID, source snapshot, epoch zero and exact filtered
validation plan; archive inventory binds those bytes. Their original fingerprint/plan digests remain
intact rather than being reinterpreted as v2 promotions. The immutable receipt exactly matches its
copy in the sealed full-Debug-build context. Its executable/FileSystem.dll/Common.dll output hashes
are historical receipt identities, not fresh live-file checks: the authorized active Affected writer
is rebuilding/re-attesting this same Debug profile, so mutable current bytes are not historical proof.

Both selected FileOps runs have three native passes (Setup, requested case, Cleanup), zero failures,
zero skips, native process exit zero and suite PASSED. Outer wrappers exit 2 / repository
NOT_EVALUATED. Preserve their non-clean **non-blocking** sandbox audits: Review 100 issues; Identity
99 issues (19 root-child, 8 stale-run, 72 manual/unparseable sibling-run). Exact original lists/reasons
remain in the copied aggregate and `sandbox-audit.json`; ownership of unrelated residue is unproven.
These are not current-run root-escape/reparse findings, a clean-root claim or permission to clean.
Owning contract: `Testing_SelfTests.md` aggregate-audit/non-follow/preflight sections.

## Queue selector component

`Debug/Identity/perf/perf_metrics.jsonl` retains all 48 paired groups: 96 duration rows plus 48 capacity
rows. All three fixture sizes (4/64/256), storage/key directions and four counterbalanced pair indexes
are independently checked by `prepare_packet.py`; `Debug/queue-pairs.csv` retains every observation,
and `component-summary.json` includes all samples, medians, ranges and positive/negative deltas.
Durations measure **16-call batches**, not whole queue latency. Capacity's raw JSONL unit is `count`;
the protocol defines `value0` as baseline collection capacity bytes and `value1` as candidate bytes.
This resource observation is neither allocator calls nor heap/process peak. Correct selected IDs and
every observed scalar were asserted by the native case; packet checks additionally match expected
fixture ID, pair order, successful HRESULTs, actual raw rows and zero candidate collection capacity.

The actual production helper is compared with the frozen old algorithm in the same executable.
`source-excerpts/` binds the immutable baseline excerpt and current helper/adapter/harness. Both paths
retain matched per-call ID plus relaxed atomic capacity outputs. Baseline capacity-read/local adapter,
common output stores, nested task mutex, timing and compiler costs remain included. Tiny differences
cannot be assigned solely to selector logic. No p95 threshold or historical application speedup is claimed.
Debug is diagnostic; no Release latency, NAS, queue-start, allocation-saving or whole-queue acceptance.

## Guard and power witnesses

Review includes three real private admission/cleanup controls, each with one prompt and zero Delete
source offers: namespace-different IDs for the same physical entry (Cancel), distinct bound files
(Cancel), and distinct identityless overwrite (publication plus typed Completed/Retained). Exact bytes
and independent positive identity refusal were asserted natively. The namespace seam models a reported
ID-domain mismatch while retaining physical handles; it is not physical NAS/export evidence. The policy
conservatively withholds destructive distinctness for differing qualified Local roots, even real distinct
files or GUID-looking roots. Ordinary bound cleanup has separate source/publication authority.

All six actual conflict-wait controls are retained: Preanswered, Precancelled, QuietCancel, QuietResume,
StopOnly and Spurious. Counts describe bounded event/deadline calls under the test-only short quiet
budget; fake power APIs assert exact errors and failed-clear retirement. Private unadmitted tasks neither
exercise real Windows power/sleep nor measure historical polling CPU. Timings include declared responder
drain, and added correctness scenario durations are not equal-work historical comparisons.

## Preservation and remaining steps

Raw metrics exceed the 2-MiB file / 5-MiB run Git archive caps and remain in sealed local evidence.
Each compact JSONL preserves exact original row bytes/order; adjacent indexes bind full raw SHA256,
byte/line counts and every retained original line. Archive intent/completion records bind those raw
streams. Exact result/trace/receipt/workspace copies preserve original bytes and terminal newlines.
`Release/PENDING.json` is an honest empty slot, not a complete report.

Before publication: finish current native operation, obtain and review matched-source test-enabled
Release evidence, retain failures/skips and all samples, review this packet, add an exact archive
`-text -whitespace` attribute rule (currently text:auto would normalize digest-bound rows), then validate
the **complete** candidate packet with `Tools/Test-TestRunArchive.ps1 -RunPath <packet>` before any
force-add. Final closeout also requires the owning inventory validation. Neither validator has been
run by this preparation task. Do not claim Full/ARM64/ASan/physical-device/interactive qualification.
No source files, Specs archive, builds, tests or current active native run were changed by this task.
