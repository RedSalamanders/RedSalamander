# Test-run evidence archival

Status: DONE. Owner: RedSalamander test tooling. Scope: runtime-run retention and tooling lifecycle.

The runner must retain completed and interrupted runs under the selected marked sandbox's
evidence directory, without duplicate payload copies or automatic destruction of unique
evidence. Current contracts remain in `Specs/Testing/Testing_SelfTests.md`.

- [x] Trace runner publication, cleanup, reporting and evidence identities.
- [x] Add same-volume, create-only, verified archival and abandoned-run preservation.
- [x] Cover successful/failed/interrupted runs, owner protection, collisions, links, streams,
      disk receipts and interrupted publication recovery with filesystem tests.
- [x] Update help, inventory, normative contracts and CI artifact collection.
- [x] Run scoped tooling validation, fresh ReFS/NTFS controls, and junction read probes.
- [x] Record results and close the plan after relevant gates pass.

The runner retains complete runs under `evidence/completed-runs`, including failed,
skipped and early-exit witnesses. Create-only intent/completion records preserve path
mapping and verified content/stream/security metadata. Abandoned runs are preserved;
live owners and unsafe or unresolved resources remain protected. Same-volume renames
preserve shared file identity and read-only state without duplicate payload copies.
Containment tests record and retire their own synthetic negative-fixture links.

Local evidence: `Z:\RedSalamander.Perf\evidence\test-run-archive-workflow-20261021`.
The broad focused tooling pass contains 509 passing tests; the later archive/containment
pass includes the added initialization regression. Final native controls, audit results,
source digests and junction probes are retained in that directory's report and logs.
The first fresh x64 Debug app build and focused ReFS execution passed; final
receipt-bound repetitions verify the closed source snapshot. Broader product/platform
and interactive qualification remains governed separately.

Five failed validation/debug fixture trees were retained with independent full SHA-256,
stream, identity and metadata indices. Seven synthetic test links were retired only
after their raw reparse buffers and identities were preserved. No payload file was
deleted in that preservation step. Historical failures remain failed evidence.
