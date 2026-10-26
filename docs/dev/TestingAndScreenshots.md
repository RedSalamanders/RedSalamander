# Tests, Evidence and Documentation Captures

Start with [Tests/README](../../Tests/README.md) and the authoritative
[self-test contract](../../Specs/Testing/Testing_SelfTests.md).
Use [Test-Changes.ps1](../../Test-Changes.ps1) to explain the affected scope
before iteration, and PrePush mode to account for local/CI obligations.
An affected, focused or pending-CI result is not a complete repository pass.

## Keep generated data in the owned sandbox

Tests use an initialized, marked `X:\RedSalamander.Perf` on a fixed local drive.
The selected root is shared by native and tooling tests. Scratch lives under
`runs/<run-id>`, durable validation evidence under `evidence`, and machine
resource declarations under `config`. An environment variable selects a root;
it does not authorize an arbitrary directory or replace the ownership marker.

Repository `.build` holds build inputs/outputs. It is not a test-data sandbox.
Run tests through the canonical runner so process temporary paths, artifact
ownership, build receipts and desktop leases follow the same contract.

## Maintain a test machine safely

Review each reported audit path and its owner/process use. Preserve useful
results, failures/skips and diagnostics before retiring obsolete scratch.
Verify archive bytes and unchanged source identities. Keep recovery bundles
and validation evidence separate from run-directory deletion.

Preview one exact reviewed run:

```powershell
.\Tools\Clean-TestSandbox.ps1 -TestRoot D:\RedSalamander.Perf -RunId '<reviewed-run-id>'
```

Add `-Apply` only after review. The command protects parseable live owners using
the runner's PID-reuse policy. A failed removal returns `Status: Failed`, warns
and exits nonzero. WhatIf, preview, absence and success exit zero. Audits require
current root/marker authority even when a requested directory does not exist.

Manual names never establish automatic cleanup ownership. Reparse points are
not followed; a mount or a test-owned virtual disk requires its recorded safe
teardown. Empty misplaced root remnants need exact containment and nonrecursive
removal. Restore selected archived regular files into fresh contained scratch,
never over the active root or by automatically recreating recorded links.

## Produce a documentation screenshot

Use a deterministic test-harness scenario and the shared owned-window screenshot
helper. Seed realistic invented names, paths, accounts, hosts and document
content before capture. Check titles, breadcrumbs, rows, details and tooltips.
Actual test I/O still stays in the marked sandbox; display-only paths grant no
filesystem access. Screenshots are saved unedited: redaction labels, masking,
blur and later pixel edits do not meet the contract.

The File Operations documentation fixture can run without foreground input:

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_ReviewRemediationContracts -NonInteractive -Configuration Debug -Platform x64 -TestRoot D:\RedSalamander.Perf
```

Its ten `fileops/documentation` artifacts show running transfer facts, open
counting, totals-only Move, file-only completion, paused/queued/completed Copy,
a destination conflict, the complete wrapped terminal note, and a selected issue
in the real Issues pane. The fixture uses private synthetic state and opens none of its displayed
example paths. Use the broader visual gallery for additional presentation
states, with its documented selection/desktop rules.

Review the complete native frame before copying an accepted artifact into
`docs/res`. Preserve scenario, source/build receipt, executable identity,
language, theme, size, DPI and image hash in adjacent provenance. Record any
synthetic state and qualification limits. An image is an illustration; it does
not establish storage-performance, foreground-input, Full or other-platform
qualification.

Static capture defaults to no activation. Interaction scenarios require an
agreed time, the existing warning and runner lease, owned-window targeting, and
cursor/foreground restoration. Computer Use and manual desktop grabs are not
substitutes. Keep proposal mockups clearly separate from application captures.

The [screenshot inventory](../res/README.md) records the published set and known
capture gaps. Remove an unsafe illustration from publication until a compliant
harness replacement is available; preserve its review evidence outside the
published image directory.

## Capture the wider product

The expanded `cmd_product_documentation_gallery` Commands case includes the
original 26 workspace, Search, Preferences, plugin, connection and palette scenes
without activation or global input. Mixed/broader Commands filters keep their normal
desktop-input contract. Search rows are seeded presentation state; no search runs.
Connection profiles are private drafts and the MTP picker uses the fake backend.

`ViewerPETests TestViewerDocumentationGallery --no-activate` captures eight actual
viewers from generated data; `TestViewerSpaceLifecycleCaptures --no-activate`
captures thirteen Space lifecycle/workspace states. Build ViewerPETests and the app/plugin
runtime from the same source snapshot and profile, verify both receipt output
closures, and preserve both receipts. A targeted receipt does not qualify Full.

Keep shared plugin modules loaded through all viewer roles, close/release consumers,
drain callback references before providers retire, and use the plugin shutdown
quiet point before unloading. Preserve failed captures separately; publish only
the reviewed successful run. The [Visual Tour](../VisualTour.md) is the public entry.

The expanded product case has 45 frames and includes modal utilities, Terminal and
provider browsing. Cancel every modal; use injected shared-directory rows rather
than NetShareEnum. Terminal transcript mode consumes one thread-local factory arm
and starts no process. Provider work uses the existing loopback APIs or RAM MTP,
with workspace destruction before endpoint teardown.
For successful media, run `TestViewerVlcDocumentationPlayback --no-activate` and
retain its generated-media/runtime hashes and real decoder pause observations.
