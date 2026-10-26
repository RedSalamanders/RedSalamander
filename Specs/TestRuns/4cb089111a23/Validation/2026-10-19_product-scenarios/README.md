# Product documentation scenario gallery

Focused native x64 Debug presentation evidence, 2026-10-19. This extends the
ten File Operations illustrations with 45 Search, Space, plugin, Preferences,
connection, viewer, command and workspace/theme frames. The public entry is
[Visual Tour](../../../../../docs/VisualTour.md).

The app and ViewerPETests were built sequentially from the same frozen workspace
snapshot `852e660077215c655a5e36fcc8f3f5b4d9c8e447def440088c755203c511d71f`.
Both targeted builds completed with zero warnings/errors. Before viewer execution,
the capture driver verified both receipts against that exact snapshot/profile and
every output digest. The component receipts are preserved separately; neither is
a Full-solution receipt.

| Component | Receipt / result |
| --- | --- |
| App and plugin runtime | [App receipt](app-build-receipt.json), `36439b61e4eda860ece6330bd2dba33df0ea32f88784169912448965006be697` |
| ViewerPETests | [Viewer receipt](viewer-build-receipt.json), `fab9ce3622ade62ec2419a32e5065a993a905231cd9188b5bda7b080688ae20f` |
| Product gallery | [Commands result](commands-results.json): 1 passed, 0 failed/skipped; [26 scenarios](product-scenarios.tsv) |
| Viewer gallery | [Viewer results](viewer-results.json): 8 captures, 13,910 ms |
| Space lifecycle | [Viewer results](viewer-results.json): 11 captures, 5,044 ms |

The private Settings draft, in-memory Demo filesystem, fixed timestamps, fake MTP
picker and generated viewer content prevent real names, history, accounts or
devices from entering the frames. Search rows are seeded presentation state; no
search is admitted. The VLC frame shows its missing-runtime/download guidance,
not playback. The PE frame and native compiled-string assertion verify the UTF-8
resource correction preserves U+2022 separators.

Every native frame was reviewed in full. Twenty-three product frames were byte
identical to the already reviewed intermediate pass; the three changed product
frames and all nineteen viewer/Space frames were inspected again. OCR completed
for all 45 frames and found no identifying paths/names, sandbox paths or redaction
labels. Images are unedited native output, with adjacent source/build/DPI/hash
provenance. Space captures retain their original native state witnesses too.
Scrollable pages show their initial viewport; workspace/Space captures show the
complete native component surface, as identified in the public captions.

All three final native cases passed without owned foreground activation. Ambient
cursor/input changes were observed without restoring or taking over the desktop.
This evidence covers static presentation and fixture teardown, not Full, Release,
ARM64, real-input, real-device/provider, media playback or storage performance.
Case elapsed times are observational, not controlled performance comparisons.

The first startup rejection, rejected intermediate visual passes and viewer
teardown/module-lifetime failures are retained privately beneath
`D:\RedSalamander.Perf\evidence\product-scenarios-20261019`. The original hang has
a [continuation record](../../Continuation/2026-10-19_product-viewer-gallery-hang/README.md).
[Raw preservation metadata](raw-evidence-provenance.json) binds complete successful
run ZIPs with independently verified entry digests. Earlier published assets were
preserved before replacement; three obsolete safe configuration duplicates were
retired. No failed-attempt frame is published.

Documentation publication, metadata and tooling guards were updated after native
execution. The archived workspace is the capture origin; those later edits do not
turn the old build receipt into an attestation of the final documentation tree.
Focused tooling, link/schema coverage, inventories and sandbox cleanup results are
recorded alongside this evidence after their final checks.

The eleven Space frames and native witnesses are in the [companion Space archive](../2026-10-19_space-scenarios/README.md), keeping each archive within the standard 5 MiB budget.

Final focused validation: 396 unique Pester checks pass with zero failures/skips,
including resource localization, exact activation planning, source contracts and
published-capture identity/privacy. The published guards were repeated after the
archive split. Five native startup variants (prefix, case, multiple cases, flaky
injection and mixed suite) each return the expected exit 2 before creating UI.
Documentation checks cover 47 pages, 696 local links, 25 settings sections and all
16 plugin families with zero gaps. Spec inventory has zero errors/warnings.
OCR completes for all 77 currently published images; the sole broad path-pattern
match is the previously reviewed invented `Z:\Example\Reports\ExampleFolder`.
All changed native source bytes and every output digest still match their capture
origin. Later guide/tooling metadata edits are explicitly outside that attestation.
The preliminary tooling attempt's one missing-profile-context failure is preserved;
the configured final run passes. See adjacent compact validation summaries.

Cleanup removes nine exact task-owned run directories only after independently
reading back all 2,682 regular files from their private preservation ZIP. Long
fixture paths use the extended-length Windows path form for that preservation;
no links are followed. The original failed WinRT OCR setup and its empty scratch
run are included. The final marked-root disk audit is clean with zero issues and
no allowed-run exceptions. Raw captures, failed attempts, minidump and useful
evidence remain preserved outside the disposable run directories.
