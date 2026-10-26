# I27 sandbox containment diagnostics

These records qualify the bounded tooling change and its focused tests. They do not
qualify native File Operations or a Full repository run. `retained-records.json` binds
the eight byte-identical originals, their sizes/hashes and external retention paths.

The existing private ancestor walk now uses fresh `File.GetAttributes` and
`Path.GetDirectoryName` calls. It caches no prior authority. Exact marked fixed-drive
authorization, lexical containment and alternate-data-stream rejection stay in the
owning helper. Only file-not-found and directory-not-found permit missing components;
other failures propagate. Existing and dangling reparse components are rejected.

The unchanged diagnostic fixture has 192 ordinary files, 32 directories and two
reparse leaves. Baseline/candidate/candidate/baseline order retains identical ordinary
identity/content inventories and raw reparse proofs in all four trials:

| Trial | Capture wall time | Independent reader wall time |
| --- | ---: | ---: |
| Baseline 1 | 50.342 s | 9.611 s |
| Candidate 1 | 7.787 s | 1.666 s |
| Candidate 2 | 7.799 s | 1.760 s |
| Baseline 2 | 57.506 s | 10.184 s |

Filesystem/JIT caches and simultaneous original-archive activity were uncontrolled.
This small diagnostic supports removal of provider overhead; it is not a large-archive,
native latency, GPU, memory-peak or leak measurement. Handle counts are endpoint samples.
Fresh checks detect replacement between calls; they do not make a walk atomic against
concurrent mutation. Named injected I/O/access failures are not real OS DACL evidence.

The parity overlay executed 48 controls with no failures, including real junctions,
file/directory/dangling symbolic links, deep French paths, inter-call ancestor replacement,
D-to-C capture, empty trees and corrupted archive rejection. The corrected diagnostic
Pester overlay passed ten cases. Its first attempt failed eight because a Pester 3 module
mock leaked; that original failed report remains here. Root replaced it with explicitly
scoped injection restored in `finally`, without changing shared markers or source files.

After canonical adoption, focused testing passed 353 cases, zero failures/skips:
containment 10, tool inventory 17, tooling governance 12, harness contracts 191,
runner planning 59, documentation drift 18, spec architecture 37 and environment 9.
Tool/spec inventories reported no blocking findings. These focused results do not
replace the remaining fresh Full Debug/Release qualification.

The failed Full `20261004T054036Z-80612-60e6bb68d8934d4aa370d59535713f24` remains FAILED.
Its original archiver was stopped after native containment ended; the original archive
is incomplete. A separately sealed later capture passed independent readers for
52,751 C files/seven reparse leaves and 80 C-validation files. D was absent only at the
later observation. Failed-only retention preserved two immutable receipts and passed
all 39 canonical epoch/output-map bindings. Original completeness and qualification
remain false. The retained capture/wrapper/report state those limits explicitly.

Large archives, raw link blobs, original console logs, scripts, failed drafts and
source snapshots remain under the external retention root in the map. No original
link-target payload, native pass, global sandbox cleanliness or platform qualification
is reconstructed from this packet.
