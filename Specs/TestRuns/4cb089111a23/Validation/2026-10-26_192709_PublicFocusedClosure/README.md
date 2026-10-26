# Focused public migration qualification

Candidate b5ea0d679dd1fe7d4d418ee7e8d8cb49a7cb7d26 passed both x64 and ARM64
ASan Debug solution builds, all eight selected standalone native checks and
18 original-regression checks (12 Commands observations across two repetitions,
three search checks and selection setup/case/cleanup). There were no selected
skips. The negative ASan control intentionally exited nonzero with a real
heap-buffer-overflow report. A separate fresh-process theme run passed both
repetitions. The 291 focused harness/routing/specification controls passed.
These receipts qualify their recorded source, dependency, output and scenario
identities. Full runtime reruns were omitted under the user's accepted scope.
They do not promote earlier failed Full or hosted CI receipts.

The retained 03cb2510b failure is the same-machine ASan baseline for the terminal
queue repair: the unchanged 200k-file/20k-directory case missed its 180-second
deadline while new analysis builders repeatedly held the final update tail.
Candidate b5ea0d679 prioritizes the completed producer's pending bytes before a
new projection. The protocol process completed in 64.05 seconds; its large-tree
case reached exact 200,000-file/20,001-directory totals, zero pending bytes,
8,388,544 peak pending bytes and zero validation errors. Posted 220,013 minus
58,000 coalesced equaled 162,013 drained. Existing deadlines, transport/resource
bounds and data predicates remain unchanged. Additional unbuffered diagnostic
output makes this an observed before/after qualification, not an isolated
production throughput benchmark or a general performance guarantee.

The physical-pixel small-work-area control in 03cb failed at 144 DPI because it
could not represent the mandatory analysis panel. The candidate expresses the
control as 1024 by 680 DIPs, bounded by the actual work area; all original
analysis/state/layer/foreground witnesses pass. The included native harness PNG
shows the actual synthetic fixture, with matching unmodified provenance.

The earlier cold theme handle-guard failure stays failed. The candidate adds
exact counts and preserves the +2 GDI/USER, same-HWND and 8 MiB bounds. In the
fresh original cluster it observed GDI 64 to 65, USER 95 to 97; its second burst
observed no handle growth. The measured private-byte changes were negative in
both repetitions. No causal claim is made for the earlier intermittent handle
failure; no sanitizer option or quarantine setting was weakened.

Raw results/logs are copied byte-for-byte. Theme metrics are explicitly selected
retention records; the manifest identifies and hashes their complete external
raw originals, which remain retained. Build and observer records bind the full
receipts and artifact hashes. No raw original was replaced or deleted.
Public PR/main CI and package/manual/installed-helper gates remain separate.
Hosted x64 b5ea0d679 Debug compiled successfully. Suite PR retained 1,018 passing,
one failed aggregate viewer entry and two existing optional skips. Every failed
capture rejected its owned visible HWND at IGraphicsCaptureItemInterop::CreateForWindow
with E_INVALIDARG, before frame/pixel validation. The complete failed job log is
retained here. The next candidate repairs the missing Windows Runtime initialization
precondition, preserving STA/MTA and using WIL's balanced runtime cleanup. Whether
that clears the hosted error remains unproven until current-source CI runs.
