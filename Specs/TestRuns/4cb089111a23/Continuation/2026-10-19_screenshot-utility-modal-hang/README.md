# Utility screenshot fixture hang

Status: original evidence preserved; corrected focused capture passes.

The x64 Debug build passes with zero warnings/errors. The exact nonactivating
product gallery stalls after six base workspace frames, before any utility frame.
The owned-window inventory contains no visible modal prompt. This is a hang, not
an observed crash. The case used a worker with synchronous HWND calls and an
unconditional join; that arrangement is unsuitable for an absent/early-return
modal and is being replaced with UI-thread timer capture and posted cancellation.
Local-only forms will use explicit presentation-only entry points.

Contained child 71816, parent 32576, was matched by executable path, full arguments
and creation time before collecting the minidump and stopping that exact instance.
The independent Release application was outside this operation. No failed frame
is published. Raw run data has independent ZIP readback verification.

The dump, matching executable/PDB and source are preserved privately under
`D:\RedSalamander.Perf\evidence\screenshot-gaps-20261019\first-attempt`.
Digests are recorded in preservation.json. No symbolic stack analysis or exact
original wait-site proof is claimed. The corrected 45-scene native App gallery passes with no failed/skipped cases.
The later missing-namespace compile failure is preserved privately too.
Successful evidence is recorded in the supplemental scenario validation archive;
this does not prove the original symbolic wait site.
The test-root policy overrides the forensic skill's historical legacy/AppData
search examples; this collection stays in the selected marked sandbox.
