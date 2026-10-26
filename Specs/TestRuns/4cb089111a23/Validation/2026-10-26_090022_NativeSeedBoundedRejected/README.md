# Rejected bounded native seed pool

Source `487a99bc2b37f3f67e16924fffb0ba096447dab5` built x64 ASan Debug
and Release without warnings/errors. All 17 focused native checks passed,
including the original service startup watchdog. The paired comparison failed
Release memory admission, so no Full was started and no qualification is claimed.

| Profile | Seed wall-time median ratio | Sampled private-memory median ratio | Verdict |
| --- | --- | --- | --- |
| ASan Debug | 0.642555 | 0.728908 | PASS |
| Release | 1.044102 | 1.325656 | FAILED |

Release sampled peak private-memory median rose from 844,038,144 to
1,118,904,320 bytes. The baseline remains
`fac05d5b70b9a5933b12093c1b2337f76aee8cdc`; neither the admission limits,
native watchdogs nor sanitizer settings changed. These are fresh-process runs
on the same machine/volume/scenario, with evolving MFT record counts recorded
in the original observations rather than a frozen volume image.

`manifest.json` binds all 38 raw payloads: the failed state/comparison, six
native results with complete metrics/memory samples, parser/service results,
two build receipts and exact external drivers. Failure remains explicit.
The next scratch-storage/borrowed-parser candidate requires the same retained
baseline comparison. Close repair, six final builds, all three Fresh Full
profiles, ETW/ASan detection and public migration remain pending.
