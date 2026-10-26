# File Operations device recovery qualification

The exact source snapshot, build receipt, pass/fail/skip results and all timing samples are
retained alongside this file. Both Debug builds report zero warnings/errors and clean sandbox
audits. The first run passes 3/3; the expanded second run passes 6/6, with three repetitions.

The harness forces three host device recreations per repetition. The consumer's previous
context remains retained until normal rendering replaces it, making the distinct-context
assertion meaningful. Every recovery checks frame validity, hosted graph/disclosure state,
French foreground/background pixels, and collapse/expand actions. No application operation
or independently launched process is affected. Screenshots use the shared WGC harness.

The expanded run pumps real messages during three hidden intervals per repetition, requiring
unchanged hosted render counts/context identity, then resumed painting when shown without
activation. Its nine intervals last 359–375 ms. This does not establish idle CPU/wake-up,
private/GPU memory retention, actual adapter removal, native DPI or Release/Full acceptance.

Metrics-compaction binds the original raw file's bytes, row count, SHA-256 and selected line
numbers. Selected rows are unchanged and ordered; their digest uses UTF-8 without BOM and LF
terminators. All DeviceRecovery rows are retained in perf_metrics.jsonl. Three recovery redraws
form each RedrawUs sample; these are not presented FPS. HiddenFrames values are before/after
host render counts; duration is the observed hidden interval. Both fixtures exercise the same
production implementation, so no before/after optimization claim is made. Full raw data stays
under C:/RedSalamander.Perf/evidence/i26-ui/device-recovery* and completed S-drive evidence.
