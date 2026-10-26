# C0 Curl lost mutation reply — test-enabled Release

- Source: `58f22d0e` plus the same C0 working changes as the
  [Debug candidate](../2026-09-05_c0_curl_reply_loss_debug/README.md).
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter C0_ -FailFast -RunId c0-curl-reply-loss-release-20260905`, with process
  `RSBuildEnableTests=true` explicitly set and restored afterward.
- Accepted build receipt: `84fb98a9fea4b800b242f45483c9a9533be20f9a85466419c9ff4184bc4bf871`;
  verified `tests_enabled:true`, x64 Release. Zero warnings/errors; build 3:53.
- Harness: 9/9 passed, zero skipped; duration 6,935 ms. Native Curl delete guards
  29/29; native lost-reply checks 33/33; real host Delete feedback witness passed.

All eight native mutation scenarios issued exactly one mutation. Lost-reply cases
preserved the replacement and sibling bytes, retained the first Rename result,
reported an uncertain completion, and did not issue incidental NLST. Ordinary
controls stayed successful. Fault-case durations were 2,713 / 2,434 / 3,631 /
2,922 us for DeleteItem / DeleteItems / RenameItem / RenameItems. These Release
timings are separate from the pre-fix Debug baseline, not an optimization ratio.

The host Delete witness took 63 ms: one request, one backend commit, preserved
replacement/sibling bytes, immutable Indeterminate result, `ERROR_IO_INCOMPLETE`,
localized `Outcome unknown` summary, and zero conflict prompts. The two existing
stalled-list cancellation checks returned cooperatively in 1,000 ms each, with
zero mutations. Google Drive, Local publication and ABI guards stayed green.

Evidence files are intact copies from
`C:\RedSalamander.Perf\runs\c0-curl-reply-loss-release-20260905\artifacts\selftest\last_run`.
The runner reported 36 retained-run disk-audit issues; no old data was removed.
This is focused test-enabled Release evidence, not the separately required
non-test Release staging transition, real SSH-server/native-walker completeness,
or final Fresh Full qualification. C0–C6 are not declared complete by this run.
