# Local Debug Full and long-path fixture correction

Frozen RedSalamander `b2be60e421df1a25d84bcfa83ba78c9e8e32513f`, DxUi `13788e95`,
completes Fresh Full locally on 2026-09-13. The root Debug build passes in 3m57s with
zero warnings/errors. All 14 module sidecars match the actual binaries after Full's
build-tooling re-attestation.

The run `20260913T005200Z-93948-3aca8cd69cf74c458a0b151bc62940dc` is **failed**:
19/20 entries are promoted; 2,116 cases pass, one fails and 52 skip. Commands passes
898/0/2, Compare 242/0/30 and File Operations 180/0/20. Both interactive viewer entries
and the performance entries pass. ViewerWeb's earlier close-time assertion passes
at 29,192 microseconds with zero provider timeouts; persistent direct menu hover
passes in 157 ms. These results do not explain or erase earlier failed attempts.

The sole failure is in the tooling Pester entry (769 passed, one failed): the new
compiler-host fixture runs `git clone` without long-path support. Full's deeper
sandbox makes `.git/hooks/fsmonitor-watchman.sample` exceed MAX_PATH, so Git rejects
the clone before restore assertions. The correction adopts the same process/fixture-local
`core.longpaths=true` clone settings as the existing dependency fixtures. The fixture
also pads short focused-run roots and asserts that the hook path exceeds 260 characters.
No global Git setting, product code, compiler policy or test assertion is relaxed.

All eight dependency tests pass after correction, including real six-profile MSBuild
host evaluation and environment restoration. Their summary records the candidate test
file hash and the base source revision; the change was uncommitted during that focused run.
The failed Full remains failed. A new final three-profile sequence is required.

`retained-files.json` hashes eight original receipts. The complete local archive,
entry receipts and validation evidence remain under
`Z:/RedSalamander.Perf/evidence/I19-20260909/b2be60e4-final-x64-Debug-Full`.
No hosted CI ran. Native ARM64 runtime, resource advice and final package acceptance
remain separate gates in the opening I19 checklist.
