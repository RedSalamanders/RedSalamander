# Current PrePush clipboard source-guard triage (read-only; corrected literal check)

Run: `20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f`.

## Recorded failure

`D:\RedSalamander.Perf\runs\20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f\artifacts\selftest\last_run\commands\results.json` records `file_operations_phase0_typed_contract_source_guard` as failed, duration 550 ms, with reason `Clipboard Move must be consumed at accepted admission, never deferred to terminal success.` The corresponding `commands\trace.txt` lines 500–502 record the same case and reason; the following `file_operations_preparing_lifecycle_source_guard` passes. The reason is a generic assertion message, not a term-by-term report.

## Exact guard and source read

Candidate `RedSalamander/SelfTest/Commands/Commands.Tests.PluginConfig.cpp:5254-5256` defines the case and calls `SelfTest::TryFindRepoRoot()`. Lines 6033–6040 read `repoRoot\RedSalamander\FolderView.FileOps.cpp` and require all of:

1. text contains `ConsumeMoveClipboardForAdmission`;
2. text contains the exact literal `request.moveClipboardSequence = clipboardSequenceNumber;`;
3. text does not contain `request.completionCallback = [clipboardSequenceNumber]`.

`ReadSourceFileUtf8` at lines 4798–4808 copies the `ifstream` bytes into `std::string` with stream iterators; it does not normalize whitespace. `TryFindRepoRootImpl` in candidate `RedSalamander/SelfTest/Common/Common.Tests.Runner.cpp:760-798` first accepts `REDSALAMANDER_REPO_ROOT` when valid, otherwise walks upward from process current directory, then from executable module directory. The run artifacts reviewed here do not record the environment value, process current directory, or the resolved absolute `repoRoot`.

## Exact literal predicate results and diff

I read candidate and primary `FolderView.FileOps.cpp` contents and evaluated each ASCII needle with ordinal, whitespace-sensitive string containment (PowerShell/.NET `String.Contains(..., StringComparison.Ordinal)`; same literal semantics as the guard's `std::string::find`). Both copies yield:

- `ConsumeMoveClipboardForAdmission`: present (`true`);
- exact `request.moveClipboardSequence = clipboardSequenceNumber;`: absent (`false`);
- `request.completionCallback = [clipboardSequenceNumber]`: absent (`true` for the guard's required absence).

Both files have the same SHA-256: `AF95943866040FFE7AECB9CBAB9D07815036F1841748D58A43DB64EB7EBDDDF0`. At line 868 the actual spelling is `request.moveClipboardSequence   = clipboardSequenceNumber;` with three spaces before `=`. This fails the one-space literal. Candidate HEAD is `993c3ff14ce60eb7331a0f36bd54475e227859e9`; primary HEAD is `1277ffdec9cf78b5452d29f695aae7806fffe207`. The candidate implementation file and guard test file are unchanged from candidate HEAD, so the scoped working-tree diff for them is empty.

Run trace plugin paths are under the candidate worktree's `.build\x64\Debug\Plugins`, but do not prove which repoRoot the source guard resolved. This ambiguity is not causal to the observed predicate mismatch: candidate and primary source copies both have identical bytes and both fail the exact second predicate.

## Conclusion and narrow follow-up

The recorded failure is explained by a whitespace-sensitive false negative: the source spells the assignment with three spaces, while the source guard requires one. The guard's generic message is consistent with the second term failing; the two other terms pass against both inspected source copies. A narrowly scoped correction can align source spelling or guard expectation, then validate the existing source guard under the governed workflow without changing clipboard admission semantics. No test, build, source edit, or rerun was performed for this triage.
