# File Operations decisions and progress review

The requested fixes are implemented. The final test-enabled x64 Debug build passed with zero
warnings/errors; all three new focused cases passed on receipt
`047f98907061f915d00f1f490b967e89fab62bf15c3f8aa6169c531d0240b5af`.
The sibling README and preserved aggregates give the exact assertions and qualification limits.
All seven final exact runs passed: three new scenarios plus the existing error-message, checked-Skip,
single-folder and multiple-folder regressions. That is21 passed cases including each run's setup and
cleanup, with zero failures or skips. [validation-summary.json](validation-summary.json) binds every
run ID, archive and the final receipt.

| Scenario | Cause and resulting behavior |
|---|---|
| Checked Skip repeats | The broker excluded ordinary Skip from checked decisions. Checked Skip now records the same eligible scoped grant as Skip All. |
| Keep Both repeats | Some offered contexts could not reuse their saved choice. Selected-root and descendant contexts now have separate Keep Both grants, with reuse among compatible descendants. |
| Metadata consent repeats | The key used an individual destination path. It now uses the admitted destination namespace root, preserving the risk/task/profile boundaries. |
| Error `0x80070002` | This is Windows file-not-found. The UI now explains that a required file could not be found and preserves its code. The screenshot alone does not establish why it disappeared or which prerequisite was missing. Known path errors and other Windows errors also get meaningful descriptions. |
| Skip counting | The UI kept discovery and completed-work rows together, with a selected-root denominator. Copy now uses one `Copied x files; y folders; z written` summary; unknown counts use a truthful fallback. |
| Permanent Delete progress | Bound deletion did not provide completed-child facts, and the fixed card height overlapped extra rows. Confirmed object removals now supply file/folder/byte counts; layout follows the visible wrapped rows. |
| Disk full | The generic conflict menu offered irrelevant choices. Qualified destinations now show Retry / Skip All / Cancel directly; Skip All covers later matching disk-full errors. Unqualified destinations show Retry / Cancel and explain the missing scope. |
| Read-only deletion | There was no explicit read-only deletion action, and selected roots could bypass that specific consent. A retained exact item/tree now requires explicit authorization before read-only deletion; the grant does not carry to another selected root. |
| Other wording | Unencrypted Move now explains removal of the original and distinguishes copying while keeping it. Invalid-name and link-target messages avoid guessing which endpoint failed. |

## Number of questions

There are **26 conflict categories** and **36 base question/problem resources** in the shared
question path, plus six composition resources. Dynamic paths, consent details and Windows/provider
descriptions vary, so these are template counts, not a count of every possible displayed sentence.
`question-template-inventory.json` lists the exact IDs and counting boundaries.

**64 is the remembered decision-scope capacity**, not the number of question types or files.
Item kinds, profiles, namespace roots and Keep Both context can produce more than64 distinct scopes.
Cache exhaustion now withdraws the repeat option with an explanation. A separate64-slot consent
receipt table reuses one checked same-risk/root receipt; unchecked per-item receipts can still
exhaust that independently bounded table and fail closed before the affected mutation.

## Evidence boundaries

Review covered all26 File Operations categories in source and policy checks. Runtime evidence covers
the named native/bridge scenarios, with36 English/French layout checks and18 harness captures.
These are focused passes, not a Full repository, Release, ARM64, remote-provider or foreground-suite
verdict. The preserved aggregates retain three unrelated legacy sandbox-directory findings.
The new measurements are bounded witnesses; no before/after throughput improvement is claimed.

The earlier CRT assertion was caused by the direct-provider test fixture omitting the executor's
active callback scope. The corrected fixture preserves that assertion and passes on the final build.
An intermediate bridge fixture also required a parallel-only metric from the sequential route;
its corrected guard retains exact file/folder counts and payload checks for both routes.
