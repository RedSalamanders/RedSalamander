# I66 restart qualification

## Hosted Release on the integrated main revision

Run `45907308619`, attempt 1, built and tested `09c663878b9902385db71600367fc9acec984008` with the shared engine pinned at `7c2f428fa8d49791093bc782db65c7e0a3c97b18`. Both native Release regression jobs failed. The test gate and production jobs were skipped. The green admission-only runs for earlier merges also skipped regression. [Hosted workflow](https://github.com/RedSalamanders/RedSalamander/actions/runs/45907308619)

| Lane | Canonical run | Result |
| --- | --- | --- |
| x64 | `20261126T234216Z-8804-2c476e53b9a347faa530c15334f32331` | 3,021 passed, 6 failed, 71 skipped. All six interactive entries were promoted. The only failed entry was Tooling Pester. |
| ARM64 | `20261126T235409Z-9592-5bf543a64bb84b2e8d6156ec6e75e8b8` | Tooling Pester failed the same six checks. Compare Directories was promoted. Commands recorded 962 passes, two failures and six skips in its original native result. The isolated ViewerSpace hover retry failed again, then hosted retirement refused the foreground console. File Operations and all four interactive viewer entries remained pending. The complete aggregate verdict is unavailable. |

The ARM64 retirement log records `Progman` at capture and `ConsoleWindowClass` at refusal. The foreground process was in a job, and membership in the exact test job and the parent process were unavailable to the identity query. An offline comparison of the recorded SHA-256 UTF-16LE image-path digest matched the standard System32 WSL executable path. The console’s ownership remained unresolved. The fixture refused to restore over it. [ARM64 regression job](https://github.com/RedSalamanders/RedSalamander/actions/runs/45907308619/job/140213035613)

The archived original Commands result has SHA-256 `9a263f04967b9866a43f78f92f35d617f1bb0d9020280ece5cfa717341c1ed02`. The file-action version case failed before its body while settling foreground ownership. The ViewerSpace hover case found the pointer over `ConsoleWindowClass` and never observed its own tooltip. Its isolated retry retained the same pointer failure and returned native exit 1 before the retirement refusal. Retry result SHA-256 `fa3dbb60b6c2aaa0b01081b597211403faae3cdb6f4022b62132bf7509b90d0b`.

The previous run on `94516001bcf582b646a6496002ab0cad48f32a69` had 3,017 passes, five failures and 71 skips on x64. ARM64 also failed at the foreign-foreground restoration boundary. [Previous hosted workflow](https://github.com/RedSalamanders/RedSalamander/actions/runs/45890003061)

Raw hosted artifacts and runner evidence were retained beneath the marked local test sandbox in archived import `20261127T071012Z-47220-ba04f9bd23c843feb938db4c61230e50`.

The read-only ARM64 DesktopProbe completed on the same source in a separate hosted run. It recorded session 2, the Default input desktop, and a `Windows.UI.Core.CoreWindow` foreground owned by `WWAHost` at that sample. Its repository verdict was `NOT_EVALUATED`. The console in the failed Commands run remains unidentified. Archived probe import `20261127T073108Z-55748-ce4e060020f242868c80cbbc77483c06`. [ARM64 desktop probe](https://github.com/RedSalamanders/RedSalamander/actions/runs/45958485341)

## Tooling repair

| Failure | Corrected witness |
| --- | --- |
| The shared PrePush behavior fixture omitted `test.ps1`, the compatibility entry point's target. | The shared fixture stages the real root `test.ps1` when present. It still checks hand-edited changelog refusal and ordinary pending-fragment acceptance. |
| The offline RedMonitor success fixture omitted the two pinned command files. | Stage both exact tool bytes with the client and lock. The existing corrupted and missing input controls still execute. |
| The RedMonitor tool-hash refusal used a wildcard in a Pester 3 error-message assertion. | Verify the reported command and the exact mismatch text. |
| The Static fixture expected a removed direct specification inventory invocation. | Exercise the actual Static branch and both root validation entry points. The recorded specification inventory call requires both failure-on-findings and report suppression. |
| The vcpkg process-ownership source guard expected a list conversion removed during argument construction. | Require delegation and the repository working directory at the streaming call with its selected argument array. |
| WinGet expected the shared renderer to omit even caller-supplied release dates. | The hosted publication entry point still renders without a release date argument, and the generator still avoids a date derived from rerun time. |

Focused Pester 3.4.0 run `20261127T072222Z-60988-39a5d39b3a94466583ce9b4295a8efa5` passed 312 checks with no failures or skips. It selected the RedMonitor, Static-policy, source-contract, and WinGet test files and excluded the unrelated MSBuild-tagged cohort. It included all 12 offline RedMonitor controls. The first attempt `20261127T072014Z-55868-af8d7e1f1b4a49a1ad50331cdd0f98e3` was archived with 311 passes and one failure before the corrected Static fixture created its stub directory.

The shared changelog behavior fixture passed all 40 cases against the RedSalamander product on the reviewed shared fix `4d5f00c056631047d2ff24e22a47ec2c4216ca1a`, run `20261127T072108Z-46748-fee3e9d72f264db8b126e0f5cceee9d2`. [Shared-engine PR #130](https://github.com/RedSalamanders/.github/pull/130)

The product integrated merge revision `317dfd0f03b21ef338cd250fd26847f215b5c3f9` after required Windows and Linux checks passed. On that pin, focused Pester 3.4.0 run `20261127T073600Z-19528-9d3bfc25a8c045b598d9ae411105f86e` passed 315 checks with zero failures or skips, including all six formerly failing checks. Whole-inventory archive validation passed for 18,765 files. The broader local engine suite had 128 passes and one existing PowerShell 7.6.9 schema-test mismatch. The unchanged schema returned false for the invalid manifest on that shell, and the old fixture expected an exception. Hosted Windows and Linux engine suites passed.

## Local x64 Release desktop pass

The shutdown profile refused reuse before taking focus. The follow-up used the byte-verified x64 Release solution receipt `5aaaae600267450f00e8a3e8c0fcabcd32331ba59c72a6de7d86dad247a47e29` on `96e495135b127a79f48c8ecd6e2157930991c898`, the head of PR #348. Fresh CI run `20261127T070925Z-27348-dd7cbf0dec974bf49844d7b72bb08804` began at 08:09 CET.

The selected Plug-in Contract and Viewer PE non-interactive entries reached promotion in this run. The original Tooling Pester invocation reported 1,491 passes and five failures, all in the tooling cases covered by the merged repair. CI then repeated the Pester entry for failure classification. At the 08:55 CET cut-off, that retry was still running and all six desktop entries remained pending. The contained launcher closed the owned process tree. It retained stdout and stderr in capture `20261127T070923Z-58036-5fac4f1c7d584a48966bdaa56fb4a435`. The runner evidence was archived as interrupted under the original run ID. The archived launcher stdout has SHA-256 `50d0fadf2ea4845f91233628330b291e5b8ac4f96aa64c20e4a743ee4d6eca3b`. The final native aggregate is unavailable.

## Integrated follow-up

RedSalamander merged the tooling repairs and shared engine pin in PR #357 at `3f2b01251c7e4bc8f8483ebbd8460444174cf1d1`. The test-enabled x64 Release rebuild took 7 minutes 36 seconds and had zero warnings and errors. All 188 output records matched receipt `946bd916ef3e93a02ca5abcf9868baa1f45f8eeac6462611d1e8982047a1b47c`. Fresh CI run `20261127T080428Z-67632-151d4572a667447aa54cc951dd320513` started at 09:04 CET under the approved follow-up window ending at 10:54 CET. The runner retains stdout and stderr in capture `20261127T080427Z-71012-6934f4b8af8642538e47be24ed795227`. The hosted Release jobs are also active on the same source. [PR #357](https://github.com/RedSalamanders/RedSalamander/pull/357) [Hosted qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/45960884674)

The integrated run promoted Tooling Pester with 1,500 passes and zero failures or skips. Terminal result digest `2974f819fb8b21cf75cdea37a08a2e9aba3f95c0573cbd5b3a2be21573460ae0`.

Commands promoted with 965 passes, no failures and five skips. Terminal result digest `4e20b346253d5e8394ec87a10c15ae32ac5e718a56c10cb2d4d5b7ff5cea4bab`.
The Commands result includes passing Pack and Unpack dialog checks, the ZIP worker round trip, the folder drop guard, and both cases affected by the previous ARM64 foreground failure.

### Local File Operations native failure

The first File Operations process (PID 71312) started at 09:55 CET and exited at 10:29 CET with `0xC0000374` (`STATUS_HEAP_CORRUPTION`). Windows Error Reporting recorded `ntdll.dll` and report ID `4ff6ce0a-6d82-4da0-b508-13a213aa9dbc`. The final trace checkpoint was `Phase7_DeleteItemsNestedRootsWake`. The last JSON is partial and provides no terminal suite verdict. CI began its automatic classification retry under the same approved lease. Four local viewer entries were still pending.

The matching private symbols resolve the exception thread to `TickPhase7` in `FolderWindow.FileOperations.Tests.Phases07_09.cpp:4425`. The heap manager reported the corruption during a `std::filesystem::path` allocation while the test prepared a nested-delete round. This is the detection site. The dump does not locate the original write.

| Preserved artifact | Run | SHA-256 |
| --- | --- | --- |
| Original File Operations trace | `20261127T093150Z-76092-b58767c5debd44aaaf1c8cc683d5361e` | `30deb671f949053f5402b39e62de7671257a8514f41b5baacb6fd301f17512f8` |
| Original WER dump | `20261127T093315Z-73340-823abf554bf24f089fbf1b781bf3b675` | `9ce818432cdcbb819bb108987fa90c3bcd0c54bf8e7767b56f44008a7473f1dc` |
| Source-resolved stack and thread stacks | `20261127T093818Z-70840-dc3487fda9a447e6954d6030e4557542` | `b90ba71166a499a611e86d3f1091798f78c60be1d198c9b7a9ece361b5418cd0` |

All 17 Local scheduler workers were present in the dump. The provider's parallel Delete used 16 contexts and mapped worker ID 16 to slot zero. Either worker could overwrite the other's path arenas and progress state during a Delete. This is a confirmed race on the failing route. The dump does not capture that write. The repair uses the scheduler's full slot count with direct indexing and extends the existing witness to check one successful completion per selected root. [Repair PR #360](https://github.com/RedSalamanders/RedSalamander/pull/360)

The matching executable, provider DLL and private symbols were copied and digest-verified in `20261127T093921Z-65580-357b912006044ade9a4442cef2d4eaf2`. Raw artifacts remain in the owned local evidence sandbox.
