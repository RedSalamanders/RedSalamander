# I66 restart qualification

## Hosted Release on the integrated main revision

Run `45907308619`, attempt 1, built and tested `09c663878b9902385db71600367fc9acec984008` with the shared engine pinned at `7c2f428fa8d49791093bc782db65c7e0a3c97b18`. Both native Release regression jobs failed. The test gate and production jobs were skipped. The green admission-only runs for earlier merges also skipped regression. [Hosted workflow](https://github.com/RedSalamanders/RedSalamander/actions/runs/45907308619)

| Lane | Canonical run | Result |
| --- | --- | --- |
| x64 | `20261126T234216Z-8804-2c476e53b9a347faa530c15334f32331` | 3,021 passed, 6 failed, 71 skipped. All six interactive entries were promoted. The only failed entry was Tooling Pester. |
| ARM64 | `20261126T235409Z-9592-5bf543a64bb84b2e8d6156ec6e75e8b8` | Tooling Pester failed the same six checks. Compare Directories was promoted. Commands stopped during hosted desktop retirement and its terminal promotion is absent. File Operations and the four later interactive viewer entries remained pending. No complete aggregate verdict was produced by this runner. |

The ARM64 retirement log records `Progman` at capture and `ConsoleWindowClass` at refusal. The foreground process was in a job but membership in the exact test job was unavailable. Its parent process was unavailable to the identity query. The fixture refused to restore over that console. This observation leaves ownership unresolved. [ARM64 regression job](https://github.com/RedSalamanders/RedSalamander/actions/runs/45907308619/job/140213035613)

The previous run on `94516001bcf582b646a6496002ab0cad48f32a69` had 3,017 passes, five failures and 71 skips on x64. ARM64 also failed at the foreign-foreground restoration boundary. [Previous hosted workflow](https://github.com/RedSalamanders/RedSalamander/actions/runs/45890003061)

Raw hosted artifacts and runner evidence were retained beneath the marked local test sandbox in archived import `20261127T071012Z-47220-ba04f9bd23c843feb938db4c61230e50`.

## Tooling repair

| Failure | Corrected witness |
| --- | --- |
| The shared PrePush behavior fixture omitted `test.ps1`, the compatibility entry point's target. | The shared fixture stages the real root `test.ps1` when present. It still checks hand-edited changelog refusal and ordinary pending-fragment acceptance. |
| The offline RedMonitor success fixture omitted the two pinned command files. | Stage both exact tool bytes with the client and lock. The existing corrupted and missing input controls still execute. |
| The RedMonitor tool-hash refusal used a wildcard in a Pester 3 error-message assertion. | Verify the reported command and the exact mismatch text. |
| The Static fixture expected a removed direct specification inventory invocation. | Exercise the actual Static branch and the live `validate-specs.ps1` wrapper. The recorded inventory call requires both failure-on-findings and report suppression. |
| The vcpkg process-ownership source guard expected a list conversion removed during argument construction. | Require delegation and the repository working directory at the streaming call with its selected argument array. |
| WinGet expected the shared renderer to omit even caller-supplied release dates. | The hosted publication entry point still renders without a release date argument, and the generator still avoids a date derived from rerun time. |

Focused Pester 3.4.0 run `20261127T072222Z-60988-39a5d39b3a94466583ce9b4295a8efa5` passed 312 checks with no failures or skips. It selected the RedMonitor, Static-policy, source-contract, and WinGet test files and excluded the unrelated MSBuild-tagged cohort. It included all 12 offline RedMonitor controls. The first attempt `20261127T072014Z-55868-af8d7e1f1b4a49a1ad50331cdd0f98e3` was archived with 311 passes and one failure before the corrected Static fixture created its stub directory.

The shared changelog behavior fixture passed all 40 cases against the RedSalamander product on the reviewed shared fix `4d5f00c056631047d2ff24e22a47ec2c4216ca1a`, run `20261127T072108Z-46748-fee3e9d72f264db8b126e0f5cceee9d2`. [Shared-engine PR #130](https://github.com/RedSalamanders/.github/pull/130)

## Local x64 Release desktop pass

The shutdown profile refused reuse before taking focus. The follow-up used the byte-verified x64 Release solution receipt `5aaaae600267450f00e8a3e8c0fcabcd32331ba59c72a6de7d86dad247a47e29` on `96e495135b127a79f48c8ecd6e2157930991c898`, the head of PR #348. Fresh CI run `20261127T070925Z-27348-dd7cbf0dec974bf49844d7b72bb08804` started in the approved window ending at 08:55 CET. Commands, File Operations, Viewer PE Interactive, Viewer SQLite Interactive, and the two Viewer Text dialogs are scheduled. The native results will be added after the owned runner drains.
