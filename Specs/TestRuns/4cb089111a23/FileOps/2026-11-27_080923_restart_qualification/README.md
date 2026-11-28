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

RedSalamander merged the tooling repairs and shared engine pin in PR #357 at `3f2b01251c7e4bc8f8483ebbd8460444174cf1d1`. The test-enabled x64 Release rebuild took 7 minutes 36 seconds and had zero warnings and errors. All 188 output records matched receipt `946bd916ef3e93a02ca5abcf9868baa1f45f8eeac6462611d1e8982047a1b47c`. Fresh CI run `20261127T080428Z-67632-151d4572a667447aa54cc951dd320513` started at 09:04 CET under the approved follow-up window ending at 10:54 CET. The runner retains stdout and stderr in capture `20261127T080427Z-71012-6934f4b8af8642538e47be24ed795227`. The hosted Release jobs for the same source began native regression. [PR #357](https://github.com/RedSalamanders/RedSalamander/pull/357) [Hosted qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/45960884674)

The integrated run promoted Tooling Pester with 1,500 passes and zero failures or skips. Terminal result digest `2974f819fb8b21cf75cdea37a08a2e9aba3f95c0573cbd5b3a2be21573460ae0`.

Commands promoted with 965 passes, no failures and five skips. Terminal result digest `4e20b346253d5e8394ec87a10c15ae32ac5e718a56c10cb2d4d5b7ff5cea4bab`.
The Commands result includes passing Pack and Unpack dialog checks, the ZIP worker round trip, the folder drop guard, and both cases affected by the previous ARM64 foreground failure.

Plug-in Contract promoted with digest `29ff826b183fef266eeed38a15ca00fbca342c25c57e82ca2e5d525722e3cce8`. Its S3 multipart performance and teardown export reported 50 passes, zero failures and `S_OK`. This is the same export that reported 49 passes and one failure in the earlier Full run.

Viewer PE Noninteractive promoted with digest `154624703cd67008b72ff7389bffef11ba93d52f957a2978c6bcb45dfdd34bd9`. Its raw output confirms the ViewerSpace window stayed responsive with `ReadDirectoryInfo` blocked, Close returned within 500 ms, and the provider exited after release. This witnesses the prior failed scenario on the newer build. The earlier Full failure remains in the record.

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

The contained launcher stopped the classification retry at the end of the 10:54 CET window. Its trace reached `R0fSmb_NestedTreeCopyToShare`. The runner did not finalize the retry. All four interactive viewer entries remained pending, and the owned process tree drained. The canonical raw run archived with disposition `interrupted` and the launcher streams archived separately.

The fix source `d8b85fac0684204f91649f1ee0c0a89687e526ae` built the test-enabled x64 Release solution in 11 minutes 55 seconds with zero warnings and errors. Receipt `ce5008d9adb017bc9c0b9b38bc7a9d152be0797a5eea5465be0f430ac2c2385f`. Static passed, and 66 selected runner-plan checks passed with zero failures and skips in `20261127T095057Z-42016-754198165825437ba521135d90ef5e15`. A broader standalone Pester diagnostic had 127 passes and two failures while evaluating extracted affected-planning snippets outside the runner context. The dedicated plan result covers exact selection, casing and wider scope refusal. The earlier canonical full tooling entry passed all 1,500 checks on `3f2b01251`.

### Hosted ARM64 regression after the tooling repair

[Hosted ARM64 Release job](https://github.com/RedSalamanders/RedSalamander/actions/runs/45960884674/job/140378452721) completed with failure. Tooling Pester passed 1,499 checks with zero failures and skips, and Compare Directories promoted. The original Commands child exited 1 after 964 passes, one failure and five skips. `cmd_viewer_space_hover_does_not_rebuild_static_cache` recorded `ConsoleWindowClass` under the pointer, no hovered item, and no painted tooltip. The desktop boundary refused restoration over an external console, leaving File Operations and the four viewer entries pending. The hosted archive has no completed repository aggregate for ARM64.

The retirement record contains a SHA-256 of the executable path that matches the standard `C:/Windows/System32/wsl.exe` path encoded as UTF-16LE with the observed casing. The parent had exited before the diagnostic could verify its identity. The repeated console observations are consistent with the runner image issue. The upstream ARM64 image fix remains open as of this run. [Runner image issue](https://github.com/actions/runner-images/issues/14264) [Upstream image fix](https://github.com/actions/runner-images/pull/15362)

The Commands result SHA-256 is `f7c6f287ea5082347e866fda8ff8b4707204738998b23807d1b4d4673db6551e`. ARM64 artifact preservation run `20261127T095914Z-73520-630da9bf9bc5425aa087ed666e9cf389`. Job log preservation run `20261127T100117Z-70728-79864a9faffc45d2900fd159f80999c6`, SHA-256 `39a6c4e4fac2f3cfa7fd362512895900dbcc312c0ed8dd6eab3a99f3d56faee9`.

### Focused Local Delete repair

Five independent native runs passed Setup, `Phase7_DeleteItemsNestedRootsWake` and Cleanup with zero failures and skips. Each run verified one successful completion for each of four selected roots in twenty rounds. The fixture emitted 31–47 ms for the twenty rounds. The same-machine samples describe the test fixture and expose no disk-throughput claim. All five runner exits were `2` with filtered repository coverage.

| Governed x64 Release run | Focused duration (ms) | Fixture metric (µs) |
| --- | ---: | ---: |
| `20261127T100117Z-55752-477a4780e3d44e75924f4d50e9ee0ca2` | 78 | 47,000 |
| `20261127T100349Z-54556-e1baf3ec1c9c488fb8b095e9f6d8cba8` | 62 | 47,000 |
| `20261127T100438Z-59028-49ee52af45604a95aa44009d239de0e0` | 63 | 31,000 |
| `20261127T100528Z-60180-3fd4fadbc7324b40aa64310512f32696` | 94 | 47,000 |
| `20261127T100620Z-70736-8a1017dcc8944ca490b44281c0962726` | 78 | 31,000 |

The initial wrapper supplied an ID associated with the still-live wrapper process. Its native child passed three checks before the sandbox correctly refused archival. The five listed runs use their runner's own identities and completed archival.

The fix merged at `3a82609a0f07313b7707d7fc7dbbfc9e889b082e` after all required checks passed. With fresh desktop agreement, the broader x64 Release File Operations replay began at 11:08 CET on the exact repair build. Run `20261127T100829Z-71032-3a35034d0a544ae7a4b3aa41840a24cf`, launcher capture `20261127T100828Z-63940-23509abeeab44bb69fc99dfb7893995f`. The approved end is 11:58 CET. [Merged PR #360](https://github.com/RedSalamanders/RedSalamander/pull/360)

### Live SMB namespace witness

The first x64 Release compile rejected the wrong completion member in the test recorder. Corrected `843f88712e819a8410b02691eb261d631e53b906` reads the captured source-row receipt. That test-enabled x64 Release solution built in 7 minutes 45 seconds with zero warnings and errors. Verified receipt `78dda21b7329c33122cc0e8b6d02f165d2f01d44a53e7635be938e4230720ca4`.

| Live endpoint and identity mode | Governed run | Result | Same-name decoys (µs) | Replaced parent (µs) |
| --- | --- | --- | ---: | ---: |
| RubyGloom Samba, natural missing exact ID | `20261127T101512Z-12936-a57a7c13a8984420aa42842a33aa8be2` | 3 passed, 0 failed, 0 skipped | 271,281 | 146,032 |
| Windows SMB, controlled missing-ID path | `20261127T101702Z-75836-df53b4ff675e49cd8f24cedccc0f31c9` | 3 passed, 0 failed, 0 skipped | 28,537 | 23,263 |

Both focused witnesses kept ancestor decoys byte-for-byte and refused publication after moving the original file back through a replacement parent. The test also exercised 100 provider Renames on each endpoint. Both fixture cleanups passed. The Windows launcher restored the sandbox resource manifest to its original bytes.

Raw result SHA-256 values are `e344c7c1e5a562858daab048a086ab606c7251b7c61ba236782708675c30569f` for Samba and `f3a2994071aa556c45e89d0b10294eeca67615169cae69e789c885abf105d702` for Windows SMB. Launcher archive `20261127T101454Z-22036-f4fc735368884ac684d575833878c2c5`. The protected merge reached `f43444dcbdd4d77c3edcd4fd8ed108f9fbfea41c`. [Merged PR #359](https://github.com/RedSalamanders/RedSalamander/pull/359)

### Hosted x64 regression after the tooling repair

The [hosted x64 Release job](https://github.com/RedSalamanders/RedSalamander/actions/runs/45960884674/job/140378451388) succeeded on `3f2b01251`. The canonical archive records 3,026 passes, zero failures and 72 expected skips. All 19 CI entries promoted, including File Operations and the four interactive viewer entries. The Release assertion probe passed. Result digest `5efb744898ad468fe08d16486d18e09b1cd27a307de83913ecd229418dcac1e6`. Raw evidence archive `20261127T102332Z-65416-6641c4f2c663406e8c0f8bde7d74c2bb` and regression receipt archive `20261127T102237Z-41692-1ab0392960054e2a8da6da1b7e3b8950`. The combined test gate and both production Release branches did not run after ARM64 failed.

### Dedicated viewer desktop selection

The exact selection in PR #362 passed 68 planning checks in `20261127T101832Z-64300-051ffb0f71c0471bb77f612885fbfd13`. Read-only ExplainPlan `20261127T101832Z-76544-fb0f2647358041c49cf98040b042d5d8` listed Viewer PE Interactive, Viewer SQLite Interactive, Viewer Text Find and Viewer Text Go To in canonical CI order. The focused source and shared-scope suite passed 252 checks in `20261127T102838Z-68900-203427e89e3b406fa68a6989cccf44de`. The test-enabled x64 Release candidate `36c07a80367f3986f05172fcd89b7b75aa5a4855` built in 7 minutes 38 seconds with zero warnings and errors. Receipt `e03407c5c0b2bc4cd0d1d3cfac87e3fe889c6a52665dac1a008b5e5edacaf0fb`. [Viewer selector PR #362](https://github.com/RedSalamanders/RedSalamander/pull/362)

The four-entry viewer run started at 11:53 CET and drained at 11:58 CET, inside the approved window ending at 12:05 CET. Canonical run `20261127T105332Z-79564-ae0ecf66a8a84f0a8379ff50bd2775aa` executed all four requested entries in order and recorded three promoted entries and one failure. The runner result digest is `4b64ba4293a5e953886095287f1b1fb762f1f645b8d1462353cd3a87f2d95fd1`. Launcher capture `20261127T105330Z-42956-4fe4f0b48b5f49feb29ae4fb9a35c5aa` has stdout digest `341bd12bc4646159002c62aac124dd607afc3d4a3f9db827556b3db10436592a`.

| Entry | Exact x64 Release observation | Output SHA-256 |
| --- | --- | --- |
| Viewer PE Interactive | The VLC and ImgRaw children exited normally in the original and retry respectively. Each parent then observed changed input and an unrelated foreground, refused restoration and stopped before the next child. | `1a3c2ec2e82ba429534ac1e9416b0590057889b9e7eb8091661c364ca483db75`, retry `9890960fcb004ba20fcc88bf5c30801001d54c77ff2bc3ca6719948dab81175f` |
| Viewer SQLite Interactive | Entry promoted after exit 0. The executable reported zero compiled interactive cases. The Tab assertions require the Debug viewer snapshot and paging hooks. | `3354a43ecd8c96552bdd39a8c9519c9d50bfaf33d70c4dfa50a31ccd34b95887` |
| Viewer Text Find prompt | Passed the owned DxUi dialog, focus and close assertions. | `a22bb5fbf9fffa79b3da92fa871b058be1f2a30acf78fdef99f28320659d230c` |
| Viewer Text Go To prompt | Passed the owned DxUi dialog, focus and close assertions. | `9f6136fa7c1994c61185726324cbad79f7434af332e77b45d9ea733f2411f49b` |

The same Debug-only SQLite scope applies to the promoted Release entry in the hosted x64 aggregate.

The Debug solution built on the same source in 6 minutes 31 seconds with zero warnings and errors. Debug receipt `eb9fa049e81ca26846afca61fcc4e6c6ceec46d9c521dfa46fd01ea6e7e6f707`. All 188 outputs in each profile verified before the follow-up. The original PE logs identified the foreground process IDs and input timestamps. Their still-live instances started before the viewer run: the initial saved Visual Studio Code window gave way to Edge, then the retry observed a separate RedSalamander instance outside both test profiles.

### Repaired x64 Release File Operations replay

The approved File Operations replay completed and retired its child at 11:52 CET, before the 11:58 cutoff. The native suite reported 300 passes, one failure and 29 skips in 42 minutes 34 seconds. The nested-roots delete case completed its twenty rounds inside the suite and emitted `FileOps.SelfTest.DeleteItemsNestedRootsWake` with `S_OK`. The earlier heap exit remains in its separate archive.

The only failure was `R4A19_DiscoveryIndependentVolumes` with `alternate-volume-test-root-not-authorized`. This replay omitted the explicit alternate-volume flag and the case failed before allocating on another drive. A follow-up must select the exact marked alternate root and pass `-AllowAlternateVolumeTestRoot`. The run keeps its failed suite verdict. Native result digest `5ddefc87c5f34e600c9c3c998475bc0b13c69281d9706a748477f987b5681cde`.

### Independent-volume recheck

After the viewer processes drained, the approved five-minute window began at 12:01 CET. The same repaired x64 Release receipt ran only `R4A19_DiscoveryIndependentVolumes` with `-AllowAlternateVolumeTestRoot`. The launcher verified the existing marked NTFS C and ReFS Z roots on distinct fixed volumes and restored the original primary resource registration byte for byte.

Run `20261127T110136Z-62592-b63a350bca84466094e2bdebe429c9e7` passed setup, the 3,782 ms independent-volume scenario and fixture cleanup. Native result digest `e02a9f997398e82d35bb23d1078b6d1df8173456d95d8cb667a1705a61344ec4`. The launcher drained with filtered repository verdict `NOT_EVALUATED` and exit 2. Capture `20261127T110135Z-69404-c1c97542efd146918c1c41870a44822c` has stdout digest `8132c9b24c8e784d45da0b3685ef41d4064284b1419dc9ffaa6f4cc2131d7b13`. The sandbox audit listed thirteen pre-existing legacy or stale paths. The isolated 16 MiB-per-task control recorded independent progress overlap. Both streams took 1,172 ms concurrently, with 1,203 ms for the slower isolated control. The original broad File Operations run retains its failed verdict.

### Hosted ARM64 Commands callback failure on the SMB Rename merge

[Hosted ARM64 Release job](https://github.com/RedSalamanders/RedSalamander/actions/runs/45981936673/job/140448172187) failed on `f43444dcbdd4d77c3edcd4fd8ed108f9fbfea41c`. Tooling Pester passed 1,499 checks with zero failures or skips. Compare Directories and the non-interactive Viewer PE and SQLite entries promoted. Commands entered the owned Default desktop and launched native process 3752.

At 12:13:07 UTC the contained Commands child exited with `0xC000041D` (`STATUS_FATAL_USER_CALLBACK_EXCEPTION`). The launch record confirms job closure before desktop retirement. The foreground guard then observed a foreign `ConsoleWindowClass` process and refused restoration. Its path digest `95a9bdfa9774cbec665bce075b37ff65cb5cf6f70e0286eabad32eb453fdfbd1` matches the UTF-16LE spelling of `C:\Windows\System32\wsl.exe`. The outer runner separately recorded a foreground process named `wsl`. The evidence does not establish what raised the native callback exception or whether the foreground change contributed.

The guarded archive refused the retained `gj_…/links/j` reparse point inside the Commands fixture. That fixture name belongs to the Go to link target junction case. The uploaded diagnostics contain no command result, last-case trace or dump that identifies the crashing callback. The canonical run `20261127T111125Z-2752-299fffdc9454406cadccf0e6ab7c0481` therefore has Commands marked running in its last saved state and leaves File Operations and the four interactive viewer entries pending.

Original evidence and workspace uploads are preserved in `20261127T121523Z-75496-146a377aaaf542b6af44be3477ea2411`. The archived run-state SHA-256 is `c2e54e26815880519af33a53337fdf66e67ee311fe3be2bb2b4cf8c4f2a9f91d`. The complete job log is in `20261127T121553Z-64572-219b8e83cbdc4f078a2a5dafca736134` with SHA-256 `1ff5ca67e6d1674376cdb83c5cf889062286fd42fdd419e4b4f85c30f4be7c26`.

### Hosted x64 success on the SMB Rename merge

[Hosted x64 Release job](https://github.com/RedSalamanders/RedSalamander/actions/runs/45981936673/job/140448172110) succeeded on the same `f43444dcbdd4d77c3edcd4fd8ed108f9fbfea41c` source. All 3,027 required tests passed. There were 71 expected skips and zero unexpected skips. The Release assertion probe passed. The native CI run `20261127T110453Z-6456-f1cac1749ded4ce5a0c44847f57623b4` promoted 19 of 19 entries and reported exit 0. Verified hosted build receipt `278c3119885a8f9532ec24f79122fd3c4f4d48e4cc2d9152283794ca7cc9826b`.

File Operations reported 294 passes, zero failures and 36 environment skips. Both the independent-volume case and the repaired nested-roots Delete case passed inside that suite. The x64 Release SQLite executable again reported that its interactive group contained zero compiled cases.

Download archive `20261127T123447Z-79056-5b3ef4e8c2a3483db68a6121d5d45438`. Native summary SHA-256 `a87ed25cdb65c686605da0fadfe53dcaff4685febbf3acf3c87b72d67bc6628b`. Regression receipt SHA-256 `91183a13591b8ee8cc3cc871305734c5884689d531f2eca802182b0968b073ec`. File Operations result SHA-256 `889983445fc7ff6827d76bc28d38c7bc88d1b16c05bdf80c230b11f3fd196454`. The combined test gate and production jobs skipped after the ARM64 failure.

The promoted Commands result includes `cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation` passing in 2,284 ms. Its ZIP and 7z folder-selection assertions prove the complete archive entry set, retained original bytes and empty source directory, and no cleanup question or Delete when a tree receipt is unavailable. The Commands result SHA-256 is `2d9db08e76b5658fe2af21e18c7f813c0e00d7e26460d9734afd3beeb9bc28be`. The File Operations result also records passing `FileOps_ReviewRemediationContracts`, `Phase10_DeferredConsentAndRecycleEscalation`, and `Phase10_MetadataPreservationAndSourceRetention`. Exact Pack descendant cleanup and combined sensitive-metadata coverage remain open.

### Integrated non-interactive x64 Release PrePush

The integrated source `fc3fe2499e11ce70eb3b1fc5bc426e3e17bfa197` built the test-enabled x64 Release solution in 7 minutes 58 seconds with zero warnings and errors. The runner verified initial receipt `36e84c4cb767b1334ec51ca9ea629f9a91fa1c13b1d3bf0c22a32bbc40e9ae39`, passed all 35 build-tooling tests, and re-attested receipt `21aa51848139198445f680aead0e3885fb0f8419b25fdaa1c14ee968bdab735f` with another zero-warning, zero-error build.

Fresh run `20261127T115257Z-66576-ae56caf6de77484ebe13ecf8bda2b90f` promoted all 14 selected entries. It reported 1,801 passes, zero failures and 30 Compare Directories environment skips. The full Tooling Pester entry passed 1,501 checks. The archived change verdict is `PASSED`. The repository verdict is `NOT_EVALUATED`, and the runner exited 2 because Commands, File Operations, Viewer PE Interactive and Viewer SQLite Interactive were outside the non-interactive plan.

Canonical result SHA-256 `9458f6fbcf01a61b9b2e511be53172ff7d3ca833baa6bcb13a49ea732794175d`. Runner detail SHA-256 `ac69705f8969031bcac214dd41b944bbd35b3bfa8a6c662c97719234a613241b`. Both files are in the completed-run archive on the marked Z test root.

### Approved viewer recheck

Both exact viewer selections ran in the approved window on 2026-11-27. The launcher began at 14:33:18 CET and drained before the 14:45:18 CET cutoff. Source `36c07a80367f3986f05172fcd89b7b75aa5a4855`. All 188 outputs in each x64 profile verified against its receipt before launch. Release receipt `e03407c5c0b2bc4cd0d1d3cfac87e3fe889c6a52665dac1a008b5e5edacaf0fb`. Debug receipt `eb9fa049e81ca26846afca61fcc4e6c6ceec46d9c521dfa46fd01ea6e7e6f707`.

- **Viewer PE Interactive, x64 Release.** Run `20261127T133320Z-71308-8ddcb037ed434ffeba8082d889aeb495`. The entry passed, including the previously interrupted VLC and ImgRaw focus checks. The runner finished in 3 minutes 22 seconds. Result SHA-256 `0c21062df063527552c028569878e47fe542f2272bdea465c4ea3fbff86bb5ab`. Child output SHA-256 `3f040be7d005363149b0b361a070e968b8d9edbaf25632f092e9a4207d1f3b74`.
- **Viewer SQLite Interactive, x64 Debug.** Run `20261127T133738Z-54716-cd7d6cb85999494b8347c348771cf577`. `TestViewerWindowTabTraversalMatchesExpectedOrder` passed forward and reverse traversal, Escape focus return to the results grid, and closure from the grid. The runner finished in 17.8 seconds. Result SHA-256 `809048662c5c56b21cafac1173d3d3cf8b6a48b8d1df83e7f0c823feafe47d6a`. Child output SHA-256 `c81b0c87447a26597e7c77da10f44cc2b3a7b8bf883e055af2061a5cc3629c0d`.

Each selection recorded one passed entry, zero failures and zero skips. Both archives completed under the marked C test root, and the launcher and owned test processes exited. Each filtered runner retained repository verdict `NOT_EVALUATED` and exit 2. The prior Release-only run retains its original verdicts.

### Commands native exit diagnostic

Candidate `d636033db3e0a8e34e1c2fe3dfe59f9fdbb12fa0` records the most recent complete case entry or return after the original Commands child exits with a negative native code. The runner emits the launch ID, signed exit and case boundary after containment closes and before desktop retirement. This preserves a bounded observation in the console when the full archive is refused. The observation identifies the last case boundary available in the trace.

Focused Pester 3.4.0 run `20261127T135346Z-73832-01a4bff9c29d49b6bfb5d6bad44ef070` passed 12 checks with zero failures or skips in the marked Z sandbox. The checks include the injected negative exit, retirement ordering, the fresh bounded UTF-16 tail, unsafe and incomplete inputs, and original versus retry scope. Log SHA-256 `52977bdc0fdebc6715be738aa29b4ebb1e71d5bf205cc8eb516d84d82490bcc8`. A read-only replay of the real x64 Commands trace preserved by rotation recognized the return of `cmd_pane_copy_text` from the last 64 KiB of the 999,718-byte trace.

A direct broader Pester launch lacked the canonical runner's module, planner arguments and exact native profile. The 397 passed checks and its failing fixture preconditions are retained in `20261127T134906Z-76344-c2125eb6a5d54d5482d13f2ed52e51df` and `20261127T140251Z-72584-d3ecb27611c047afaef27be4e8b74a0b`. Supplying the normal runner module and planner context passed all four affected-planning fixtures in `20261127T140520Z-70432-b8e09347cb3d4c549456a521bcfdfbad`. Canonical selected run `20261127T135417Z-53784-abd16f94c2fc4e14ac88dba47dc5c340` built test-enabled x64 Release on the exact candidate and bound receipt `1294c2b395a3f7f33c9537372048facc3f82c0d1d7d66b5fed3dc1ed5562c67f` with 188 outputs. Tooling Pester promoted 1,504 passes with zero failures and zero skips in 1,351.6 seconds. It completed archival with the filtered repository verdict `NOT_EVALUATED` and result exit 2. The scoped adapter returned 0 for the successful selected entry. Detail SHA-256 `dc4301325aeb10b29959734a994622acaf85c2d498873b60c71d8d6c39cd4744`. The [diagnostic PR](https://github.com/RedSalamanders/RedSalamander/pull/370) merged on `a9ca4a6472ff76212614e32a0e9b0ff03e57e38c`.

### Hosted ARM64 foreground failure after live-provider qualification

[Hosted ARM64 Release job](https://github.com/RedSalamanders/RedSalamander/actions/runs/45993909744/job/140509781386) failed on `8c3c568016794ce90178cafe56eea29f426a10e8`. Tooling Pester passed 1,501 checks with zero failures or skips. Compare Directories and the non-interactive Viewer PE and SQLite entries promoted.

The Commands child ran on `WinSta0\Default` for 44 minutes 23 seconds of native case time, then exited 1 and closed its containment job. The original native suite reported 963 passes, two failures and five skips.

- `cmd_pane_fileops_popup_throughput_disclosure` lost real foreground before Enter. Focus remained associated with the owned popup, while a `ConsoleWindowClass` window from a different process held foreground.
- `cmd_viewer_space_hover_does_not_rebuild_static_cache` observed a `ConsoleWindowClass` window under the test pointer. The failed pointer sample belonged to the console.

The hosted boundary refused foreground restoration. The outer desktop sample identified a separate foreground WSL process at retirement. The native case evidence identifies the earlier window class and the first case's PID independently. Canonical run `20261127T132325Z-5580-c78f80db1e174bd0897100ae0dfb2ec6` saved Commands as running without a promotion, while File Operations and the four interactive viewer entries remained pending. Its retained build receipt is `caa12463184a324ef994fabe20ff65c5b3d9d0f803edfe2095278e7dbd294e12`.

Commands result SHA-256 `172b173d5644c74a6f438e68c899c5c82baeef86a08169ffad631bca210741ee`. Canonical run-state SHA-256 `61144db8c7db98660b7db6085eb325d1ea0929318b793c4bb246a5c38e88d4be`. Hosted evidence and workspace import `20261127T150203Z-39044-9b1cafb5ab2e4326bdf52f86d894a89c`. Complete job log archive `20261127T150301Z-43024-0539b2767ed943d386bb212c44c8c2c6` with SHA-256 `e255cb131ec763f2acc456d78fb87978190b22e30f17f4216de8887620d589e1`.

### Hosted x64 success after live-provider qualification

[Hosted x64 Release job](https://github.com/RedSalamanders/RedSalamander/actions/runs/45993909744/job/140509781222) passed on the same `8c3c568016794ce90178cafe56eea29f426a10e8` revision. The regression executed and passed all 3,028 required tests, with 72 expected skips, zero unexpected skips and no failures. The Release assertion probe passed. Native CI run `20261127T133015Z-2032-6dda6fd0c29b4b0882a3bf27e0647d66` promoted all 19 entries and reported exit 0. Build receipt `2375463eb2247b321ac64cecb4873057ec16671b16ca12426dc4b705c6ae433d`.

Commands reported 965 passes and five skips. Both the path-detail and ViewerSpace hover cases that failed on ARM64 passed here. File Operations reported 293 passes, zero failures and 37 environment skips. The independent-volume and nested-roots Delete cases passed in that suite. The combined test gate and production jobs skipped following the ARM64 failure.

Hosted evidence import `20261127T151601Z-85572-911f2b54a6af442a8310e4beb7b0515b`. Canonical summary SHA-256 `af52a48fa45523315013421fb7f32f7f3830cfee47a5ae9317304e75fbe11a26` matches the hosted regression receipt. Regression receipt SHA-256 `f3e24dc62bb0f52aa76d8cde199fc431c886f669b3f0a2a11a9825d2d95746f2`. File Operations result SHA-256 `9e19ae1c5edf4b2897790bd0288c9db4f9df6eee75947b75cfc6130391fc5f9d`.

### Hosted qualification after the Commands diagnostic merge

Hosted replay `46020714850` evaluated source `42691d6bc0c26a84b815ee0f9b43eb2de7fe682a`. Both native jobs failed and the combined test gate and production jobs skipped. [Hosted replay](https://github.com/RedSalamanders/RedSalamander/actions/runs/46020714850)

**x64 Release.** Native run `20261127T160249Z-6956-f9ff4630348d4f1e8f5e3adb07609cf8` reported 3,030 passed, two failed and 71 skipped. Tooling Pester reported 1,503 passes and one failure in the original Commands trace witness. Its retry passed all 1,504 checks. The non-interactive Viewer PE entry failed the ViewerWeb source-origin assertion and passed that assertion on retry. The original failure log also confirms the private-route classification, script policy and isolated browser profile. The runner classified both entry failures as flaky and kept the failed verdict. Commands and File Operations reported zero failures. Build receipt `24b7bf2ac167772cbd70e649ed33da378ea802b310831d0eb4a620ae1757e93c`. Evidence archive `20261127T180731Z-85888-42014651be7f4d9f9c1ed8dfe2c8d389`. Aggregate v1 SHA-256 `7362a24bd056a427b4f50857e1f8d622a18fd40df09f1ea73dc203c69e7f7bf2`. Complete job log archive `20261127T180637Z-14176-ef972af907f54abdb0d564c36d8fd811` with SHA-256 `4ac7e5143b20c3194c9c072c99fcb90b78b101d5e247fcf91992b61a82f1fd93`. [x64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46020714850/job/140589147450)

A fresh local non-interactive Release Viewer PE selection on `36c07a803` passed the ViewerWeb source-origin assertion and the selected entry. The Viewer PE and ViewerWeb source directories match the hosted revision. The run `20261127T180858Z-81440-b6bf2626a8ff4d67aa105136c1802994` retains filtered repository verdict `NOT_EVALUATED`. Child output SHA-256 `f3c3c71b5e407f5e4c5cb5ff21a8dcc025b64740350b0907e81458f1eee8b050`. The logged assertion gives the successful predicate and test profile. The hosted failure log recorded the predicate without its observed URL, leaving the cause open.

A second local run with the same verified x64 Release source and receipt also passed the actual private-origin assertion and the selected entry. Run `20261127T183920Z-77820-e8623c12853c450f8092e98e65bafdd1` retained console SHA-256 `0002abb86140661631f4a69e6405e6b9fd5490866ee36c543ea5c87065350bed` and aggregate SHA-256 `2fb184d036e2d784bc80dec24349dcc229310c6ec682fda487fd33799308e43b`. The filtered verdict was `NOT_EVALUATED`.

**ARM64 Release.** Tooling Pester passed 1,504 checks. Original Commands reported 963 passed, two failed and five skipped. A foreign console took foreground while settling `file_action_resolution_v16_explains_priority`. The navigation refresh case failed to capture its baseline. Both isolated classification children exited 0. The shuffled Commands replay again reported 963 passed, two failed and five skipped, this time with an active-pane focus restoration mismatch and a speed-limit test that could not quiesce pre-existing tasks. The shuffled child exited 1. The desktop boundary refused restoration and the adapter observed a foreground WSL console. Canonical run `20261127T160337Z-5880-806c0eb22a45424a82b47d638fd11dda` retained Commands as running and File Operations and interactive viewers as pending. Build receipt `671e495b3d0d494e3785d9ff954d4672088b92e0269add842e052e7ea4a399f9`. [ARM64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46020714850/job/140589147439)

The stock-preparation audit retired an exact WSL startup window at 17:25:41 UTC (PID 11188, verified argument and creation-time match). The later retirement refusal at 18:04:25 UTC captured PID 8864. Its path digest matches the System32 WSL image and the retirement digest from run 45993909744. Its parent was already unavailable. The outer adapter observed PID 11216 in a subsequent sample. These are distinct recorded process instances. The creator and its relationship to individual Commands failures remain undetermined. GitHub's image issue and proposed package fix were both open at the follow-up. [Runner image issue](https://github.com/actions/runner-images/issues/14264) [Upstream image fix](https://github.com/actions/runner-images/pull/15362)

ARM64 evidence and job log archive `20261127T181319Z-81228-0fa85b30c40442d586cac1ef1ec8a853`. Original Commands SHA-256 `0c1045e9e029e59f546a47f9bd1e2b123f3a9624cba50ee30f73c865de1733f6`, shuffled Commands SHA-256 `5beb251f943a0b22d97af46c7707f81c8e0244a594614d3d52a0a9610fc12c0c`, run-state SHA-256 `ae8c643adb0e2f0673cf67ed10117106c213f6e771ba63b1fa4aeba9a3936e7f`, job log SHA-256 `97d30b7afde04eaf11a988a4d4cf2a48425a4b699178d01c8f802b5ced6df247`.

### Commands trace freshness correction

The marked NTFS probe `20261127T180730Z-82940-9bd062c88731440ebc3c36c88048b266` saw a file write time precede the launch clock read in 43 of 1,000 writes. The oldest observed delta was negative 0.4189 ms. A deterministic fixture with an earlier write time reproduced the missing post-containment record on the old code. Reproduction `20261127T180954Z-60000-babba20055554df0bf25778ee94851b2` archived two passes and the expected failure. Log SHA-256 `773494bb0445d009746514edaa2b5b95d6c040fb5f53941bda0a2552e487414c`.

Candidate `4df8b9dab` compares the guarded path's file state before launch and after containment. It permits a newly present trace or one whose observed length, creation time or write time changed, then reads the same bounded case tail. Focused run `20261127T181149Z-83140-676f3112768e4433b486a09535faa9d8` archived 12 passes with zero failures or skips. It verifies the earlier-timestamp reproduction, same-time changed file, unchanged-file refusal and the existing diagnostic ordering. Log SHA-256 `20084a9495731055e7a8ae8a5076250a5f69e20b777141acf23aedd503c6969d`. The first exact x64 Release Tooling Pester selection `20261127T181320Z-85012-28f20f5231de404e946080717581d5e7` failed with 1,518 passes, one failure and zero skips. The inventory desktop-admission fixture omitted the artifact directory supplied by the real runner. Aggregate SHA-256 `e476d5f1bebb6515d3db12d2c42a14161ae72d7424eea125ec2f6ec7efbabdcb`. Candidate `aee6062d2` supplies the owned fixture path. The exact failing test and three boundary checks passed in `20261127T184608Z-41996-e350a2921e9846f08bcaa2f3efa959ff`, log SHA-256 `d09a955b6c74e2ea8e2171e40aa57734812960f6654d8842386e7d67fc93db39`.

The corrected x64 Release build completed with zero warnings and errors on receipt `109b06e1fcdded599f7f1c651b51bc827e47b8a3b3217964a0722aec8b81234e`. Canonical `tooling.pester` run `20261127T184627Z-65804-e8c0bfe54e0b41a19382d1a859ef494f` passed all 1,519 cases with zero failures and skips, promoted the tooling entry, and retained the filtered repository verdict `NOT_EVALUATED`. Aggregate SHA-256 `0ea2bc702eb0dbf92783810105885d620206442b18e0b671e65b1bfb72f482aa`. Run-state SHA-256 `9f244715bb9b913565f778ce61de78d340f0a15967d9396f21c270bb716e79b9`. PR #375 merged at `320f1dc04bd9478e03f885c382a635eaed5649bb`. [Trace repair PR #375](https://github.com/RedSalamanders/RedSalamander/pull/375)

### Pack shared preparation and retained-originals alert

Candidate `58369588d40bb913727e1d7b33607bd0c3619009` prepares the source inventory once in the Pack worker and validates it against the approved cleanup receipt before selecting an encoder. Both stored ZIP and 7z use the prepared entries. The pane alert identifies complete publication with retained originals. The incomplete-packing alert remains specific to omissions. The worker records `archive.pack.prepare_us` alongside the existing Pack worker and publication metrics. Completion through an exact directory cleanup receipt remains an open gate. [Pack preparation PR #376](https://github.com/RedSalamanders/RedSalamander/pull/376)

The baseline ran three repetitions of `cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation` without activation in x64 Release. Candidate and baseline use the same archive fixture. Baseline source `36c07a80367f3986f05172fcd89b7b75aa5a4855`, receipt `e03407c5c0b2bc4cd0d1d3cfac87e3fe889c6a52665dac1a008b5e5edacaf0fb`. The relevant archive implementation and native fixture match the pre-change candidate base. Run `20261127T181920Z-81056-a85426aafab843ea82e27050ef435ef7` passed 3/0/0. Native case times were 2,371, 2,021 and 2,055 ms. The retained metric sample records 44,117 µs for Pack, 66,782 µs for Unpack, 16 entries and 23 bytes. Metric SHA-256 `15ae232724d81b32464502c2c293c88f94de8f1f1226d720a2ccadefe4200157`.

Resource localization run `20261127T182330Z-81076-0082ac0346a64c4899069145f66d7f3c` passed 12/0/0, including positional placeholder parity and satellite completeness. Log SHA-256 `647bd51ff87ddd07fc39d356523f051618d687bf0838044212fc6f8b33bff22d`. The pinned formatter checked 650 files.

The candidate selected-project x64 Release build finished with zero warnings and errors. Artifact receipt `154a13972c494ec4eca9f62e2ebe813b234b5ed69db77e243b11de3d670a66ab`. A fresh non-activating selection with `-Suite Commands -CaseFilter cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation -SelfTestRepeat 3 -Configuration Release -Platform x64 -NonInteractive -TestRoot C:/RedSalamander.Perf` passed all three repetitions with zero failures and skips. The test asserts the actual pane alert and source fate for complete archives with retained originals and for archives with omitted items in both formats. Run `20261127T182409Z-68868-adb9e6d1f7da4031971556cc858eb724` archived the native result with SHA-256 `80d31b9ca429d777567357550dc74db7e7e082311df22b0d0372789f35e09eff`. The filtered repository verdict is `NOT_EVALUATED`. Both baseline and candidate sandbox audits reported the same 13 pre-existing shared-root paths.

Candidate case times were 2,149, 1,927 and 2,111 ms. The retained metric sample measured 14,657 µs for Pack and 67,551 µs for Unpack of 16 entries and 23 bytes. Metric SHA-256 `6feeff39e4499e1ed8756599b2d295f40459eecc659006b73f5d2ea4c4c0d5cb`. Successful candidate preparation samples had upper medians of 511 µs across 21 ZIP samples and 476 µs across 15 7z samples. Raw performance log SHA-256 `f7a3a3001ef2513ccd00b5f7fc1c327bbb081dbe48d34c9a7cbaed0a4bcfdb60`. These observations describe this short fixture.

PR #376 merged at `ab452399690fb7e152188b377dfcdd31a00cafc7` after the required checks passed. [Merged Pack preparation](https://github.com/RedSalamanders/RedSalamander/pull/376)

### Combined EFS and sparse allocation consent

The existing worker checks EFS and sparse allocation together when applying pre-content metadata, but used to request each consent separately. Candidate `196c02d8dc53e8b1c30b1201e0cba8611fea5a2a` reports both losses in one question before streaming file bytes. One combined risk key keeps any repeat decision within its existing task and destination scope. The Move action can publish the copy and keep the original. Placeholder hydration remains an earlier separate decision. [Combined metadata PR #377](https://github.com/RedSalamanders/RedSalamander/pull/377)

The preserved x64 Release baseline `36c07a80367f3986f05172fcd89b7b75aa5a4855` passed all three non-activating `FileOps_ReviewRemediationContracts` cases in `20261127T185315Z-80428-acad43c347a34d8bb1762705e5f5e13f`. The native result SHA-256 is `a8daaccfaf93d08cbc960ff4778c82098bb50005373fdb3e7bce1a1c73543557`. The existing remediation body took 106,750 ms and used build receipt `e03407c5c0b2bc4cd0d1d3cfac87e3fe889c6a52665dac1a008b5e5edacaf0fb`. The separate Phase 10 baseline selection `20261127T184354Z-82068-c716e5616f1f4cae82a78bc8500b6500` required desktop admission and executed zero cases.

The exact candidate selected-project x64 Release build had zero warnings and errors. Build receipt `35c4bdf6abb56bac20cce67fc5f0fd4a254303dadfb88674085a097e375ac65a`. Run `20261127T185356Z-72340-d09877195e8d413faf35edb0ad109941` passed all three non-activating `FileOps_ReviewRemediationContracts` repetitions and their setup and teardown, for 9/0/0. Each repetition observed one combined question before the final path existed, checked Copy and Move question text, and verified both payloads and retained-source results.
The native result SHA-256 is `72e296bb95a5ed5175fe253c166c714b4c47fcc819fb4e172b73be5eb101c996`. Trace SHA-256 `2bd158b8907d269b18e107e34ee35248653518ed3a401657cba4ef26bb208090`.

The combined witness recorded 81,392, 73,444 and 67,579 µs through admission, the single decision and the retained-source result. Every row recorded feature mask 160 and one question. Raw performance log SHA-256 `fc98fbfb5758d6ab8abe4eeaa4de09b0d45f2fb6171966a5f8221948d8d97586`. The remediation case also runs unrelated witnesses. Its full durations were 97,078, 95,813 and 99,984 ms. These brief fixture observations cover storage and prompt behavior.

Non-activating `FileOps_RepeatDecisionContracts` passed 3/0/0 on the same verified receipt in `20261127T190714Z-83800-1c560937a3ac48ee9c0c7aac2a5bad1a`, including the action placement and scope matrix. Native result SHA-256 `56d6d92c1ff7d60f7cbb8da05d29b1d4afe1c885fd8a5142e50b07a60495f9f3`. Both candidate selections retain the filtered repository verdict `NOT_EVALUATED`. The full metadata and conflict walkthrough still requires desktop admission.

PR #377 merged at `b7bea3b68fa8a12c7ab1f138c9412520e4d9bc87` after the required checks passed. [Merged combined metadata decisions](https://github.com/RedSalamanders/RedSalamander/pull/377)

### Follow-up hosted qualification

At 20:18 CET, run `46027552244` on `bcc3f42e2` had successful admission and both x64 and ARM64 jobs had reached their native regression and Release assertion step. GitHub had queued the later runs on Pack merge `ab4523996` and combined metadata merge `b7bea3b68`. [Integrated regression in progress](https://github.com/RedSalamanders/RedSalamander/actions/runs/46027552244) [Pack qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/46050953765) [Combined metadata qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/46055245540)

### Approved conflict and metadata desktop pass

The approved eight-minute window started at 20:27:34 CET. The timed launcher finished the three child runs in 121,985 ms. Each child acquired the canonical interactive-desktop lease and closed its contained process. All three launcher results were exit 2 with filtered repository verdict NOT_EVALUATED and native counts of three passes, zero failures and zero skips.

The test-enabled x64 Release build was source `196c02d8dc53e8b1c30b1201e0cba8611fea5a2a` with verified receipt `35c4bdf6abb56bac20cce67fc5f0fd4a254303dadfb88674085a097e375ac65a`. The launcher used the marked C sandbox. Its archived record `20261127T192734Z-79568-2b272342923b4d37bab23546c163eb8c` contains the three runner logs and the result manifest with SHA-256 `793191f857959cf19aa5da89b886ed711333e48eb4ee3944cef6a8f9c646a977`.

| Native case | Run ID | Case duration | Native result SHA-256 |
|---|---|---:|---|
| Read-only replacement | `20261127T192735Z-81880-aa5f51dead374712a4b7d7efa7df59be` | 281 ms | `647935623883231e9d88f2088a420112974aaa9c2c5234deb3d5e811155a77dd` |
| Retry after a known non-commit | `20261127T192815Z-57912-c5b7e026cf8741478fefbe0944f35687` | 2,532 ms | `db402d2353bdd0f6c3038e9a8ab1fcd88702ee8f165eaa981950d32b716833b0` |
| Metadata preservation and source retention | `20261127T192856Z-67616-4f9370d504244ef28804bdff1db07170` | 562 ms | `78f4c0df39b0a68872fd39ff51ad86a364a7d77ce4346a05f7012469e966b2bb` |

The read-only collision offered Replace read-only and Keep both. Explicit replacement cleared the destination's read-only attribute. The Retry fixture verified one attempt per explicit choice after a proved sharing failure and eventual completion. The metadata fixture verified preservation of alternate streams, Mark of the Web and sparse allocation. It exercised separate hydration consent, the combined EFS and sparse allocation choice, publication with the source retained, and Cancel before publication. The `FileOps.SelfTest.MetadataMatrix` row recorded all four scenario groups. Performance log SHA-256 `a2b69a06e8a2f9384c46fcc5938b0c85537e646621992ece33be1d8c8f85bf93`.

The selected pass covers these three cases. The wider workflow walkthrough and the integrated native gates stay open.

### Local archive source revision protection

Candidate `650c2ea733068bbec8c510a5cb74471b52f548f1` excludes overlapping content writes for an approved versioned Local reader. It compares the existing four-part revision on the same handle after obtaining its sharing restriction. Pack and Unpack use this reader after exact cleanup consent. Directory sources still require a separate exact cleanup receipt. [Archive reader PR #381](https://github.com/RedSalamanders/RedSalamander/pull/381)

The frozen baseline `196c02d8dc53e8b1c30b1201e0cba8611fea5a2a` with receipt `35c4bdf6abb56bac20cce67fc5f0fd4a254303dadfb88674085a097e375ac65a` passed all three non-activating x64 Release archive case repetitions. The relevant Local provider, archive command and native fixture inputs match the pre-change candidate base. Baseline run `20261127T193408Z-23604-e42017460a3d4119908785a73d9eb1df` has native case times of 2,330, 1,932 and 2,020 ms. Result SHA-256 `81f11b4241db21e697940023699bd3b16c0ba6b027180bf766f2c0e204fe5480`. Raw metrics SHA-256 `972532cc7a9497ccee66e0a7b28f2fe02a77ef99c3d75cd7ca911ed2a5828dfd`.

The first candidate build, a single-project invocation, passed with zero warnings and errors. Receipt `8957324f6e213acee1c3b9158da5b2109279ea49de39b1891c92c13a074b971a` carried the single-project target kind. The canonical replay `20261127T194656Z-4000-0b163d44b69c47ff970583fb4fbfab71` required the explicit project union and refused that receipt before launching tests. The retry delegates the build and receipt to the same canonical runner.

The canonical x64 Release project-union build finished with zero warnings and errors. Receipt `afeb761141ca932aa7ce6ae1361f82298b9fe5ec320d732acb4826079935eade` attests 166 outputs. Run `20261127T194823Z-73460-19917ccd9d6a4cbbbbc7e4cd3726d030` passed all three non-activating archive repetitions with zero failures or skips. The native case times were 2,315, 1,888 and 2,003 ms. Native result SHA-256 `f32972c5e8a3ed65ec9386375f22debb14c3c8142a3bd64f09d0143bd3b9255e`. Raw metrics SHA-256 `67f2e44e3aa5f782af768dca09e9582038af18b01c04acb6050b6407e5f2f18b`. The filtered repository verdict was `NOT_EVALUATED`.

All six busy-source Pack attempts, three ZIP and three 7z, reached the worker and reported `0x80070020`. The native assertions checked exact cleanup consent, preserved source and pre-existing output bytes, and no active cleanup. Each repetition also denied a new writer while its versioned reader was open and refused the old binding after a same-size rewrite.

The existing successful 33-byte exact-source Pack controls give directly paired worker samples for both formats. Baseline medians were 3,216 µs for ZIP and 4,288 µs for 7z. Candidate medians were 2,872 µs for ZIP and 4,287 µs for 7z. Both runs reported the same 13 existing sandbox audit issues. The samples describe three repetitions of the local fixture.

The same byte-verified receipt passed the non-activating `FileOps_ReviewRemediationContracts` selection and its setup and cleanup in `20261127T195517Z-75080-8bd77c30628d41fdad215dc0593c0720`. The case took 85,906 ms. The archived native result reported 3/0/0, SHA-256 `481dfc00a50781d0c8aebf0db4006ae06caeac508da64d1046b7e3082d1cafb5`. The filtered repository verdict was `NOT_EVALUATED`.

PR #381 passed the required checks and merged at `e72f5f02ff619264b6c3b2db7297a4a9ce154b73`. The six changed product, test and contract files match the verified candidate. [Merged archive reader protection](https://github.com/RedSalamanders/RedSalamander/pull/381)

### Hosted qualification during the archive reader merge

At 20:58 CET, both native regression jobs on `bcc3f42e2` were still at the actual regression and Release assertion step. The Pack and combined metadata merge runs remained queued. [x64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46027552244/job/140671267975) [ARM64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46027552244/job/140671268108)

### Hosted native failure on the pre-directory candidate

Both native jobs in run `46027552244` on `bcc3f42e2e114953556605f3ef54e30919ada1bc` failed. The combined test gate and production jobs skipped. [Hosted qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/46027552244)

**x64 Release.** The aggregate reported 3,045 passes, one failure and 72 skips. Commands passed 965 with five skips. File Operations passed 293 with 37 skips. Tooling failed its original-negative-exit witness in both attempts with 1,518 passes. Its assertion expected the Commands case boundary row to follow containment closure. [x64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46027552244/job/140671267975)

**ARM64 Release.** The original Tooling attempt reported 1,518 passes and the same witness failure. Its retry passed all 1,519. A foreign console was in front at desktop retirement. The boundary preserved that console and ended the native job before its complete aggregate. The recorded process was PID 884 with image path digest `95a9bdfa9774cbec665bce075b37ff65cb5cf6f70e0286eabad32eb453fdfbd1`. Its parent was unavailable. [ARM64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46027552244/job/140671268108)

The exported failed-step logs are retained in `20261127T215313Z-67324-3779e546dc064f278be27deaecaf2c5e` under `artifacts/selftest/last_run`. x64 SHA-256 `30e77cbb6a6c742dc25b0dad22930a8c3ebe62f34c46d7ee2d78ac1002da8980`. ARM64 SHA-256 `68296802faa1a0062e947d224a495c5c30c30dbe77cf34a453ab15987c599c66`. Run `46027552244` on `bcc3f42e2` predates the merged Commands trace freshness fix in PR #375. At 23:18 CET, both architectures in the follow-up on `651a700fc` had completed their Release builds and were running native regression. This run includes the trace fix. [Follow-up qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/46061664377)

The local Pester 3.4.0 replay of the three Commands case-boundary checks passed 3/0/0. Run `20261127T215840Z-53720-fda843ee669345418be047ea99da7b0f`, log SHA-256 `d51ecb90cb3b78b2d7e54d8dbbedd7040326ce8d6be0228b5789d7360ab9df75`. The archived hosted verdict remains failed.

### Exact Pack directory cleanup

The first candidate `57e881f3644149a6bcb47387977b533ca9617f28` stopped at compilation before launching tests. The C++ diagnostics and interrupted runner evidence are retained in `20261127T215019Z-78340-9f4de4e0984e432186bf289ed99cefde`. [Pack directory PR #384](https://github.com/RedSalamanders/RedSalamander/pull/384)

Candidate `4b2dbdfc1f9cbd830b207731fddcd545af65719a` built x64 Release with zero warnings and errors. Its three native repetitions failed before the new tree consent. The directory reopen returned `0x80070005` on the first child in each capture. Run `20261127T215511Z-80200-bf26980f2caa46c6927adbc8f00a4cd1`, native result SHA-256 `c524d669acb3ec48081e304a91f1d47c96d048d40e955f0eb22fc4144d0875bd`, raw metrics SHA-256 `8bf10c1c242e91ae897ed297b4ca9d132d40791c51bf014a19adbe1957878f74`.

Candidate `1b2dd5db66813c97982fc3616e379617864aa48f` uses the retained directory handle for its cursor and relative opens. Its x64 Release build had zero warnings and errors on receipt `b98fe2d2848ccd311bf9fc1421237b686bbb74ad22750a3997870eeee9119364`. In all three repetitions both positive formats removed five of five approved entries, the Cancel control kept the previous output and tree, and late additions and same-size rewrites kept five of five approved originals after publication. The remaining test failed at the attempted ancestor move while the source handles were open. Run `20261127T220413Z-79000-2552e3d1f6be45efa70c73ebe7ffd820`, native result SHA-256 `c4592acfd0d4817af6ab199db89734dfffeff7d84e0fd52233fc0f264d7aa816`, raw metrics SHA-256 `52a45e23c1b833f5dd80ef0b3053728c90f9491efc57b046091aa0a149dd080d`. This run retains a failed verdict.

The final exact candidate `af79501c8d7b0fd1c7ecc469f0e59316e74642a2` had zero warnings and errors in the test-enabled x64 Release build. Receipt `96a4e3da9f8bb0147d3eec1f70a0c1668003a02639bbcbadcfda342e3c90e35b`. Its canonical non-activating run `20261127T221543Z-82288-583e98b28f894d0a97c52b24cf6b9840` passed three repetitions with zero failures or skips. The native case times were 2,889, 2,905 and 3,080 ms. Native result SHA-256 `7f3bea4cda61c263ae9162ec6a9b987466de1c9b22fc8a36acfe795b33f2a497`. Raw metrics SHA-256 `26b922485afa853013e79bb56faeb630a3b7a209af0de84391587cf6a9b84771`. The filtered repository verdict is `NOT_EVALUATED`.

Each format removed five approved entries and the test read the approved nested bytes back from the published archive. Cancel preserved the previous archive and the entire tree. The late addition and same-size rewrite controls retained the published archive and every source. Windows refused all three physical ancestor moves with error 5 while the retained tree was open. The existing Local identity namespace then changed the fresh ancestor's reported identity. Preflight stopped at that ancestor with zero selected bindings and zero removals in each repetition. All nine post-publication change controls reported `0x8007051A` and removed zero entries.

The approved five-entry shapes captured five entries and eleven ancestor bindings in 392–679 µs across 18 samples. Successful ZIP preflight took 493–709 µs and removal took 284–326 µs. Successful 7z preflight took 529–736 µs and removal took 286–335 µs. The earlier regular-file protection run `20261127T194823Z-73460-19917ccd9d6a4cbbbbc7e4cd3726d030` supplies the matched 16-entry, 23-byte retention control. ZIP median worker time was 4,885 µs on that baseline and 4,171 µs on the candidate, with six samples each. The 7z medians were 6,358 and 6,192 µs, with three samples each. The positive directory cleanup timings are an initial measurement for the new route.

The same final receipt passed the non-activating `FileOps_ReviewRemediationContracts` selection in `20261127T222454Z-83668-9de769048e7c4ce19957eb2857e8b41a`. Setup, the 87,703 ms case and cleanup all passed with zero failures or skips. Native result SHA-256 `48b69ebfd938d072f8176438393ae2e35f5de937c335d44c6b49984cf77fa1cf`. Trace SHA-256 `cbba4ad401be434bedc94039ac1a76796dfbde1aa218ecb7f2866f8cabb82cee`. The filtered repository verdict is `NOT_EVALUATED`.

PR #384 merged at `b632fbf29c318afc530755c1965a967f634863ff`. The required format, tooling, CI gate, dependency review and CodeQL checks passed on the tested head. Main contains the verified code, tests, specifications and user guide. [Merged Pack directory cleanup](https://github.com/RedSalamanders/RedSalamander/pull/384)

### Follow-up live provider checks

These non-activating x64 Release selections byte-verified receipt `b98fe2d2848ccd311bf9fc1421237b686bbb74ad22750a3997870eeee9119364` on `1b2dd5db66813c97982fc3616e379617864aa48f`. They used the already registered RubyGloom fixture and unchanged machine resource manifest. Every selection closed its contained process and kept the filtered repository verdict `NOT_EVALUATED`.

| Case | Outcome | Run ID | Native result SHA-256 |
|---|---|---|---|
| RubyGloom Samba corroborated Rename | 3 passed including setup and cleanup | `20261127T221059Z-6780-fa54632c44a242f7ac55c04d0e98c058` | `05377e427c5182eb620471154d74d71b28830191ef15aab077d04ac65188bd0a` |
| RubyGloom Samba mixed Local Native and SMB Managed Move | 3 passed including setup and cleanup | `20261127T221221Z-79620-a29c0abe7850425b89387d01ae389309` | `e7b364d086a9d75945c73dabcbd14a9fef98b5727a670a66ed1042d7463a1a47` |
| SFTP F2 Rename | 2 passed, provider case skipped because the isolated runner could not find the connection profile | `20261127T221304Z-68276-a5461ad816ba475e909c20602aa5e45d` | `0763046c3b5c0705912b34c0e1ccf9e2b94981171720501272ec503f90048fb0` |
| OneDrive Personal F2 Rename | 2 passed, provider case skipped because the isolated runner could not find the connection profile | `20261127T221347Z-21212-e0ce8e53e145417396a318d448a83df0` | `a29f63509b1bdfaaa3b9f29470b8c8046f9e0f2d6390b4b43791326554275d75` |

The SFTP and OneDrive skips preserve the earlier successful qualification on dedicated profile copies. The current registry needed no changes for these selections. [Earlier live qualification](../2026-11-26_221027_live_mixed_smb/README.md#existing-live-route-rechecks)

Both Windows SMB checks ran through the existing approved SINON fixture on the final candidate `af79501c8d7b0fd1c7ecc469f0e59316e74642a2`. The canonical runner verified receipt `96a4e3da9f8bb0147d3eec1f70a0c1668003a02639bbcbadcfda342e3c90e35b`. Each check passed setup, the live operation and cleanup without activation. The temporary fixture registration ended with the original C resource manifest restored byte for byte. Its SHA-256 is `1779448ad25c15b19c41f1cd0a1c81d68c7711cfe834df2b9ee7599a3bbd7a22`.

| Case | Body duration | Run ID | Native result SHA-256 |
|---|---:|---|---|
| Windows SMB exact Rename | 1,297 ms | `20261127T222303Z-58080-cccbd2e315364d7998b1ddd5730912f3` | `a03c7e9136b841a6d525bbccd55f8ecfec31686eb6a0277608bb7443b0c3d6cb` |
| Windows SMB mixed Local Native and SMB Managed Move | 110 ms | `20261127T222342Z-69648-9a70cbf8427746e98ea52ee5792cb0b3` | `bf6d38f945058f66ea88fd23391648bdb7c170795227eb248bc0bb2575bd22ab` |

Both Windows SMB runs retain the filtered repository verdict `NOT_EVALUATED`.

### Follow-up checks on the Pack build

Four canonical non-activating File Operations selections ran on x64 Release source `af79501c8d7b0fd1c7ecc469f0e59316e74642a2`. The runner byte-verified receipt `96a4e3da9f8bb0147d3eec1f70a0c1668003a02639bbcbadcfda342e3c90e35b` for each run. All passed setup, the selected case and cleanup, for 12 passes and zero failures or skips. The filtered repository verdict is `NOT_EVALUATED` for each selection.

| Case | Body duration | Run ID | Native result SHA-256 |
|---|---:|---|---|
| Identity-less Delete and Pack writer interlock | 250 ms | `20261127T225237Z-23716-884c81862d934f8d96c5a73cac489228` | `0031e67c39ef50205e18a433d2243b5ab87429c35597ca9b3c307de4d6d65540` |
| Deferred exit rejects new work | 19,828 ms | `20261127T225319Z-16148-654e72083b4b4196bf774d3492063b00` | `c3dc62bd55d45e1cfcce7d91b5782f3aa81e10da6557d7beda8677ad4eca90ac` |
| Copy-only Move retains sources without destination proof | 313 ms | `20261127T225423Z-38824-899ab5092a25449cadcdcf376e8b6967` | `b31de1df6cba3fc030b1f66c53f84ed6014a67ef1f5dc165adb5ea91cbe086b3` |
| Provider capability matrix | 2,156 ms | `20261127T225507Z-21448-da5661c651a14e0fbe3f677d4ee104f6` | `c6f3bfeb0700bd14e01e737ef9b3f9b90ae24824539cdba2e9388f16855598f2` |

### Approved clipboard, consent and File Operations desktop pass

The approved 12-minute window began at 23:56:24 CET on 27 November 2026. All four canonical runners verified the x64 Release Pack receipt and used the interactive desktop lease. Their contained processes drained at 23:59:30 CET, after about 185 seconds. Launcher archive `20261127T225624Z-73372-fbb5a7b2ca9d4883a6e6f277115d4b5e`, results SHA-256 `858fd811a4334bb60e8e06b77017d3e530df2bf16ef6a5522cae5e56dbab9d59`.

| Case | Body duration | Run ID | Native result SHA-256 |
|---|---:|---|---|
| Clipboard Move admission and retained actions | 110 ms | `20261127T225626Z-84476-385832b37f9347a981ca3487f9f1e3c8` | `53f836d46bf6c16f9deb196ce8f92ee98a404f6bd2627521b44b29d9d6d8bbe9` |
| Deferred consent and Recycle escalation | 2,219 ms | `20261127T225708Z-81900-10a2f2ae880447188b63e22070011c1e` | `c61bf363d2c9b183b7cb1f87bfd529fa61e712a6f1d5fd70b65b60768ab6673a` |
| File Operations footer menu actions | 1,581 ms | `20261127T225754Z-51968-ea7c6e600c4540479618a458ae2452ac` | `ff074b3cb5fbd6e32fe8c60a77f0ad27fcc1a6ced7317202d1696ae6beca09e9` |
| Retained source and mixed-result controls | 2,158 ms | `20261127T225842Z-77184-5b1a4df6e31b416787f56c167187638e` | `2e69e6c37084fc3e6bcac868b8180d3af5984af42bfb60e8a8592491b4dbce73` |

The two File Operations runs each passed setup, case and cleanup. The two Commands cases each passed. The total was eight passes with zero failures or skips. Each filtered run retains repository verdict `NOT_EVALUATED`.

### Earlier ARM64 Commands and desktop trace

The retained ARM64 evidence in run `46027552244` on `bcc3f42e2e114953556605f3ef54e30919ada1bc` includes the original Commands result. It records 963 passes, two failures and five skips. The Direct2D failure-surface case reported that disabling Cancel all did not leave keyboard focus on Close. The mouse-opened main menu case had observed a readable popup, but its HWND was already invalid when the driver tried to post Down. Native result SHA-256 `f5e2a36e830560c8cfcda57419f2d56138544daf0c8dd3bbd7ba178d05dfc10f`. Native trace SHA-256 `08afdd4b28b6b3e26e4cae53de4913f418f9ac53695d5020e7025f300c79a43f`. [Retained ARM64 evidence](https://github.com/RedSalamanders/RedSalamander/actions/runs/46027552244/artifacts/15231407555)

The captured image digest for the separate retirement refusal matches the SHA-256 of the `C:\Windows\System32\wsl.exe` path encoded as UTF-16LE. PID 884 was created at 21:00:27.523 UTC. Its recorded parent PID 9292 could not be opened at retirement and returned Win32 error 87. The exact test-job relationship was unavailable. Startup preparation had finished at 20:13:28 UTC after retiring the approved account and search surfaces. Launch attribution remains open. Local evidence capture `20261127T230013Z-65064-7e827c14d0f7447b96441adb2053920b`. [ARM64 job log](https://github.com/RedSalamanders/RedSalamander/actions/runs/46027552244/job/140671268108)

### Local replay of the two historical ARM64 Commands cases

The approved x64 Release replay used the same verified Pack receipt. The two exact cases passed all three repetitions each with zero failures or skips. The desktop interval started at 00:07:11 CET on 28 November 2026 and ended at 00:08:51 CET. Launcher archive `20261127T230711Z-59680-b05a8ca3256646ac9d1a3857e088b871`, results SHA-256 `a1e2057ca323e6f3c30b4a4515d6f4761cc5a994cc6bd6f36b27dcf49a6b7bc1`.

| Case | Durations | Run ID | Native result SHA-256 |
|---|---|---|---|
| Fallback Cancel all focus | 4,548, 4,593 and 4,539 ms | `20261127T230713Z-75860-e502bf5951f54d6d9f9c5ac358bf7110` | `314b60ed8b0057a23979bbcd634310d06676d2bd6475fee64382dc9d1e4dbc1a` |
| Mouse-opened main menu keyboard input | 461, 472 and 436 ms | `20261127T230810Z-43548-cfb4c4e19f4649ad99889572084627ad` | `248902ab90aba90163b25c7c2c4be511bbda426641070928a7e09458ed220363` |

The focused repository verdict is `NOT_EVALUATED` in both cases. The native ARM64 follow-up remains open.

### Hosted x64 follow-up after the Commands trace fix

The x64 job on `651a700fc66de7927ad6684b08a8abc1553e1afa` finished with 3,045 passes, two failures and 71 skips. Commands passed 965 with five skips. File Operations passed 294 with 36 skips. All four interactive viewer entries passed. Tooling failed both attempts at the HWND render-target source assertion. Viewer PE non-interactive failed its first WebView2 private-origin observation and passed its retry. The hosted classifier retained one regression and one blocking flaky result. [Hosted x64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46061664377/job/140741316649)

The original x64 job log is preserved in `20261127T231756Z-73420-0ed1a0cd3ac34ceca26562bef5678857`. SHA-256 `c27195e3d584326b434566e7283f4c9d77f7f16b6a232f24561dd3bac0e2f5d2`.

The HWND assertion expected `wil::com_ptr` for the shared factory, target and brush. Their existing `wil::com_ptr_nothrow` ownership predates both hosted revisions. The focused Pester 3.4.0 baseline reproduced the failure in `20261127T232000Z-82176-ac3ad0d4234540b3a9e6747fe4338bf2` with 2/1/0, SHA-256 `d11334610fb945f99122a4ea5148fffb89acb2ddbdd48d0351c25f2d5fe47dd4`. Candidate `fe5d6e248bade5b7e35058a035baa6d0f88820a2` passed all three checks in `20261127T232045Z-31404-2cf21bf5d28044b8ba268666f6aa853c`, SHA-256 `17c66fb622cde05978307d59acf05fce459ee66d666b25e50b5dc805d95480cd`. [Source-contract repair PR #389](https://github.com/RedSalamanders/RedSalamander/pull/389)

PR #389 merged at `db658051844f10de7c07452bf54b9b160d8dade9` after the required checks passed. Its source-contract test and specification match the reviewed head on main.

### ViewerWeb private origin follow-up

The original non-interactive Viewer PE console confirms that only its raw HTML source assertion failed within the private-origin child. The isolated profile, script disablement, raw byte checks, and local and virtual PDF checks passed. The original console SHA-256 is `5f7e9b59b9f2f3ae781c762643978b60d1ce627cae82bb5d0ff19612d87b73a8`. The full x64 artifact is preserved locally in `20261127T232122Z-37168-94cfd42c925646baa43dc150962a6e9c`. [Hosted x64 artifact](https://github.com/RedSalamanders/RedSalamander/actions/runs/46061664377/artifacts/15237492399)

The pre-change x64 Release non-activating Viewer PE entry passed on `fe5d6e248bade5b7e35058a035baa6d0f88820a2`. This contains the HWND contract correction and unchanged viewer source. Receipt `8c2cff1733f7b30d221cb8695e191eb070a9d3b7eb038acc62abec61e7d6a2f1`, run `20261127T232334Z-39372-07dad03628ad4bf997b216bca421f17a`, console SHA-256 `4cac7e242a5992be3ddc3e330c0713dd330b5b84abf5d5f647df88303fe9ea21`. The filtered repository verdict was `NOT_EVALUATED`.

The first two exact ViewerWeb candidate launches stopped before compilation because the isolated checkout lacked its pinned vcpkg inputs. Their run IDs are `20261127T232709Z-59132-39712ba67d284d32a9c6e4c3a0af64a4` and `20261127T233003Z-67196-2167e0c047014d75bf7f0def457c04fd`.

The prepared checkout built candidate `f8e90684f5099d93491dc47ea4625cd715898d29` with zero warnings and errors. The x64 Release test-enabled solution receipt was `2e9b6e0b31e2fe44ef0ae8fcdd748b0cc8a63ba63ea716e59971621882ccf6b1`. The canonical non-activating Viewer PE selection `20261127T233558Z-69264-01d67640cf0a43ce8cd65789f8cd9ddf` passed with zero failures or skips. Both controller startup and the exact browser source `https://viewer.redsalamander.invalid/web/1.html` were observed. Navigation completed successfully. The console SHA-256 is `a5d767efac6775a188a220dc88de3063b386162cbf8a1715814f78d4f7f49669`. Repository verdict `NOT_EVALUATED`.

PR #390 merged at `ae26292aee7c36f2a40e4dcce9b589d89426b085` after the required checks passed. The original intermittent hosted symptom did not reproduce on either local run. [ViewerWeb repair PR #390](https://github.com/RedSalamanders/RedSalamander/pull/390)

### Hosted ARM64 follow-up after the Commands trace fix

The ARM64 Release test-enabled build succeeded on `651a700fc66de7927ad6684b08a8abc1553e1afa`. The native regression step finished failed. The tooling suite reproduced the HWND contract failure in both attempts. The non-interactive Viewer PE entry promoted. [Hosted ARM64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46061664377/job/140741316586)

Commands originally reported 963 passes, two failures and five skips. The failures were `cmd_pane_fileops_conflict_pending_inspection_teardown` and `cmd_viewer_space_hover_does_not_rebuild_static_cache`. The hover witness recorded zero tooltip paints and `ConsoleWindowClass` at the pointer instead of its viewer. The conflict-inspection case passed in the single-case replay. Original and replay result SHA-256 digests are `56bdbbf8c2fbe929ae39cc16ab7e31cdb886b484c228efcea179f7cd90ac5a51` and `b7e6abedce776c70eac07e922d21e5f068e2ac1ea664bd09bae26217b2d35361`.

At replay retirement the Default desktop boundary refused restoration over a foreign foreground. The captured image digest `95a9bdfa9774cbec665bce075b37ff65cb5cf6f70e0286eabad32eb453fdfbd1` matches the exact UTF-16LE path of System32 `wsl.exe`. PID 10916 was created at 23:42:28.4513297Z. Its reported parent PID 8212 could not be opened (Win32 error 87). The stock preparation snapshots showed no WSL console before either boundary. The exact arguments and launch ownership remain unproved. The hover replay, File Operations and all four interactive viewer entries stayed pending.

Job log and the downloaded native evidence use local run `20261127T235131Z-79960-2fde120a957943a4a7f6fdf4bd17fcd8`. Raw job log SHA-256 `0768232a9690ccb397a17d2fbf0580ca7208b09fff29f50bcb7461cf60bfee14`. The direct Pack run and the two intervening Pack descendants finished with native regression skipped at admission after main advanced. [Pack merge admission](https://github.com/RedSalamanders/RedSalamander/actions/runs/46076994462) The next Pack-bearing source, `ae26292aee7c36f2a40e4dcce9b589d89426b085`, was admitted and both native regression jobs started. [Current x64 and ARM64 qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/46085203828)

### Controlled mixed Recycle and metadata recheck

The exact x64 Release solution tree at `f8e90684f5099d93491dc47ea4625cd715898d29` matches the tree merged by PR #390. The canonical non-activating `FileOps_ReviewRemediationContracts` run `20261128T000055Z-70540-431839e6d3ef4e6b91580433544b69aa` byte-verified receipt `2e9b6e0b31e2fe44ef0ae8fcdd748b0cc8a63ba63ea716e59971621882ccf6b1`. Setup, the case and cleanup all passed with zero failures or skips. The case ran for 90,703 ms and the filtered repository verdict remained `NOT_EVALUATED`.

The case required both exact paths in a controlled Local mixed Recycle selection. Both files survived until the consent was accepted. Approval removed one through Recycle and one through the permanent route. A fresh Cancel selection retained both. The same case passed the late metadata-loss result, staged Move Skip and combined pre-content consent checks. Result SHA-256 `b4ffc55b31d594136c986ab78d1749ba981aa4ddf624987f376627af7e061ac8`. Trace SHA-256 `216510d19fa8f40c139be52bed0a6ce45152e049ab453760c53d934a1e25d9b8`.

### Upstream ARM64 runner image evidence

The GitHub runner-images tracker records the same recurring WSL prompt on the native Windows 11 ARM64 image. Another project's process-start capture traced its prompt to the hosted compute agent. The image maintainers' open PR proposes preinstalling the native WSL package. This is a plausible explanation for the matching WSL image and unavailable parent on this run. The RedSalamander artifacts still lack the arguments and parent launch chain required to establish attribution for this instance. [Runner image issue](https://github.com/actions/runner-images/issues/14264) [Upstream image PR](https://github.com/actions/runner-images/pull/15362)

### Commands evidence from the local Pack Full run

The x64 Release Commands entry on `05252ae38332641222c55ba1fb11a9f1abf264ae` passed 965 cases with zero failures and five skips in 1,279,206 ms. The native result SHA-256 is `23f04ec980182e56abc702efdfbd7dd57a9d9919797726136e4e49fc69ca5e23` in run `20261127T231106Z-58756-cb5c3c7c52a1429aa066841fab1ae55b` at `artifacts/archive/4cb089111a23/Commands/2026-11-28_013004/commands_results.json`.

Both cases from the preceding ARM64 Commands run passed locally. Conflict-inspection teardown took 154 ms and ViewerSpace hover took 76,817 ms. The Pack and Unpack round trip, Find result shortcuts and file actions, Cut and Paste, and the host queue paste route also passed. These are local x64 results. The broader Full sequence continued to File Operations and the ARM64 gate stayed open.

### Pack cleanup source-contract follow-up

The canonical x64 Release Tooling Pester selection on `f8e90684f5099d93491dc47ea4625cd715898d29` had 1,518 passes, one failure and zero skips. The HWND factory, target and brush check passed. The failing Observatory Pack source guard expected the previous pane-only cleanup dispatch. Exact Pack directory cleanup now runs in the Pack worker, and regular files still use the receipt-bound Delete request. The focused gate rejected `tooling.pester`. Run `20261128T000631Z-58780-a5e96683c36d488982248a59600ab07d`. Legacy result SHA-256 `08ffa0cbf078eac797b026dab321c2ae7d4be2102cd5191b5b5c4f5080d6b558`.

Candidate `c092f3a60d8509bbc861eb68583b1f7e5600ac16` scopes the static checks to the Pack completion and worker paths. The source contract checks publication, instance and receipt binding before the regular-file Delete request. It also checks that tree removal is conditioned on published output, approval and successful strict rebind. The focused Pester 3.4.0 run `20261128T003626Z-62128-8af469d17c83462e9b9a2666f2926de5` passed all 244 cases with zero failures or skips. Result SHA-256 `96e9248c0249dc0d44672db2afabdb05018abc6bab3c93c3922d6a747a2c3738`. The exact specification inventory and changelog checks passed. A first build-free Tooling launch stopped before execution because its receipt belonged to the earlier revision. Interrupted run `20261128T003802Z-80852-d75ce95db7e447e2bd7f195277b1182b`. The exact x64 Release solution build passed with zero warnings and errors in 7 minutes 31 seconds. Receipt `40b6b04330edf92a2ceb6034311fcfe788fca998d538f8e37028136c6109c0b3`. Its fresh Tooling selection `20261128T003926Z-53672-4f18d26075944d2fbb5fb76efb94521a` promoted `tooling.pester` with all 1,519 cases passed, zero failures and zero skips. The selected entry took 1,361.32 seconds. The filtered repository verdict is `NOT_EVALUATED`. Results SHA-256 `30bed50b012d52b0f6c3ced2d50a654d3fc5ad95962bdabae37c1030a05d4756` and promotion evidence digest `4bee2e3e8b425317e36d63831b0933c513bbf28c1dc03db6177e089a095f454d`. PR #395 merged at `73c70e33037273b3727835d9b61a265883f15470` after the required checks passed. Hosted run `46085203828` is still on the source that reproduced the old Pack guard locally. Run `46089888885` on the merged repair is queued. [Hosted follow-up](https://github.com/RedSalamanders/RedSalamander/actions/runs/46089888885) [Repair PR](https://github.com/RedSalamanders/RedSalamander/pull/395)

### Non-activating rechecks on the merged Pack repair tree

The reviewed source `c092f3a60d8509bbc861eb68583b1f7e5600ac16` has the same tree as merge `73c70e33037273b3727835d9b61a265883f15470`. Both selections verified x64 Release receipt `40b6b04330edf92a2ceb6034311fcfe788fca998d538f8e37028136c6109c0b3` and retained repository verdict `NOT_EVALUATED`.

- The Pack and Unpack Commands case passed three repetitions at 2,632, 2,549 and 2,555 ms. Run `20261128T011207Z-80416-2cffd106732d4d30b7b8f62c29b8e3df`. Native results SHA-256 `22717aca1b5d9ae3f9758981772c4a76a96bb26a62d8a613782a6e9435882896` and archive metrics SHA-256 `4b0571438442ffbdf0fcb22329bf5de4be8bb7448a19f23ead96103a55063081`.
- The non-interactive Viewer PE entry promoted with one pass, zero failures and zero skips. WebView2 created its controller and completed navigation at the exact expected private URL. Run `20261128T011325Z-38820-242b04608cfc41d9bd6dd4694073e27d`, console SHA-256 `c2cb263adf14c7344929bb4371573d5e09f5c566605a6fc2b1e81ab0004fde30`.

### Local Pack Full terminal result

The fresh x64 Release Full run on `05252ae38332641222c55ba1fb11a9f1abf264ae` finished with 3,091 passes, two failures and 58 skips. Both failures were in Tooling Pester. The HWND ownership assertion and the original Pack finish-layout assertion were corrected in PRs #389 and #395 and passed the later exact-build Tooling entry. The older Full verdict remains `FAILED`. Its final artifact receipt was `c4d1306bb2cbfd63c3ccc4a4783a5d58f50552bfcc899bd4417df2d3b4a7873f`, following its initial receipt `7288592462615c8e6b2b35ca6b172faae671e88d108e0fad517ada685a855bdb`.

File Operations passed 306 cases with zero failures and 24 skips in 2,663,900 ms. The independent C and Z volume case, guarded SMB Recycle, SMB mixed Move, deferred consent, clipboard admission and remediation matrix passed. The isolated SFTP and OneDrive Personal Rename cases skipped because the Full profile lacked their registered connections. File Operations results SHA-256 `0bf96e6dca676b1223303e52f70788238354e3b522dbe89f8359aebb18102ab9`.

Commands, File Operations and Viewer PE Interactive passed. The Release SQLite interactive entry exited successfully and reported zero compiled interactive cases. The previously recorded x64 Debug SQLite Tab result remains the interaction witness. Compare Directories passed 255 with 29 skips. Seventeen entries promoted and Tooling remained incomplete.

The contained runner drained at 02:21:38 CET after roughly 130 minutes. The original C machine-resource manifest was restored and independently rehashed to `1779448ad25c15b19c41f1cd0a1c81d68c7711cfe834df2b9ee7599a3bbd7a22`. The inner archive `20261127T231106Z-58756-cb5c3c7c52a1429aa066841fab1ae55b` stores the canonical result with SHA-256 `c13fd5dfdde6f86756b4e6d0cacc7c4557b6b2d5a7d7632aea513e736acac33f` and the detailed result with SHA-256 `cedb364430091a423b41ddd60a3ebd138afe09d2a5de637de28224e6baf9d5cb`. The launcher archive `20261127T231104Z-35576-a767e8721c16434087f1f522b7f6b3dc` stores the original console with SHA-256 `9e89c1d99a7a334b1c5ad3f607506ccf626312c45006eb6201b6c1380236f3df`.

### Focused replay on the merged Pack repair tree

The replay on the exact x64 Release tree `c092f3a60d8509bbc861eb68583b1f7e5600ac16` used receipt `40b6b04330edf92a2ceb6034311fcfe788fca998d538f8e37028136c6109c0b3`. The desktop interval ran from 02:23:20 to 02:28:51 CET. Both Commands selections passed all three repetitions with zero failures and skips, and retained the scoped repository verdict `NOT_EVALUATED`. The native ARM64 gate remains open.

| Case | Durations | Run ID | Native result SHA-256 |
|---|---|---|---|
| Late conflict-inspection teardown | 154, 140 and 139 ms | `20261128T012321Z-63660-4c03b154a87243a69dee3b28932daeac` | `4548dc92b00cb2e900430fd5dacd48cb1bed12e2be6337581f744df565b81be3` |
| ViewerSpace hover and static-cache reuse | 74,356, 75,031 and 74,961 ms | `20261128T012415Z-59792-36856b46fff94bd68860feecdf030885` | `d152ab9c1b8a3bcd6a718837e84be37f4fbf7dd878e48c00eed5a7be9cea4660` |

Launcher archive `20261128T012320Z-83620-ac6af69b045d4678887c892de1ad796a`, result SHA-256 `98534fdb8f741ba9e0c31e8e31314391060f8eeebb13922931e79cc8a208be67`. A fresh corrected Full run started after the replay drained. The native runner is `20261128T013006Z-46208-de4321aec9534bfabbec57a5e7d4e641` and the timed launcher is `20261128T013004Z-67052-fa4ead283b86415b99ac7c9578fd104e`. Its build re-attestation passed with zero warnings and errors, receipt `f77916da01c53b838f912968a3ba82bb0138ec24e07fac9e1edc60cf8b989c1a`. All 1,519 Tooling cases passed during Full. Commands began at 03:18 CET and passed 965 cases with five skips and zero failures. Native Commands result SHA-256 `370232025248e1b4bd13842fdbb94da3bb99ea3db7f3d0dfaab1e5702aaa20dd`. File Operations started after Commands retired.

### Hosted x64 result on the ViewerWeb follow-up tree

The hosted x64 Release regression job on `ae26292aee7c36f2a40e4dcce9b589d89426b085` finished with 3,045 passes, one failure and 72 skips. The HWND ownership assertion passed. The stale Pack source guard failed in both attempts and was the only blocking regression. The ViewerWeb witness observed the exact private URL, and the non-interactive Viewer PE entry passed. Commands passed 965 with five skips and File Operations passed 293 with 37 skips. Viewer PE Interactive, the Release SQLite entry, and ViewerText Find and Go To all exited successfully. [Hosted x64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46085203828/job/140793216069)

Native evidence was downloaded as `20261128T023227Z-21272-1448c113c8f442f2aa441f8207ef10d9`. Raw job log SHA-256 `a3063ba7f955ca008e124d7cd6346ccf3d97d478b6b13f9a5373701d20274447`. The queued follow-up contains the merged Pack source-contract repair and still needs both native verdicts. [Hosted x64 artifact](https://github.com/RedSalamanders/RedSalamander/actions/runs/46085203828/artifacts/15244621114) [Corrected native follow-up](https://github.com/RedSalamanders/RedSalamander/actions/runs/46089888885)

### Corrected x64 Release Full terminal result

Fresh Full on `c092f3a60d8509bbc861eb68583b1f7e5600ac16`, the tree merged by PR #395, passed 3,093 cases with zero failures and 58 skips. All 18 entries promoted and the repository verdict is `PASSED` for this x64 Release Full run. The artifact receipt at completion was `f77916da01c53b838f912968a3ba82bb0138ec24e07fac9e1edc60cf8b989c1a`.

Tooling passed all 1,519 checks. Commands passed 965 with five skips. File Operations passed 306 with 24 skips, including the independent C and Z volume case. Viewer PE passed in both activation classes. The Release SQLite interactive entry again reported zero compiled cases. The earlier approved x64 Debug run supplies the SQLite Tab interaction evidence.

The runner drained at 04:29:37 CET after 7,174 seconds. The C machine-resource manifest matched its original SHA-256 `1779448ad25c15b19c41f1cd0a1c81d68c7711cfe834df2b9ee7599a3bbd7a22` after restoration. Canonical result SHA-256 `13a2306094b12ed615eb253d0b41931b1142507435b777a756c752023c3318eb` and detailed result SHA-256 `d2d43134b85f431cca9c20925a561b2690c3f8caaaf5ada9dc0db19d2f809193` are retained in `20261128T013006Z-46208-de4321aec9534bfabbec57a5e7d4e641`. Launcher archive `20261128T013004Z-67052-fa4ead283b86415b99ac7c9578fd104e` has result SHA-256 `4da2c186e2f095d01108ca70c9ea49fd870e9cfffc3cd260e7d3e39f3e31baf6` and original console SHA-256 `d031d7a3e878a8b03960a192d75cf6e8e2a57b6afb32e6eb0d5a845761b3a117`.

### Hosted ARM64 timeout on the ViewerWeb follow-up tree

The ARM64 Release build succeeded on `ae26292aee7c36f2a40e4dcce9b589d89426b085`. The native regression step ended when the owned PowerShell invocation reached its 18,000-second limit. The preserved run state for `20261128T003926Z-3616-b491431d3e6745b097cd8ab83094c6eb` has promoted non-interactive Viewer PE, failed Tooling and Commands, and File Operations still marked running. All four later viewer desktop entries remained pending. Run-state SHA-256 `ff1bd514c4c916c147818cb1438affaa5ea3ff7a382ac7e6f3592c94b1235460`. [ARM64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46085203828/job/140793216128)

The original Commands entry reported 964 passes, one failure and five skips. The Completed group case captured `paintGeneration=0` for the re-expanded success card while layout, auto-resize and group motion were active. It observed zero published actions and failed the Dismiss assertion. The isolated replay passed. Three shuffle attempts retained three, six and seven failures. These included interrupted Move presentation, ViewerSpace hover over a console and several Terminal creation failures with `0x800700AA`. The original Commands result SHA-256 is `9eb0df3134740492b829339443758807fcc031fb1263d420686c4705daa1cc97`.

The original File Operations result reported 292 passes, one failure and 37 skips. `FileOps_PopupStableHostedTree` retained the routed press, avoided full tree rebuilds and reported a change in native focus or foreground ownership. The trace gives no identity for the new foreground. Results SHA-256 `585d84017532651b8cfca440084d50e39bfe41b9b3974f2765521cac2adab40d` and trace SHA-256 `c996984deb776bb8127434f542b570d3f0f6110bf10c0a48784d583a75af16e2`. The retained viewer dump is from the expected crash-capture negative control, which passed. It does not account for the timeout.

The raw log and native artifact were archived locally in `20261128T054209Z-76168-dbb2940b26694021b1dfcb50c9120eff`. Raw job log SHA-256 `e63f2892466989668b52b538a79181627b299de16812071b261df81e6a2ace67`. The corrected Pack-bearing follow-up on `73c70e33037273b3727835d9b61a265883f15470` was admitted and has started both native jobs. [ARM64 artifact](https://github.com/RedSalamanders/RedSalamander/actions/runs/46085203828/artifacts/15251080129) [Corrected native qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/46089888885)

### Focused replay after the ARM64 timeout

The corrected Pack tree `c092f3a60d8509bbc861eb68583b1f7e5600ac16` used its re-attested x64 Release receipt `f77916da01c53b838f912968a3ba82bb0138ec24e07fac9e1edc60cf8b989c1a`. Three focused selections passed all nine repetitions with zero failures and skips. Each filtered runner retained the repository verdict `NOT_EVALUATED`. The desktop window ran from 06:47:26 to 06:50:07 CET and the contained launcher drained.

| Native case | Durations | Run ID | Native result SHA-256 |
|---|---|---|---|
| `cmd_pane_fileops_completed_group_and_navigation` | 1,073, 937 and 951 ms | `20261128T054727Z-82264-8fb75abf50a44a4083e603896f2c53f7` | `1aa74a9805b386313b5b083c7aaffe96a80f18a0d08045705e34a3c67c05e59d` |
| `FileOps_PopupStableHostedTree` | 1,735, 1,734 and 1,719 ms | `20261128T054822Z-77756-028207fc8fa94ae5b1410d1f3ba2410f` | `a35630cdc0630b991bbfa6d8ba7b5983236b3f958623135268a9401e9d7c4ac0` |
| `settings_terminal_plugin_roundtrip` | 99, 95 and 94 ms | `20261128T054916Z-67720-6a4f172136154f0a830b48bde5efe3f3` | `4a76c606a351e82ebde5b47dbc5cd38cc8a7d8498d82d641c12f02a7a9239ea9` |

The Z sandbox retained the launcher archive `20261128T054726Z-73492-130671ecea544ba2bec8888b10f2c058`. Its results digest is `373cd6fad8e87442db01e2f6ff7688bd061196b85875750339e6a09317b7b733`. These isolated x64 passes leave the hosted ARM64 focus and Terminal outcomes open.

### Completed-card paint and queue desktop checks

PR #398 merged at `02755d02bde37f7e8050a366d70db45cb75adefa`. It uses the expanded card's painted body to decide when to check the original Dismiss and action-count assertions. The initial build compiled with zero warnings and errors. The source barrier rejected its receipt after the formatter edited whitespace during compilation. The clean formatted source `70a2593a130690d130465e37e9a990d527b99eb3` then built x64 Release with zero warnings and errors, receipt `a8b9a8988348b3b7b411ec961b951a91a63ac985859a86b6776f39b37f4dacb2`. [Merged PR](https://github.com/RedSalamanders/RedSalamander/pull/398)

| Selected case | Native totals | Case durations (ms) | Run ID | Native result SHA-256 |
|---|---:|---:|---|---|
| `cmd_pane_fileops_completed_group_and_navigation` | 3 passed | 1,012, 983 and 1,058 | `20261128T061526Z-56756-f553354d8c524dbc85b38fec1ba51620` | `585a50330b65842c11cc1e00039e6d01cbdae238ffaf47ecbbc7436a71eb08cb` |
| `Phase5_DiscoverySingleTraversal` | 3 passed | 141 | `20261128T061621Z-21020-1f2bd5967c20490b93efdaa69a0e9f06` | `20af940d922f2517a76efa77327984184212765932dcd97d5ef75523c97663ff` |
| `Phase6_PopupSmokeResizeAndPause` | 3 passed | 20,203 | `20261128T061715Z-60188-1c7eac83e6c849cdbac6c541305aba9f` | `990b1548841a9dc58a58cbaae4260582ede65a0fc845d924b3913066ba3b9b3d` |

All three selections had zero failures and skips, with filtered repository verdict `NOT_EVALUATED`. The File Operations totals include setup and cleanup. The Start now control released the selected task's scheduling gate and kept the competing mutation interlock. Pause all and Resume all were invoked through the popup on an active transfer. The launcher ran from 07:15:24 to 07:18:31 CET and drained. Z archive `20261128T061524Z-71128-50d56e814b354eb9b02eb600842ac867`, launcher result SHA-256 `698aa31fa94e5d68d308fb5324f1e79ee1433fac3e503b95175d1c0ba7ec33a1`.

### Virtual-pane Move folder picker

The exact `15ab5b4d7771c226cd5d9d0159be42a8c98f9305` x64 Release solution built with zero warnings and errors. The first compile found one redundant lambda capture. After correction, a selected-project build passed, and the runner refused its receipt before running a case. The accepted full-solution receipt is `246fb3e053b0c5108b3b939da54205e5b3f83fead20625b0139191c859e76e1a`.

Run `20261128T063212Z-82892-0030aa0230364b128307b599cad7e729` passed `FileOps_ReviewRemediationContracts` and its setup and cleanup with zero failures and skips. The private Dummy pane sent the production item-menu command and used a known Windows folder in place of the Shell picker. Cancel left the source unchanged and created no task or destination file. Accepting the one Copy-only question yielded one completion, matching Local and source bytes, and a published/retained receipt. The runner stayed non-activating, and the virtual source fixture retired. Filtered repository verdict `NOT_EVALUATED`.

Native results SHA-256 `ac4baa8e6a6dca26cfaf287d2ad46888856629e7013156b1023216c49f15f05e`. Native trace SHA-256 `bd23659369b928ffb13619113ee711d7b0da7a610c277eb051a7667fc157ee87`. PR #399 merged at `2592e621ce22925270d0487235e85f12780fcba7`. The corrected Full run also passed `Commands_FileOpsFolderPickerMoveUsesHostQueue` and `folderView_drop_integrity_guards` for Local and same-folder Ctrl+drag. [Merged PR](https://github.com/RedSalamanders/RedSalamander/pull/399)

### Idle Queue Start default

The test-only x64 Release baseline `4c3021081326856b89a89b06a2d2404cb645ce0a` compiled with zero warnings and errors, receipt `a4d864467252afa5fa174b402877b6a44d4e92e5ae249692c8352a70a94fd100`. Focused run `20261128T063633Z-64448-9ad2f46af98e4e8a998d334426ed63ed` reported zero passes, one failure and zero skips. The actual Copy with Options dialog was opened with the Queue mode selected and no active task. Cycling Start did not produce the expected Parallel override. The case failed before reaching the second accepted Move submission. Native result SHA-256 `718a494ac54995fffb9a430596bee44a8ad21d77c378f5fdd2551ffc098c2125`.

The desktop launcher `20261128T063632Z-10556-e3b21274aedc412f9aec1255118fec1b` drained at 07:37:24 CET. Its result SHA-256 is `7a2782b0af0222b1eb48fc8ffb3c0b662b4317a4bff26e5740b05ca2e7dc214b`. The pane commands initialized Start by asking whether the new task would wait. PR #401 initializes Copy and Move with Options from the current mode even when idle. The corrected repeat is recorded below. [Start default PR](https://github.com/RedSalamanders/RedSalamander/pull/401)

Corrected `930a157a0a603d02041620abac233d15e4e43d4f` built the complete test-enabled x64 Release solution with zero warnings and errors. Receipt `9de49d5aa379df4729adb8a7e152adf882fc24e3a6451bf660ccb7513122b260`. The same Commands selection through the actual Copy and Move with Options dialogs passed all three repetitions at 1,037, 967 and 985 ms, with zero failures or skips. Both submissions started from the idle Queue default, used the selected Parallel override in their immutable plans, published matching destination bytes and kept the current and saved mode. Filtered repository verdict `NOT_EVALUATED`.

Run `20261128T064719Z-69280-0f5e956c2334493d98bb27f3c123ab4f`, native result SHA-256 `5ca721fc48a2e09b6d4183d7c3b5478734e397b8c2b94cbee94e67bf2045c0f7`. The contained desktop launcher drained at 07:48:14 CET. Z launcher archive `20261128T064718Z-84100-a65b9bd284bc411684a69a851d73ca6a`, result SHA-256 `735c47eba846645669e6259a9d1e3fca992ac7f5ecf835cbc13cd95a49f6b68e`. PR #401 merged at `7ef16d9bcbe89160fe905bd65972ca0b8bc2a10f`. [Merged PR](https://github.com/RedSalamanders/RedSalamander/pull/401)

### Hosted x64 result on the corrected Pack tree

Hosted run `46089888885` admitted main `73c70e33037273b3727835d9b61a265883f15470`. Its x64 job `140881818899` finished at 08:54:18 CET. Native run `20261128T062315Z-8740-20f3bc9edfbf4f3482ec337e4ed15418` used receipt `20e7bb444ed7a8f7ba27f2d6a63e0ee0ec90951125fe283dbb332076056b8bf8` and reported 3,046 passes, one failure and 71 skips. Tooling passed 1,519 checks including the corrected Pack guard. Commands passed 965 with five skips. The archive command case passed in 3,447 ms. File Operations passed 294 with 36 skips. [Hosted x64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46089888885/job/140881818899)

Viewer PE Interactive was the only failed entry. In the long-running Shell viewer sequence, the ViewerSpace Escape child exited successfully and left its job quiet. The harness observed no foreground window and a changed last-input tick. The saved identity and allowed foreground predicates passed. Input ownership remained unproved, so the nested harness and outer churn parent refused restoration. The isolated entry retry passed. The canonical run classified the original failure as `FLAKY` and blocked qualification.

The raw job log SHA-256 is `2526e9e1000acd0ccd7788284f901d2fe74301ee898e63c8d33ecea74ff89de2`. The original run state SHA-256 is `4249d4f1c3d94c768689a42397f74c52c6b535c15de67ec9727febe7582074ce`. The raw log and GitHub artifact `15255235116` are retained in Z archive `20261128T075556Z-62008-b3591ad0716c42f09322ee4eacfa40a7`. ARM64 continued in the native regression step when the x64 result was captured.

The local x64 Release Viewer PE Interactive selection on `706790ac621f600b73d75c398a5fd2a5497ae3b3` verified receipt `f3cbf8b9b97d7f3da04e754b46d2c432a8a6c32f848ee96829986adb0d98fce9` and passed in 182.1 seconds. It recorded one passing entry, zero failures, zero skips and repository verdict `NOT_EVALUATED`. The hosted restoration refusal did not occur. Native run `20261128T080011Z-61776-291351387bc244639ac28eddb9138474`. Launcher `20261128T080010Z-80288-758d6ab972d14727a35e7bca40970c03` drained at 09:04:13 CET, result SHA-256 `7d37162c18807c95cab10a92a7386460684d26e96c5e739d4773d6fe1b2c23c5`, stdout SHA-256 `8f24d52c20cce43dcc3d1d99e3772f4cf4dab4244049f29b13d7e14ad51c1a1f`.

### x64 replay of the ARM64 Commands order

The x64 Release Commands replay used original ARM64 seed `253058945` and exact `930a157a0` with receipt `9de49d5aa379df4729adb8a7e152adf882fc24e3a6451bf660ccb7513122b260`. The native process exited 1 after recording 84 passes, zero completed failing cases and one skip. The runner treated the partial coverage and exit as one unclassified failure. The last begun case was `cmd_compare_directories_options_uses_dxui_labels_without_visible_legacy_statics`. The shared trace records a ValuePattern read that exceeded its client deadline and then ignored COM cancellation. The self-test deadline guard terminated its process. No dump is present in this archive. The later Terminal configuration assertion did not execute.

Native run `20261128T064920Z-56668-0ded44cb0c7e49779ade4e06b323dbe3`, results SHA-256 `0c337f16745a54bea21f7ed100120c8b16e0973b6a9db61bd716f5c61052ee86`, shared trace SHA-256 `e9cf2f0474840d7fa83ebf03f81f8e2665a4169a3c54e220ae95982eb8deb14b`. Launcher `20261128T064919Z-67704-3ffaef5a91a34d77b0dc2e739c3809e2` drained at 07:54:28 CET, result SHA-256 `801283ac6f496dbb7dab91fc310bbf5db79d31f36f0621b92c0b9ade2a37248e`. The exact isolated case reproduced the timeout.

Isolated run `20261128T065834Z-80040-30443817559c4a6eb0e39bfab6a32899` reproduced the same ValuePattern deadline and ignored COM cancellation. The original app stopped before publishing the selected case result. The later listing response is a static declaration. The runner reported the incomplete execution as failure. Shared trace SHA-256 `c55bf63bb970267319a2a3a67a55fad37bcd890cc2970639e732bfab183fba90`. The contained launcher `20261128T065833Z-83872-c92d2ee7eae843f09600c0533ed0b139` drained at 07:59:38 CET, result SHA-256 `922b3283926165cb60876876180978a6862c8119cbc1f1e545d27f7eb23685dc`. [Compare options correction](https://github.com/RedSalamanders/RedSalamander/pull/403)

The initial correction `28019b86e346776575c5e8be32793abb10e26521` built x64 Release with zero warnings and errors, receipt `1830a380efe341fca4c17bebc78e06a48af3813f5be5031b3182d88ccbe239d5`. The selected case reached the subsequent command-button name read, which timed out and ignored COM cancellation. The deadline guard again terminated the original app before it published a case result. Run `20261128T070952Z-70832-68da8b350c8e41f5a2661a7094f619f4`, shared trace SHA-256 `6554352d40cc9956bdfca5a4a7c174865d2fc08bc4a659ae57a26d32c28f4e93`. Launcher `20261128T070951Z-84912-09871119d5454165b91253f4c3b5bd8f` drained at 08:10:47 CET, result SHA-256 `56bafec213e4cbf3d250b9275ba4da3d2b2656515d3c594b15cb6acb95afa48f`. Candidate `f872029e6f9e6492fdb2e7b1f935181f379156bb` uses the owned host for both reads.

Exact `f872029e6f9e6492fdb2e7b1f935181f379156bb` built x64 Release with zero warnings and errors and receipt `d2f663a2d11f02891455dc4327706ca0c2a772b7cf91fa0602c081773221131d`. The baseline selected run `20261128T072015Z-85876-0838cde812f842ddaf1dab5559e8bbf1` passed three repetitions in 485, 424 and 396 ms, result SHA-256 `a15ec7624cf1803c97cb5a1240d5478c790ac24e1dfb9e9668c04b9bf959b406`. Repeated open and close passed in 4,168 ms, run `20261128T072109Z-61252-5189603b3fa6465985d0a46dbde2144d`, result SHA-256 `dbe6c63efeb53240db1a472cb665c8f7ca6277743cad5465f6313394c17fe27d`. Both scoped runs recorded zero failures and skips. The launcher `20261128T072014Z-78932-5f2a56b9867a41daa7203943de62a66b` drained at 08:22:04 CET, result SHA-256 `e0d0f061a64be3cbdf809034baed9d4500f5424e377b79b76fe9a7c048ed6956`. Seed `253058945` restarted at 08:22:19 CET on this same receipt. PR #403 merged at `039bda08afd85db2fd52b27de18f8a524f212b32` after the required checks passed. [Merged PR](https://github.com/RedSalamanders/RedSalamander/pull/403)

The complete seeded Commands run `20261128T072220Z-82632-512a7c7dd2dd41eb9f8887e92803d66b` exited 1 with 958 passes, seven failures and five skips. The corrected Compare case passed in 409 ms. Swap Panes failed while enumerating its saved left path during cleanup. The trace identifies the already removed fixture directory from the earlier Batch Rename source-replacement case and `ERROR_PATH_NOT_FOUND`. The Terminal lifecycle test passed. The later embedded tab, Shell exit, command-state and floating singleton cases failed to open another Terminal. The deferred plugin catalog and Terminal settings cases both reported `ERROR_BUSY` (`0x800700AA`) on creation. [Batch Rename fixture correction](https://github.com/RedSalamanders/RedSalamander/pull/405) [Terminal client correction](https://github.com/RedSalamanders/RedSalamander/pull/406)

Native results SHA-256 `a852f7192a2cb3c2e650b88464098cd3e3945d80b628b92e6274af9d067b6370`, shared trace SHA-256 `5087b8394ac8cb154c12a18ad22de99ce55e9f3509715af939dea1b574c129e7`, Commands trace SHA-256 `7facf2e8343c35e60811dd92aeba2d2262a2891a6ec3632a767a02378523b38b`. The contained launcher `20261128T072219Z-79260-e5756c04783c4a4c998ecdf70d851419` drained at 08:46:18 CET, result SHA-256 `30e9eb8b00b172758b16ae116e6b06a7e9c0dae2748056f0087963321cd82a6c`.

#### Ordered failure replays

- Batch Rename source replacement followed by Swap Panes reproduced the failed left-path enumeration. The replacement assertion passed and Swap Panes failed. Native run `20261128T074643Z-76260-8208e823b2cb4ab190ea124df3f1c66c`, results SHA-256 `cf68394ee60ee7fe587a55df5c33e48db456d1f68c803a39ffb4a5dc7c0df76e`, Commands trace SHA-256 `c644db39edb77778a7cebc7161a0ae9d79a5f699e10d2448a70e82b418d73eab`. Launcher `20261128T074643Z-76260-52fb9d63fb964398bd51c3d71024ee7b` drained at 08:47:37 CET.
- The Terminal sequence with seed `3` ran the embedded lifecycle, deferred catalog settings, embedded visibility and configuration round trip in that order. The lifecycle passed. Both creation checks returned `ERROR_BUSY`, and the tab did not reopen. Native run `20261128T074737Z-76260-a414abb2baaf4eeba167c46c6fdf3189`, results SHA-256 `6f3882ce1608feaf7a599965ae6f59116958d39a045239c9f13a5fd595fe33d4`, Commands trace SHA-256 `c72b8d33b483629e4b499abdb229160e9b11ede57f6b016872ce48dce7894515`. Launcher `20261128T074737Z-76260-7557e7de08a740a091ae9760dd0ade83` drained at 08:48:44 CET.

The local outer launcher initially supplied its own process ID in the native run identifier, so the canonical archive refused a different owner after both original cases had finished. The stale-run archive verified the owner exited, archived only these two runs, and preserved both original result hashes. The corrected launcher lets the native runner allocate its own identifier.

Exact Batch Rename correction `706790ac621f600b73d75c398a5fd2a5497ae3b3` built x64 Release at zero warnings and errors, receipt `f3cbf8b9b97d7f3da04e754b46d2c432a8a6c32f848ee96829986adb0d98fce9`. Pane refresh, source replacement and Swap Panes restoration passed in order three times, for nine passes with zero failures and skips. Native run `20261128T075721Z-70408-c19386ba025c45b5b7a1c5a43d02f3cd`, results SHA-256 `db3e6f255e96a97f7b9141e11810e67d0a9277123aa3b208dbc0c8ffcc257b4d`. Launcher `20261128T075720Z-77340-4e807dfa66fa481eb0b67df0cfcc6945` drained at 08:58:18 CET, result SHA-256 `f1a7d51eac242a10f862004b08c5fd9541f53d4bcdcb03831dcf459bd82632e4`. PR #405 merged at `ce3f6b218c5d3d96ae39ad7c987eba2c4e108c42`. [Merged PR](https://github.com/RedSalamanders/RedSalamander/pull/405)

The first Terminal candidate required the pinned vcpkg inputs in its worktree. After restoration, the zero-warning build guard rejected four implicit copy/move diagnostics on the test probe. The final candidate declares the operations deleted and built x64 Release with zero warnings and errors. Exact source `c3516178defea4c4bc90c32fe382e1a677bba2c9`, receipt `84f101096098b85f44db3c8498ee618ef21795fb92f5838ab00537a8f539f427`.

The corrected Terminal sequence ran lifecycle, catalog, tab reopening and settings in order for all three repetitions. All twelve passed with zero failures and skips. The result includes the normal client resolving the Terminal provider and successful creation after release, with the intentional deferred unload checks preserved. Native run `20261128T082445Z-55812-656f7ced4f2f494b95a358e9ce97e7ca`, results SHA-256 `e73b2dd79413fe3c8f0428867e6318f8a14d0752c2090b7f8c3cab574982abe4`. Launcher `20261128T082443Z-35592-7c1846213b2944a18b5b32f3c32e3a22` drained at 09:25:58 CET, result SHA-256 `f160d0635e8a84598461ac020a017fdf1f5acf9a5577d67c2726c65702cf5185`. PR #406 merged at `d76a3104d764a0de1a5f1ad54ad92fd9c5d147e8`. [Merged PR](https://github.com/RedSalamanders/RedSalamander/pull/406)

### Virtual-pane clipboard preservation

Exact `594d58373708f05ccc58f2d49efd963f450cd55a` built the test-enabled x64 Release solution with zero warnings and errors, receipt `12cf77b57ea452ab32bdfce24e9472440bf82e0a4a332d496b0496b5e337c502`. The real Windows Cut fixture placed two Local paths with Move effect on its clipboard. It then selected a virtual file and invoked both Copy and Cut through the pane command dispatcher. Each displayed the Local selection warning and preserved the clipboard generation, the two paths and the Move effect. The virtual fixture was removed.

All three repetitions of `cmd_pane_clipboardCut_sets_move_drop_effect` passed with zero failures and skips at 946, 709 and 703 ms. The filtered repository verdict is `NOT_EVALUATED`. Native run `20261128T070433Z-34864-e090bb7e70154097b2c3aa47892cee23`, results SHA-256 `a6aea2b91da64a4f217f9f3e718629866f65a9e2fa33103e41261793b64e990d`. The launcher `20261128T070432Z-80680-bd92df4748fc480eacd1d5f795f26bad` drained at 08:05:41 CET, result SHA-256 `75fb563d01ee83b567618b5468837af8b15cdcaa11c5dacd8e55feb9a75e5c02`. PR #402 merged at `7369f2d95b94aa4bece6c8f1344d3083d05878bf`. [Merged PR](https://github.com/RedSalamanders/RedSalamander/pull/402)

### Live SMB Recycle and scheduled Rename cancellation

Baseline `aad8b5f0a1fa08447425004a8950791510d9dfa2` built the test-enabled x64 Release solution with zero warnings and errors, receipt `5b0c8945242dc3b41cce64f3a8922eedbc6d69ea034411861e79c7afc64e0e38`. All selections ran without activation. The mixed Recycle selection used a fresh owned file on an approved SMB share and a separate Local file in the marked sandbox. The provider returned unknown Recycle availability for the SMB file. The task asked once for that group, Skip retained the exact SMB bytes, the Local Recycle completed, and cleanup purged one journaled test item. The SMB-only control selected a by-name prompt on Samba and an exact-identity prompt on Windows SMB.

| Baseline selection | Passed/failed/skipped | Native run | Results SHA-256 |
|---|---:|---|---|
| Samba SMB-only and mixed Recycle | 3/0/0 | `20261128T083418Z-36388-e2c7c82348e4490785ef16e5cf960353` | `d9bc70d21149e73e20056a6b8a1922a6a06d1967b79b0ebb14d1b507296aaf50` |
| Windows SMB-only and mixed Recycle | 3/0/0 | `20261128T083703Z-57496-be7934272a444ba3a37d49ddf521c649` | `22f4b6d040257223aae4ab179112a0cd20c749cf09b9cd16d8c2a30942d97225` |
| Windows SMB exact and scheduled Rename | 2/1/0 | `20261128T083748Z-57496-d8464ff837874bc38d9d63acc5a48a35` | `29f0be599160483283013e0b82bcfb42bf2fdb9fd7d6e222ee7b41690a5427cc` |

The mixed Recycle task settled in 126,757 µs on Samba and 56,816 µs on Windows SMB. The initial Windows Recycle selection also passed. After that run, the temporary fixture registry restore failed because the PowerShell call bound an empty backup argument. The original registry was restored from its unchanged backup after verifying both hashes. The repeated selection used the corrected restore and verified the original bytes after exit. The configured resources on both marked local test volumes were verified before the corrected build.

Both normal Windows SMB scheduled Rename origins completed, with measured task times of 36,927 µs for Batch Rename and 30,747 µs for Change Case. Cancellation reached the worker barrier and returned `ERROR_CANCELLED` in 49,391 µs. The case failed its combined assertion for exact retained fixtures and the canceled item result. Trace SHA-256 `4c29c01d6be78928ae897d729a55e06bddb9a0f07f8ba104862ff3e9f403b1c9`. The scheduler marked rows as possibly mutated on entry before its read-only preflight. The correction in #404 moves that transition to the execution handoff and measures the fixture checks and typed result separately. [Live fixture and scheduler PR](https://github.com/RedSalamanders/RedSalamander/pull/404)

Corrected source `48f917d31922c250d0826763aca7acccae7f97a3` built the test-enabled x64 Release solution with zero warnings and errors, receipt `7e62eb656bf4c557d6e9ea96096fc30afb37ad2ef7f0a828f4a09deab2842565`. All corrected selections passed setup, the selected case and cleanup without activation. The C and Z resource files matched their original hashes after the launcher exited. Filtered repository verdict `NOT_EVALUATED`.

| Corrected selection | Passed/failed/skipped | Native run | Results SHA-256 |
|---|---:|---|---|
| Windows SMB exact and scheduled Rename | 3/0/0 | `20261128T084929Z-55164-27d8399df0c948ef99cda6217295a1aa` | `6212557ca0e7dfcba8e43c1593bec6018c8ae482d5843c108a05f02d6df4ef90` |
| Samba SMB-only and mixed Recycle | 3/0/0 | `20261128T085012Z-55164-030270fbbaa247b58b770aefca4c7c06` | `6de231017deffbf21dd88390b12ae9393e6d02a6e89e52e9198fb5b91a0cb0db` |
| Windows SMB-only and mixed Recycle | 3/0/0 | `20261128T085051Z-55164-193b6279e943409eb84cbafab57537a4` | `414e15d1868ce36c0f67433a0ff898656fc1bf24950aad7c16ed9d844f345566` |

Windows SMB cancellation reached the barrier and returned `ERROR_CANCELLED`. The witness separately confirmed the exact original leaf and bytes, unrelated sibling bytes, and the absence of the requested destination. The one returned item reported `NotAttempted + Retained + Canceled` with the same status. Trace SHA-256 `2b19c6e8ee431abfe46507f43f5e511100401c58afc0b2b36fcc12a0b237245c`. The blocking watcher returned `ERROR_SHARING_VIOLATION` on both builds.

| Same live Windows SMB fixture | Baseline | Corrected |
|---|---:|---:|
| One exact Rename with independent rebind verification | 4,075 µs | 5,091 µs |
| 100 exact Renames with independent rebind verification | 370,381 µs | 333,820 µs |
| Scheduled Batch Rename | 36,927 µs | 40,555 µs |
| Scheduled Change Case | 30,747 µs | 30,641 µs |
| Cancel request to completion callback | 49,391 µs, result assertion failed | 51,410 µs, passed |

The timing observations are one sample per build. The corrected raw performance log SHA-256 is `5a32371d4f29c74ca0365e3fb2f04036bbd72294bd47385d4e4240911b5314fd`. Live legacy fallback and whole-process handle peaks are separate gates. PR #404 merged at `6213ed82ceb665140492f278c86df646a07f0a17`. [Merged PR](https://github.com/RedSalamanders/RedSalamander/pull/404)

### Integrated x64 Commands replay

The full Commands entry used exact PR #404 source `48f917d31922c250d0826763aca7acccae7f97a3` and x64 Release receipt `7e62eb656bf4c557d6e9ea96096fc30afb37ad2ef7f0a828f4a09deab2842565`. It used the earlier ARM64 shuffle seed `253058945` and included all three merged Compare, Batch Rename and Terminal repairs. The 970 cases returned 965 passes, zero failures and five opt-in ViewerSpace performance skips. The canonical runner reported no failure classifications and returned its scoped `NOT_EVALUATED` verdict with exit 2.

Native run `20261128T085225Z-65844-b52e0fcf7ccd4839a72403d8e8200038`, results SHA-256 `3de4ba8151c46e86eb6f4c09850620168deccd4e0339d60fbf277e6cf95446e1`. Contained launcher `20261128T085223Z-60992-56721d9950464adf86846cf9c0b08102` ended at 10:15:50 CET after its process group and output drained. Launcher record SHA-256 `a8979d2599ef9a1b107e8e1f7a668ecffc4128ab9c62d0d98779bcdd3f48674a`. This x64 replay used the earlier seed. The later hosted ARM64 attempts below used their own orders.

### Hosted ARM64 result on the corrected Pack tree

Hosted run `46089888885` built the test-enabled ARM64 Release profile on exact `73c70e33037273b3727835d9b61a265883f15470`, receipt `a6ab5c44ad473f4ae55770a8460092fadae2361bb08898a200219cf73c47e179`. ARM64 job `140881818883` failed and finished at 10:15:51 CET. Native run `20261128T062815Z-9560-f3f1fa1e96aa4c638ee888c5f332f6d3` promoted 11 entries, including Tooling and the Pack source guard. Commands remained marked running when evidence was captured, and seven entries remained pending, including File Operations and the four later desktop entries. Both Pack and Unpack command cases passed in the original and both shuffled attempts. [Hosted ARM64 job](https://github.com/RedSalamanders/RedSalamander/actions/runs/46089888885/job/140881818883)

| Commands attempt | Passed/failed/skipped | Recorded failure | Results SHA-256 |
|---|---:|---|---|
| Original | 963/2/5 | A console took foreground during the resource-format case's input settling. The two-pending-task disclosure case reached the expected action and state while its actual focus and foreground checks failed. | `6561a91d628854571e26ef43be3152412e2a04ee3694239822e1c2bd1e7f0661` |
| First shuffle | 960/5/5 | Two Connection Manager close cases, Preferences backdrop Apply, File Operations Preferences Apply, and ViewerSpace hover. A console owned the pointer during the hover failure. | `ba09a70ff0c621283567a0741e9ec28da03170d6ef911ff141c9ea0b5e2e3347` |
| Second shuffle | 964/1/5 | ViewerSpace hover again observed the pointer over a console instead of its test viewer. | `0043b31d1e7b5533c0eb56c8c57f03aa30d69650e1c7988717cd53f4fdfdc927` |

Both original failed cases passed their isolated ARM64 retries. The last shuffle child exited 1 and its process group drained. The retirement guard saw a new `ConsoleWindowClass` foreground and refused to replace it. The final hosted desktop snapshot identified WSL in the foreground, with a different PID from the earlier retirement observation. The logs leave the creator and the four first-shuffle UI Automation failures' causes unresolved. Production packaging was skipped.

Artifact `15257430951` and the original job log were retained in archive `20261128T091645Z-81340-359fb13819e34143a0cb4d56a0936518`. Job log SHA-256 `6b4f21ac5121a23e88b6bab292c7a9ac4bbfc5a26bb48c88f3f8d82c82ea078a`. Run-state SHA-256 `d05d43a3f35e286fef97dfcc44cf09a1c124786fe8ea01414113804ef9b8f914`. The later #406 and #404 workflows concluded at admission and skipped both native branches after main advanced. [#406 admission](https://github.com/RedSalamanders/RedSalamander/actions/runs/46126299605) [#404 admission](https://github.com/RedSalamanders/RedSalamander/actions/runs/46128227489) [Newer main candidate](https://github.com/RedSalamanders/RedSalamander/actions/runs/46129343995)

### Local scheduled Rename result replay

The first non-activating selection reported File Operations as `PENDING_INTERACTIVE` and ran no cases. The two focused Local controls then ran with the desktop lease on the same x64 Release receipt as the full Commands replay. The first checked the exact mutation's conflict and authoritative result after the preflight barrier. The second checked that rows the engine never reached remain `NotAttempted` with retained sources when an earlier operation stops the schedule.

| Selection | Passed/failed/skipped | Native run | Results SHA-256 |
|---|---:|---|---|
| `FileOps_BatchRenameExactResultCarrier` | 3/0/0 | `20261128T091923Z-49844-fcb79a14ad35404fbf6c5a329da2215a` | `cced19818ca13d31085b9cac2179a75136e1921bbb85873588ecfe1c8008cb38` |
| `FileOps_BatchRenameNeverRunRowsNotAttempted` | 3/0/0 | `20261128T092020Z-72380-196fb1aa3cfa4d91a0015dd9a3e6f416` | `6ef121a7219f1cf3933b02b31a47f90162349d2f9a2988bbec772e5fa5faa5a8` |

Both runners drained at 10:21:10 CET. The scoped repository verdict was `NOT_EVALUATED`.

### Guarded-output build and live legacy Rename

Exact `b618093c466aede952e78a23602c7e8b6871647f` built the test-enabled x64 Release profile with zero warnings and errors, receipt `b216bff5df3356165b278683f41df4f68720eea5a454c9f45f40dee8b5de3e87`. The earlier integrated build found seven implicit special-member warnings in main's settings guard and fixture, plus an unqualified task type in the new Cut witness. The guarded-output merge replaced the original settings guard. Its next build exposed the WIL wrapper-address type mismatch in the recovery record call and an implicit-constructor warning in the guarded replacement state. The corrected candidate uses the address of the owning wrapper and explicit state construction.

All three live Windows SMB selections passed setup, the exact Rename case and cleanup without activation. Each performed the blocking watcher, independent binding and content verification, both normal scheduled origins, and cancellation before dispatch. The destination-scoped Ex hook fired three times per run, for occupied-name refusal, free-name publication and exact case-only publication using the legacy no-replace API. The provider recorded 108 legacy attempts per run, including the three forced controls. Each selection restored the original C configuration byte for byte.

| Windows SMB run | Native results SHA-256 | Raw performance SHA-256 |
|---|---|---|
| `20261128T094247Z-38864-21e5136e3e524fd78b260325ac14e92d` | `a6cf0b14b4c4278264cc3538aa86a3948bd0e729e0d1f4bbe99bb9e743f41bd6` | `5f62fe2a3a6f5d1b134c3385702aab7892635622b174ae58b6bf0a8633e2cd23` |
| `20261128T094333Z-38864-1b25a80d2be244d3b998b2a3da9bdc25` | `42d0ed5fa1abf8bb03dd3e6d133ea2db8da6c8e5dd0ea5151c863d5382b40145` | `68cb374a201a292a7b25ddecc93dc32794f5efd3b7725c176fc65789b360ae66` |
| `20261128T094416Z-38864-7e3c7117e3794b7398217840fae6e0d3` | `e7104af57f1da4f37b809158c214838b2ce4dfb0b9240590a129a177ecbc8901` | `747233768ea1866f3031af5a091b723f02a9612c6ec8148c2da21b50025b5a9d` |

| Matched fixture observation | First | Second | Third |
|---|---:|---:|---:|
| One exact Rename with independent rebind verification | 3,322 µs | 3,550 µs | 3,591 µs |
| 100 exact Renames with independent rebind verification | 371,344 µs | 331,888 µs | 340,152 µs |
| Legacy occupied-name refusal | 4,364 µs | 3,758 µs | 4,390 µs |
| Legacy free-name publication | 2,831 µs | 3,148 µs | 3,159 µs |
| Legacy case-only publication | 2,602 µs | 2,897 µs | 3,455 µs |
| Scheduled cancel request to callback | 49,643 µs | 62,961 µs | 52,197 µs |
| One item and 100 items, before → sampled process handle peak | 2,645 → 2,647 | 2,645 → 2,647 | 2,645 → 2,647 |

These handle samples include the source, publication and independent rebind while their owners remain live. A transient handle can fall between the sample points. Samba run `20261128T094459Z-38864-5ae4e8c4ba2d4062b721408cfaa81797` passed setup and cleanup with one exact-binding skip. The provider opened the owned source and established its exact binding limit. Both scheduled origins refused before mutation and preserved the fixtures. Results SHA-256 `ba6a920e43942416d7b27296f30582ed1da7f9f8bafd3086e61ff0bf74a49221`. The scoped repository verdict for these runs was `NOT_EVALUATED`.

### Copy-only Cut and staged metadata

Exact x64 Release `b618093c4` ran `FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof` three times without activation. Each queued admission presented one Copy/Cancel question with the Cut warning. Two declines on the same synthetic Cut generation left the consumption callback unused and preserved the task order. The same case exercised the accepted Copy-only transfer, late writer proof, the guarded changed endpoint, and mixed Native and Copy-only outcomes. All nine setup, selected and cleanup cases passed with zero failures and skips. Selected case times were 250, 218 and 235 ms.

Native run `20261128T094541Z-38864-ec73557c15c846ca90bb8622576bc86b`, results SHA-256 `1c69aa00288740b4add00c1a7ea162413d64fb34c4a1f660c3ccdf2d441072ef`, trace SHA-256 `b344f4b1500500542462c28f0a8e128cc12805b6dbddc1f19318e28ac5e34a4e`. This is the controlled admission witness. The real Windows clipboard and task execution have separate desktop evidence earlier in this record.

The staged Skip baseline passed setup, the selected case and cleanup without activation. It preserved both file contents and both source streams and retired the unpublished stage after one metadata-loss decision. The explicit metric check caught a coverage gap. `FileOps.Metadata.StageSkipped` recorded mask 6 for Mark of the Web and the named stream. The test-only presence override ran before backup inspection cleared the stream feature bits, so the simulated extended-attribute bit never reached the decision.

Native run `20261128T094627Z-38864-0b7c0bbe72f34667adca0169adba0b8e`, results SHA-256 `9c27373d5b1ced0b2eaf87eb10d388b423ceeda02e61c711e5041bedf205b64f`, raw performance SHA-256 `efe2b84d1069d2a5834fc09f482e6fc0a85a5b33f92348870ef324b9eb967d13`. Both sandbox resource registries matched their original hashes after the selections.

The correction applies the test-only presence mask after the actual backup-stream scan. Exact source `877f278c519d793634bd39c47b2fc25e1f8167cd` built x64 Release with zero warnings and errors, receipt `2d8f5e38f7408493a1aeffa63a1db845243014f4d103ebd61c0f32589590b5d2`. The non-activating replay passed 3/0/0 and confirmed the preserved files and streams and the removed stage. Exactly one `FileOps.Metadata.StageSkipped` row reported mask 14 for the combined loss, with the extended-attribute bit supplied by the test hook. The selected case took 87,453 ms, including its other remediation controls.

Native run `20261128T095900Z-83640-223ac72abeb34d8baf9f7a53055878aa`, results SHA-256 `a38de9b213ba2895f148a12111cc318d837c817d0a785764a9a1b446cd28c572`, trace SHA-256 `f25c95bd8f6c4b73fc702c1f31966306ad00144171fd4536f2943cb1a3896da2`, raw performance SHA-256 `0753205c2b2fee643a1dbd98eecf3bc576b9ff35e8b7347f8e355105754e7a62`. The focused repository verdict was `NOT_EVALUATED`.

### Guarded Pack and Unpack recheck

The exact guarded-output x64 Release build `b618093c4` passed all three non-activating repetitions of `cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation`, at 2,729, 2,456 and 2,469 ms. Native run `20261128T094906Z-73864-a1e648b9179b43fcbf68697962c261c0`, results SHA-256 `facd9964db574b58d1bb41909c084837028e0d30b662e166eacd3efacc724795`. The scoped repository verdict was `NOT_EVALUATED`.
