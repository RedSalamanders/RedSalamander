# I66 live legacy Rename, guarded publication and Release build qualification

Local profile `4cb089111a23` and hosted Windows x64 and ARM64 Release. This record preserves the 2026-11-28 captured build, provider and self-test results for I66. Source commits, receipts and supported claims are given with each selection.

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

Native run `20261128T094541Z-38864-ec73557c15c846ca90bb8622576bc86b`, results SHA-256 `1c69aa00288740b4add00c1a7ea162413d64fb34c4a1f660c3ccdf2d441072ef`, trace SHA-256 `b344f4b1500500542462c28f0a8e128cc12805b6dbddc1f19318e28ac5e34a4e`. This is the controlled admission witness. The real Windows clipboard and task execution have separate [desktop evidence in the earlier archive](../../../../../Specs/TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md).

The staged Skip baseline passed setup, the selected case and cleanup without activation. It preserved both file contents and both source streams and retired the unpublished stage after one metadata-loss decision. The explicit metric check caught a coverage gap. `FileOps.Metadata.StageSkipped` recorded mask 6 for Mark of the Web and the named stream. The test-only presence override ran before backup inspection cleared the stream feature bits, so the simulated extended-attribute bit never reached the decision.

Native run `20261128T094627Z-38864-0b7c0bbe72f34667adca0169adba0b8e`, results SHA-256 `9c27373d5b1ced0b2eaf87eb10d388b423ceeda02e61c711e5041bedf205b64f`, raw performance SHA-256 `efe2b84d1069d2a5834fc09f482e6fc0a85a5b33f92348870ef324b9eb967d13`. Both sandbox resource registries matched their original hashes after the selections.

The correction applies the test-only presence mask after the actual backup-stream scan. Exact source `877f278c519d793634bd39c47b2fc25e1f8167cd` built x64 Release with zero warnings and errors, receipt `2d8f5e38f7408493a1aeffa63a1db845243014f4d103ebd61c0f32589590b5d2`. The non-activating replay passed 3/0/0 and confirmed the preserved files and streams and the removed stage. Exactly one `FileOps.Metadata.StageSkipped` row reported mask 14 for the combined loss, with the extended-attribute bit supplied by the test hook. The selected case took 87,453 ms, including its other remediation controls.

Native run `20261128T095900Z-83640-223ac72abeb34d8baf9f7a53055878aa`, results SHA-256 `a38de9b213ba2895f148a12111cc318d837c817d0a785764a9a1b446cd28c572`, trace SHA-256 `f25c95bd8f6c4b73fc702c1f31966306ad00144171fd4536f2943cb1a3896da2`, raw performance SHA-256 `0753205c2b2fee643a1dbd98eecf3bc576b9ff35e8b7347f8e355105754e7a62`. The focused repository verdict was `NOT_EVALUATED`.

### Guarded Pack and Unpack recheck

The exact guarded-output x64 Release build `b618093c4` passed all three non-activating repetitions of `cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation`, at 2,729, 2,456 and 2,469 ms. Native run `20261128T094906Z-73864-a1e648b9179b43fcbf68697962c261c0`, results SHA-256 `facd9964db574b58d1bb41909c084837028e0d30b662e166eacd3efacc724795`. The scoped repository verdict was `NOT_EVALUATED`.

### Hosted guarded-settings build gate

Run `46129343995` admitted exact `1e05fafbc7ee8f97083df6c6cc1d89f7f67f6903`. Both Release jobs stopped at the build gate with the same seven warnings and zero errors. The settings publish guard triggered four warnings for implicitly deleted construction and assignment. Its legacy-fallback fixture triggered three more. Both jobs stopped before native regression. [Hosted x64 build](https://github.com/RedSalamanders/RedSalamander/actions/runs/46129343995/job/140933400082) [Hosted ARM64 build](https://github.com/RedSalamanders/RedSalamander/actions/runs/46129343995/job/140933400142)

The local build had reproduced those warnings. The subsequent guarded-output main merge replaced the old settings guard. The corrected x64 Release candidate explicitly constructs the newer save and displacement state, passes the owning WIL wrapper address and gives the test fixture its own path. That build completed with zero warnings and errors.

| Original job log | Local archive | SHA-256 |
|---|---|---|
| x64 | `20261128T101715Z-64280-44a0cfdc5bbf4ed0a884f351ec82a8cf` | `4e493416a6b7571ba510cde16b5b3ae9435e73e8a8219e652c04cb5b677d038c` |
| ARM64 | `20261128T100456Z-13664-0fefc9a123ef4ebd907e7fc0f2e29dd9` | `42bc90c807c071d556069c91c46054ed3c11a3a0fb4d496c8d842c025f29a49f` |

### Hosted standard-preprocessor build gate

Run `46133779153` admitted main `d0d61c5ccdc09364b76943df4bc737ace0aa372a` with the new CI engine and the standard MSVC preprocessor. Both test-enabled Release jobs stopped before native regression. Each logged one C2664 error and five warnings. The settings recovery record passed `void**` where `Commit` requires `wil::unique_hfile*`, because the WIL owner overloads address-of. The warnings came from implicit special members in `PendingSave` and `Displacement`. [Hosted x64 build](https://github.com/RedSalamanders/RedSalamander/actions/runs/46133779153/job/140946508411) [Hosted ARM64 build](https://github.com/RedSalamanders/RedSalamander/actions/runs/46133779153/job/140946508421)

| Original job log | Local archive | SHA-256 |
|---|---|---|
| x64 | `20261128T103549Z-27444-c2b4a9af60304dd9957ff9d4723eea33` | `0321d49eb976dec1f7b729fc39479bd698a78d9f82ce6986006122e4bbd9f232` |
| ARM64 | `20261128T103549Z-54792-ec1b43c210d74c1c988af6750e36c394` | `11c6cea2c9037190c91cef2d05dce01076decb807e351f5f5b26db9148aa3261` |

The local integrated build first reached missing Shell API declarations in the icon cache and host argument parser after the guarded settings fix. The explicit Shell headers cleared that diagnostic group. The next x64 Release build reached standalone Viewer PE and failed in the Windows SDK UI Automation declarations. Its log, `msbuild-20261128_112916_123-pid78268-8cf6ca7b.log`, has SHA-256 `75800dac0dd42bb7e587fa59d10de6774e7b45acfda9d955aae70d97e7ce3fe7`. The next candidate adds the Win32 and COM prerequisites before the Viewer PE UI Automation include.

### Integrated local x64 Release rebuild and routing check

Exact `3db5a417702763ef5399544e1946c90dabddf68b` built with zero warnings and errors using the updated preprocessor and explicit Shell and Viewer PE headers. The test-enabled receipt is `a9e9c4e816115a850ee66bdedb476ce9a0ea1252bfb8e7adbc65a9f33d0ef2d6`.

The focused FailureReportRouting Pester selection passed all 20 checks. It verified the single exact production crash-probe exception, ordinary routing immediately after that dispatch check, the exact four product switches, and both Release ProductUi assertion switches. Sandbox archive `20261128T104626Z-22496-8fd118129fbf4b56a1c2cdc9f79a2e0c` records result SHA-256 `a52a08f8d8ec8240f215b43cd50491d5ac42a9bae44c0d7584d0982bd36a06b0`.
