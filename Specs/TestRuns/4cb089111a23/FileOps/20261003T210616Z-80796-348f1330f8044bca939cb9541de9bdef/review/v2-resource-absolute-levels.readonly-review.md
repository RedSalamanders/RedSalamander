# Integrated v2 resource levels: pending root review

The ten closed originals confirm median final-minus-cold process-handle deltas of baseline **36** and candidate **37**, a candidate difference of **+1**. Baseline pair 4 remains an adverse **+265**. Neither observation has a proved resource owner or leak cause. This review does not qualify resource nonregression or resolve those observations by excluding them.

Read only the ten literal `interleaved-run-bindings.json` archives in `v2-pairs-20261003T210306Z-487089b39880442e8158b96b95fc97dd`. Streamed each archived metric file without loading it whole. The native metric names are `Memory`, `Cpu`, `Handles` and `GuiAndAttachments` under `FileOps.IntegratedPresentation`; no literal `Resource*` metric exists. Each original has 116 resource rows covering 29 phases. The companion JSON retains every selected original row string and line number, all absolute phase levels, IDs, receipt IDs and recorded archive SHA256. Resource row machine/process/native-run/commit/Release/hr0 identities agree with capture controls; copied raw SHA256 agrees with controls. Original hashes were not recomputed. No native execution, source edits, builds, tests or desktop interaction occurred; the proposed curator remains unexecuted.

For this table, warm means the exact `popup-open-before` witness; final means `scan-file-list-closed`. Values are absolute process counts, followed by final minus cold.

| Run | Handles cold / warm / final | Handle delta | USER cold / warm / final |
|---|---:|---:|---:|
| Baseline 1 | 2644 / 2667 / 2680 | +36 | 72 / 76 / 75 |
| Baseline 2 | 2644 / 2667 / 2680 | +36 | 72 / 76 / 75 |
| Baseline 3 | 2644 / 2669 / 2681 | +37 | 72 / 76 / 75 |
| Baseline 4 | 2436 / 2681 / 2701 | +265 | 72 / 77 / 77 |
| Baseline 5 | 2644 / 2667 / 2679 | +35 | 72 / 76 / 75 |
| Candidate 1 | 2684 / 2707 / 2721 | +37 | 72 / 76 / 75 |
| Candidate 2 | 2644 / 2667 / 2681 | +37 | 72 / 76 / 75 |
| Candidate 3 | 2644 / 2667 / 2681 | +37 | 72 / 76 / 75 |
| Candidate 4 | 2646 / 2669 / 2681 | +35 | 73 / 77 / 76 |
| Candidate 5 | 2642 / 2667 / 2679 | +37 | 72 / 76 / 75 |

| Run | Private bytes cold / warm / final | Final minus cold |
|---|---:|---:|
| Baseline 1 | 275791872 / 267894784 / 265842688 | -9949184 |
| Baseline 2 | 271650816 / 276725760 / 265048064 | -6602752 |
| Baseline 3 | 271654912 / 276549632 / 264871936 | -6782976 |
| Baseline 4 | 256540672 / 262385664 / 265080832 | +8540160 |
| Baseline 5 | 268505088 / 273399808 / 265957376 | -2547712 |
| Candidate 1 | 275283968 / 269250560 / 266096640 | -9187328 |
| Candidate 2 | 272519168 / 267476992 / 265519104 | -7000064 |
| Candidate 3 | 268513280 / 267534336 / 265154560 | -3358720 |
| Candidate 4 | 271691776 / 267882496 / 264495104 | -7196672 |
| Candidate 5 | 271581184 / 268296192 / 266588160 | -4993024 |

For all ten runs, attached hosts and current UI-thread shared-host attachments are **7 / 8 / 7** at cold / popup warm / final. Every `popup-destroyed`, `issues-destroyed`, `issues-cleaned` and four `scan-*-closed` witness has **7 hosts and 7 attachments**. All ten cold and final GDI levels are **59 / 59**. These observations establish only those sampled counters. USER ends +3 above cold in nine originals and +5 in baseline 4. Private bytes are process-wide point samples, not driver allocations or a peak measurement.

Lifecycle semantics come from [the exact fixture](C:/RsI27Followup/RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.IntegratedPresentationPerf.cpp:139). Memory maps value0/private bytes and value1/working-set bytes; Handles maps value0/process handles and value1/attached hosts; GuiAndAttachments maps durationUs/thread attachments, value0/GDI and value1/USER (emission lines 170–173). CPU values are cumulative counters, not percentages. Snapshots are taken at `sample()` calls, stored, then emitted together by a scope-exit guard; native row timestamps mark emission, not individual sample times.

The [cold sample](C:/RsI27Followup/RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.IntegratedPresentationPerf.cpp:234) occurs after language/theme/settings changes and settings-save settlement in an already running owner process. The [popup warm sample](C:/RsI27Followup/RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.IntegratedPresentationPerf.cpp:373) occurs after create/show, fixed geometry, queue preparation and eight redraw warmup frames. Popup and Issues scopes destroy their local windows before their destroyed witnesses (lines 479 and 649). Issues cleanup at line 659 follows removal of synthetic diagnostics. Each scan closed witness at line 871 follows its local popup scope and explicit host/attachment equality checks. These are surface lifecycles while the owner and process remain alive. There is **no resource sample after final settings restoration (lines 875–878), owner destruction or process exit**; no post-owner-cleanup witness can be inferred.

Baseline pair 4's anomaly begins before the first warm witness: **2436 → 2681 (+245)** by popup-open-before, followed by **+20** to final. Other baseline cold levels are all 2644, and their final levels are 2679–2681; pair 4 ends at 2701. Thus the +265 is not solely an unusually low starting count, and most of its increase precedes measured popup redraw batches. The cause remains unknown. Preserve its low cold private level, +8540160 private-byte delta, USER +5 and all intervening witnesses; do not discard or silently normalize it.

The +1 median is also not a uniform candidate increase: within-run handle deltas are baseline `[36,36,37,265,35]`, candidate `[37,37,37,35,37]`; paired differences are `[1,1,0,-230,2]`, median +1. Absolute cold levels differ, notably candidate pair 1's 2684 and baseline pair 4's 2436. These samples lack handle identities, allocation stacks or an owner-teardown comparison, so they cannot identify one persistent candidate-owned handle.

Focused F4 review: [ArmCompletionRecoveryTimer](C:/RsI27Followup/RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp:2766) installs a 100 ms timer on the existing owner HWND; [owner dispatch](C:/RsI27Followup/RedSalamander/FolderWindow.FileOperations.cpp:4442) kills it when no deferred presentation delay remains. [Orphan completion recovery](C:/RsI27Followup/RedSalamander/FolderWindow.FileOperations.State.Queue.cpp:2033) transfers an existing worker `std::jthread`, joins it, then allows owner consumption; its native handle is not a newly created owner event. The focused F4 fields include mutexes/condition-variable/counters, not an explicit new owner-retained HANDLE creator. This does not exclude library/runtime resources or prove a timer's relationship to process handle counts. No recorded handle identity or sample connects F4 to the +1; **F4 causality remains unproved**.

Root review should accept the lifecycle interpretation only after inspecting the fixture references and complete absolute phase JSON. Retain the +1 and +265 as unresolved resource observations unless independent evidence supplies ownership and later cleanup. The present evidence supports sampled surface-host teardown and the exact observed levels; it supplies no leak verdict, p95 inference, GPU allocation/peak claim, Full or I27 qualification.
