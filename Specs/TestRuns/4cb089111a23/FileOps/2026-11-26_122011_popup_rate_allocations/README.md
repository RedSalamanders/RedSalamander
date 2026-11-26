# I66 matched popup rate allocations

The baseline is the last main commit before credential-wait progress, `8c7775db997ec4fc80a6cd3f9ed9c68bf3f2daa9`. The candidate is the integrated active-peer source at `78b8c46d35a0c8d3a11cd944016a4be0269eec19`. The same machine ran both test-enabled x64 Release Embedded builds. Their build arguments, compiler, vcpkg and Ghostty identities matched. Both exact `FileOps_PopupIdleCadence` selections passed setup, validation and cleanup without activation.

| Evidence | Baseline | Candidate |
| --- | --- | --- |
| Build receipt | `8de2f82b17c06d784425bf0a2d687c37dd4ec332bc633e325499552029c4cd56` | `90b1cf616a1dc1d7ad25f3ba1ae294c31c29ec394da8d37ed5f5753fbc8b8963` |
| Native run | `20261126T111819Z-76456-97f19ad3afb6478ab89cf41163e0f744` | `20261126T111414Z-66144-696b309e6d114d1cb6ce04a37c3b86f6` |
| Owned heap capture | `20261126T111817Z-632-f130cf5d1fc145c39845f491fbe1f5ca` | `20261126T111411Z-74148-12d8760ba5c2459a92014bda6f0e0d3c` |
| Stack analysis | `20261126T112011Z-6916-998dc073f88148bca8d30bedc66896c4` | `20261126T111930Z-9520-5b47bb1619974cdb8f532cba2d2c74f8` |
| Full traced rate samples | 92 | 93 |
| Stack-attributed allocate and reallocate events | 2,231 | 2,280 |
| Reported ETW event and buffer loss | 0 / 0 | 0 / 0 |

A per-process Xperf heap session captured the runner-owned native PID. The analysis loaded each exact Release PDB, selected stacks in `FileOperationsPopupState::UpdateRates`, joined each stack to its heap event by QPC and matched it to one native rate marker within six milliseconds. The ETL start event and FILETIME supplied the clock origin. Both sides attributed every rate marker and every selected stack. The earlier lost and wrapped traces remain in their governed archives. [Microsoft heap capture reference](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/enabling-heap-data-capture)

| Matched engine window | Rate snapshot shape | Baseline | Candidate |
| --- | --- | ---: | ---: |
| Prompting | 5 tasks and rates, no updates, no silent tasks | 1 sample × 21 | 1 sample × 21 |
| Running | 6 tasks and rates, no updates, 1 silent task | 9 samples × 24 | 11 samples × 24 |
| Paused | 6 tasks and rates, no updates, no silent tasks | 2 samples × 23 | 2 samples × 23 |

Each matching window kept the same allocation count per sample. The baseline was still compiling during the candidate capture. The separate untraced run supplies cadence and latency evidence. The candidate's modeled credential-active-peer interval captured eight samples across its transition from six to eight task snapshots. Live authentication and a matched Stopping allocation window remain open. [Untraced active-peer record](../2026-11-26_095648_credential_active_peer/README.md)
