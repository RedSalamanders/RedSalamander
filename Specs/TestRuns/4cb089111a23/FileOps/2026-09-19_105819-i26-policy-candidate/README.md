# I26 four-primary-action policy — candidate

Focused x64 Debug qualification of engine-owned primary placement and Replace/Remplacer labels.
Build: zero warnings/errors. Setup, `Phase10_DeferredConsentAndRecycleEscalation` and Cleanup all
passed (3/3, zero skips). This is focused evidence, not full redesign qualification.

- Runner: `20260919T105419Z-9864-eeb40d5c1084458f9c4e6009023f8957`.
- Receipt: `984e876a986d02cc271f0b1af75d9c8995417bcd4a06749e166b9bd300325b18`.
- Source snapshot: `2a7136c4614ba746c09d426ba72792c9212510c290ccf6d2c4539667603e67f9`.
- HEAD: `a212edec433d658748ae1273cba68633c86de8b2`, with dirty implementation/test inputs.
- Exact fixture SHA-256: `90688FBD9B4343BBF323622013B3AF17AA3A45CB87D8A75551093123BC8A612E`.

The identical retained fixture covers 24 classes x 64 permission/scope combinations, unique complete
primary/overflow partitioning, absent authority, typed defaults/Escape, and the existing explicit
class/loading/scope table. Eligible collision choices are KeepBoth, replacement, Skip, Cancel;
unavailable actions are omitted. SkipAll remains overflow. Same-host and Recycle semantics remain
authoritative. Underlying allowed-action ordering, operation dispatch and cached consent are unchanged.

Paired diagnostic `FileOps.ConflictPolicy.MatrixBatchUs`, microseconds per 24576 policy constructions:

| Input | Five rounds after one full warm-up | Median |
|---|---|---:|
| [Unchanged placement baseline](../2026-09-19_105159-i26-policy-baseline/README.md) | 32802, 33219, 35004, 33337, 33084 | 33219 |
| Candidate | 34039, 34250, 34666, 34292, 34199 | 34250 |

The Debug median increased 3.1% (1031 us per batch). This is diagnostic construction-cost evidence,
not Release performance acceptance or a rendering/input/I/O non-regression claim. Required paired
Release/resource qualification remains open. Both runs use the exact same fixture; the checksum
changes from 3490816 to 3454976 because placement changes. Full raw metrics are retained.

Command: `Run-AllTests.ps1 -Suite FileOps -CaseFilter Phase10_DeferredConsentAndRecycleEscalation
-Platform x64 -Configuration Debug -TestRoot C:/RedSalamander.Perf -TimeoutMultiplier 3`, with
`RSBuildEnableTests=true`. Full log: `C:/RedSalamander.Perf/evidence/i26-ui/i26-policy-candidate.log`.
The separate sandbox audit still reports 225 preserved pre-existing run directories; this is not
clean Full closeout evidence. Native layout and action acceptance are additionally exercised by the
[Commands integration case](../../Commands/2026-09-19_110240-i26-policy-integration/README.md).
