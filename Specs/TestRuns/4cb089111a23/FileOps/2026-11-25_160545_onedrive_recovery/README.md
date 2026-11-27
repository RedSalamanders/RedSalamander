# I66 recovery of earlier OneDrive Personal fixtures

The two runs used an isolated copy of the approved test profile and its
saved sign-in without activating the desktop. The first run verified Recycle
of the earlier roundtrip seed and the empty Move directory. OneDrive refused
to recycle the Upload directory nonrecursively with `0x80070091`.

The second run recycled the files and empty directories from the earlier
Rename fixture and verified that exact root was absent. A separate read
against the approved sync folder also confirmed absence. It retained the
remaining Upload directory when the expected file could not be matched to
the original local upload bytes, returning `0x8007051A`. An independent sync
read reported the parent and its one Upload child. This does not establish
a terminal result for the earlier Copy.

| Recovery-only source | x64 Debug build receipt | Governed run | Outcome |
| --- | --- | --- | --- |
| `2ce6d94b8` | `09d9fbcebc4ba7b2be154d4e2748076573ed6b190f3e08e0b2099cbcf4106322` | `20261125T145221Z-84500-45eca834211347f686c0e3bf57028827` | 2 passed, 1 failed, 0 skipped |
| `a08db7f8d` | `c16e7ddf4fba62acde145ed1b10860d2b9cddb6fc8982394353fdf0c3be8fd25` | `20261125T150545Z-68624-6138d1b422834569a3160f0abba000b0` | 2 passed, 1 failed, 0 skipped |

Both builds emitted zero warnings and errors. Both runs verified the
installed settings digest and removed the isolated settings files before
archival. The local recovery sources reused the focused selector to launch a
bounded fixture-specific body. The original broad timeout and failed cleanup
remain in the [earlier run records](../2026-11-25_144010_graph_live_preflight/README.md).

## 2026-11-27 exact recovery recheck

The approved OneDrive Personal sync folder showed the old Upload directory present with no direct children. The same recovery-only `a08db7f8d` x64 Debug executable passed byte verification on its recorded receipt. It used a fresh isolated profile and the original archived upload payload as the verification input. The exact provider recovery returned native exit zero and 3 passes, zero failures and zero skips in `20261127T104046Z-68048-e1b06bd2ae174334bf25653349222c76`.

The recovery used nonrecursive Recycle and recorded that both exact historical fixture roots were absent. An independent sync-folder read confirmed both absences. The launcher verified the installed settings digest and removed its isolated profile before archival. Native result SHA-256 `0d324f8e8f732ee18c0494bc881eea829054381d68f639de16aa394c0f1a706f`. Archived input SHA-256 `826d03fd16604a69f74ed238cb50f97edf3f3cd2cbcacbb3b05fd49c1a8ada6f`. The earlier Copy remains without a terminal result.
