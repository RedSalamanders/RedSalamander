# Guarded renderer passes; resource admission rejected

Source `bdf657a9559ed63579ad48e30f6fb306eb0e621a` built x64 ASan Debug
and Release with zero warnings/errors. All 54 focused native checks passed without
skips: 19 seed/parser/service checks and 35 ViewerSpace checks. The original
100 ms Close case passed three ASan and two Release repetitions, with new held
ownership, module-wide capacity, synchronous HWND destruction and drain checks.
ASan trace analysis records all 30 admitted device releases off the UI thread;
the maximum whole native Close span across this cluster was 91,013 us.

ASan scratch-storage checks also verified individual object red zones and ended
allocation poisoning. Native assertions, watchdogs and sanitizer settings remain.
These passes do not clear the earlier failed Full.

| Profile | Seed time median ratio | Sampled private-memory median ratio | Admission |
| --- | --- | --- | --- |
| ASan Debug | 0.282515 | 0.675939 | PASS |
| Release | 1.055365 | 1.113858 | FAILED |

Release sampled peak private-memory median was 940,138,496 bytes versus the
retained 844,038,144-byte baseline. The failed comparison remains failed. No Full
was started, published code is unchanged, and no baseline or limit was replaced.

The live-volume corpus changed: the first retained Release baseline enumerated
4,169,603 records, the earlier scratch candidate 4,171,702, and this candidate
4,205,320. The installed MSVC hash implementation rounds bucket counts to powers
of two; crossing 4,194,304 records at load factor one is a plausible explanation
for much of the additional bucket storage. This is an inference from recorded
counts and implementation inspection, not a measured bucket count for this run
or proof that all memory differences have one cause. `bucket-growth-analysis.json`
binds that analysis to the installed header digest without copying its source.

`manifest.json` binds 43 complete raw payloads. Two viewer metric selections are
explicit UTF-8 reencoded subsets, with source digest/size and selected line numbers;
they are not complete raw traces. Full originals remain in their owned E-root runs.
The next candidate uses modest temporary-map density and records actual bucket
counts. It must pass the original resource admission and final qualification.
