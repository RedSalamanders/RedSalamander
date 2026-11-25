# I66 Unpack Delete after source identity

The ZIP command case extracted an unchanged compressed archive and removed
that exact file after consent. It then paused cleanup before the Delete bind
and replaced another approved archive. The task reported the identity change
and kept the replacement and the moved original. A third run within the case
edited the archive in place, kept its size and restored its last-write time.
The Local revision detected the change and the edited file remained.

- Test-enabled x64 Debug source `5ac478f9be90cbbb00c46c0e6b8acb7278890ed8`
  built with zero warnings and errors on receipt
  `e8cef9aad010e6586d88acdc51b3d5f9d88720c56da22262669501c3729c305a`.
- Non-activating run `20261125T170848Z-64868-04786ec6dd31403b93f90f28f266b6a6`
  passed the exact ZIP case in 1,243 ms with zero failures and skips. The
  repository verdict for this selection was `NOT_EVALUATED`.
- The first Debug candidate `03f0f4c91` built, but run
  `20261125T165603Z-70256-fbfdc954041048b88589043a14195ac1` failed before the
  expected confirmation. The Local provider had supplied its persistent
  identity without a revision. The failure is retained. The corrected
  provider supplies a revision only when the host requests it.
- The paired Debug command samples each extracted the same 16 entries and
  23 payload bytes. Plain Unpack took 80,381 µs on `63504c595` in
  `20261125T161424Z-8584-f111c4d7cf174542b2d63fa66cce8229` and 101,942 µs
  on `5ac478f9b`. The positive and negative cleanup checks ran after that
  sample. Each timing comes from a single small archive.
- Focused run `20261125T170419Z-36572-4dd117826c864851a95347102df8e66c`
  passed 67 runner-plan checks and 276 source and tooling checks.

The exact PR head `bee641dfc` built for x64 Release with zero warnings
and errors on receipt
`e22001d7917501f2fe2697ae17c67312619c755a11311f87077cbaa65d407937`.
Non-activating run `20261125T172132Z-87204-de1de0fff1a74e5eafd6bad097707841`
passed in 1,266 ms with zero failures and skips. The whole evidence
inventory passed all 18,733 tracked files. [PR #292](https://github.com/RedSalamanders/RedSalamander/pull/292)
merged at `4e24d76ca` after the required checks.

Full hosted Release and archive worker migration remain open.
