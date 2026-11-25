# I66 Copy-only Move when destination proof changes

The existing Local-to-Dummy control retains its source and checks both file bodies.
The second real task accepted one Copy-only question and reached the preparation
barrier without destination writer proof. Its private destination then advertised
writer digest proof. The task still completed with S_FALSE, retained the 65,536-byte
source, published the same bytes, and displayed no additional question.

- Full test-enabled x64 Debug build at `bd7e133bf` emitted zero warnings and errors.
  Receipt `83763f1817f39b901b67103598d58cac0ccbd46c452348055a465e61d7a3a910`.
- Exact non-activating run `20261125T132735Z-27364-4e82094f79da4bb49d7043ffce5bb3fb`
  passed setup, the two-task case and configuration cleanup. The case took 266 ms.
- The focused runner plan passed 66 of 66 in
  `20261125T131655Z-5784-d89eac6bee8b43038214475c0975ed39`.

Mixed selections, changed endpoints and clipboard preservation remain separate
qualification rows. The focused runner reports repository coverage as NOT_EVALUATED.
