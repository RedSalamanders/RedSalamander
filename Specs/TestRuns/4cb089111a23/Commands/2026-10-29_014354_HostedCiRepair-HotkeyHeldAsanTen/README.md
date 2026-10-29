# Held hotkey: ten ASan repetitions

Source `7a3a7057669c5cd9d2ae5c629133f6cdebd23875`, full-solution x64 ASan Debug
Embedded receipt `b6bdd3222d79f32b820af42923b4f0edb0fba1b2445acb006c3b4b1c33e70232`.
Run `20261029T014354Z-41464-bed122fc907c49c68180897a1c1d61f9` passes ten repetitions
with zero failed/skipped/retries and canonical wrapper exit zero. Native duration
is 3,673 ms. All twenty independent real dispatch witnesses see F24 held and
verify exact original hook/property retirement. Actual foreground, final key release,
registration reuse and refusal controls remain unchanged, including the 500 ms bound.

The original-helper control is retained separately and fails the stronger key-state
oracle before the fix. This ASan run retains final linker PDBs and instrumentation.
It is local focused acceptance, not Full or native ARM64 hosted qualification.