# I66 Cut and Paste clipboard generation

Paste captures the file list, preferred drop effect and Move sequence during one
clipboard open. The real pane journey held an admitted Cut/Paste in Preparing,
then used pane Copy to publish a newer generation. The stale Move completed
without mutation. Both files kept their bytes and the newly published clipboard
remained intact. A fresh Cut then completed the existing retained-source recovery
journey. The fixture restored and verified all original clipboard formats.

- Test-enabled x64 Debug `35cf70eefe6d7814c90d8542ad9a1a0c1a931a9b` built
  with zero warnings and errors on receipt
  `5f5d222a2319a0546a53959c559aa324ca8791fdf692a7bfc684c25c47d54b16`.
  The first build at `d43676638` was rejected with two C4389 warnings in
  the optional drop-effect comparison. The accepted build compares DWORD values.
- Approved desktop run `20261125T184421Z-86040-8c95817a3fe74e9ca49a67973b5674a6`
  passed `cmd_pane_fileops_retained_source_and_mixed_journeys` in 3,010 ms
  with zero failures and skips. The owned native process ran for 8.334 s.
  The launch closed its contained processes, released the desktop lease
  and archived the evidence.
- The same receipt passed setup, the Copy-only Move case and teardown in
  non-activating run `20261125T184631Z-504-bbc14ae3566d4b679f3d9c8af84569b5`.
  That case kept the source after the destination acquired content proof.
  The two focused repository verdicts were `NOT_EVALUATED`.

Mixed endpoint plan qualification remains open.
