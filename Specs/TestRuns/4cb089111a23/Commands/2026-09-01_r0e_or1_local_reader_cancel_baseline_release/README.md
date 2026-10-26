# R0e-OR1 Local reader cancellation RED baseline

- Machine profile: `4cb089111a23`.
- Exact RED commit: `8dd8f305eb319ae3808550409eb186f47158ae2c`.
- Build: test-enabled x64 Release, RedSalamander project receipt
  `63dd70a85845342fd4c508dc2df0a229570ed342c4e55ce0e3cbbc29050e0ec2`;
  diagnostics `0 warning(s), 0 error(s)`.
- Command: five independent processes running
  `RedSalamander.exe --commands-selftest
  --selftest-case=cmd_pane_fileops_local_blocked_reader_cancel_is_bounded`.
- Scenario: read the same deterministic 8-MiB fixed-local payload through the legacy
  and exact-bound Local readers, verify seek replay, then attempt a real Local reader
  read from a connected outbound byte pipe whose server produces no bytes.
- `FileOps.SelfTest.R0eOr1.OrdinaryReadUs`: `7,688`, `6,437`, `9,861`,
  `6,865`, and `7,724` us; nearest-rank p95 `9,861` us.
- `FileOps.SelfTest.R0eOr1.LocalBlockedReadCancelUs`: `528,468`,
  `527,050`, `527,005`, `528,759`, and `528,449` us.

All five RED processes fail the intended assertion (`passed=0, failed=1`): the
reader performs only its initial operation-control check, does not remain as an
observable cancellable kernel-pending request, and returns `0x8007006D`
(`ERROR_BROKEN_PIPE`) only after the 500-ms fail-safe closes the server. Every
process still reports zero bytes after Cancel. The ordinary control is stable in
all samples: exactly `16,777,216` combined bytes and successful seek replay.

The archived metric files are bound as `(rows, bytes, SHA-256)`:

- sample01: `(318, 119671, 7F8CE336E9D7FC21BBAF996A6B3CC52E4F7B2831D90399CE00B0A69F7ED8EB04)`
- sample02: `(318, 120227, 1702EABE3078E993EB25C49B22DA383FAE7867B0F851C2769894B2FF255BA67D)`
- sample03: `(318, 119923, 4A853DC1BC82B5397AC847E95AC96722D628F247ED4AD3B2417EC58C121A2C2A)`
- sample04: `(318, 120195, 3A39A1B66828394D5576338AC571EFB9ABF6F6B17A0FFC827C9CE4F24AE3AA62)`
- sample05: `(318, 120191, 1ED2775AC32DD516834D0EEF158A4013D14755AF5226BD65DEC2D873116EDE3E)`

This RED archive proves the missing pending-read cancellation behavior while
retaining a bounded fixture. It is not disk-throughput, writer-cancellation,
remote-provider, or generic UI-shutdown evidence.
