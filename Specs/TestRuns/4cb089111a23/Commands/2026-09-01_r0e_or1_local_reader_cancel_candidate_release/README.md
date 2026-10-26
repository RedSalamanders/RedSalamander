# R0e-OR1 Local reader cancellation candidate

- Machine profile: `4cb089111a23`.
- Exact implementation/spec commit: `d7decb405f088a3ad95180a028e8f31f1324ae23`.
- Build: test-enabled x64 Release, RedSalamander project receipt
  `7c099b7742ac87d76d891eaf86ad272e4c7174ed402f1352005a46c85265d30f`;
  diagnostics `0 warning(s), 0 error(s)`.
- Command: five independent processes running
  `RedSalamander.exe --commands-selftest
  --selftest-case=cmd_pane_fileops_local_blocked_reader_cancel_is_bounded`.
- Scenario: the same deterministic ordinary-read/seek control as the RED archive,
  followed by a real Local overlapped reader blocked on a connected outbound byte
  pipe whose server produces no bytes.
- `FileOps.SelfTest.R0eOr1.LocalBlockedReadCancelUs`: `30,943`, `46,574`,
  `31,010`, `31,025`, and `31,628` us; nearest-rank p95 `46,574` us.
- `FileOps.SelfTest.R0eOr1.OrdinaryReadUs`: `6,878`, `5,728`, `5,500`,
  `6,950`, and `6,810` us; nearest-rank p95 `6,950` us.
- Ordinary-read ceiling: `59,861` us
  (`max(1.50 * 9,861, 9,861 + 50,000)`); result PASS.

All five processes pass (`passed=1, failed=0`). Each blocked read enters the
pending state, performs three operation-control checks, returns `0x800704C7`
(`ERROR_CANCELLED`), reports zero bytes, and cleans up before the 500,000-us
absolute/p95 bound. Every ordinary control reports exactly `16,777,216` combined
bytes and successful seek replay. Candidate ordinary-read p95 is below both its
RED baseline p95 and the accepted regression ceiling.

The archived metric files are bound as `(rows, bytes, SHA-256)`:

- sample01: `(318, 119954, 2B241BC1C5E850E5147ADA67477BDB4092B07706E360A9628134F9CD26BCDA69)`
- sample02: `(318, 119889, 37CA9B7BC058502E264D10066DD424C3ECBC3F7559EBB397868CE46DE18E3B68)`
- sample03: `(318, 119888, E43B88E36110C5D3EBC0515414F11AA7BD277D9256A3FAD4EEA0876C2896BB9A)`
- sample04: `(318, 119898, 4C02FB67AD99D93F4D73720478D8F9ECF5CD8076BCD13B66AC5CD0BDF2EDBA53)`
- sample05: `(318, 120185, F508312CED7613168F2DD6AA484FE6901500EF4C941B52FC2B4AD525054AC2BC)`

This evidence qualifies only admitted fixed-volume Local regular-file source
reads through the exact option-bearing reader. It does not add a public ABI,
qualify writers or remote providers, establish a generic host join deadline, or
re-enable any route classified `uncontained`.
