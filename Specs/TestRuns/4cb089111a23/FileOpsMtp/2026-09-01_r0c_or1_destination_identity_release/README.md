# R0c-OR1 MTP destination identity replay

- Machine profile: `4cb089111a23`
- Source commit: `faa4abd47eab02cb97171783daa9761274206b91`
- Build: test-enabled x64 Release, `MaxCpuCount=1`, RedSalamander project receipt `30f307d6773fd6ee6cf4292c76b59f9cd37f214610a88be247e05e752a336f88`
- Command: `RedSalamander.exe --compare-selftest --selftest-no-activate --selftest-fail-fast --selftest-case=mtp_overwrite_journal_destination_identity_mismatch_quarantines`
- Scenario: an identified overwrite temp contains the only replacement bytes while a different-PUID object occupies the recorded destination path
- Result: `passed=1, failed=0, skipped=0`; whole case `24 ms`
- `FileOps.Mtp.R0cOr1.DestinationIdentity.SelfTestUs`: `23,492 us`
- Parent enumeration: `10` entries; one `mtp.overwrite.journal_replay_destination_identity_rejected` event
- Resource shape: one backend enumeration vector plus constant temp/destination identity and count state

Replay preserved both byte payloads, quarantined the journal, and performed no device mutation. The focused metric is below the `2,000,000 us` deterministic ceiling. This is fake-backend safety/resource evidence and does not claim live-WPD latency.

Exact-source x64 Debug full-solution build receipt is `642bf5ecd17da1cb4aec532bdf7245369ef4c27380ad1368968f9bcf35e39a6b`. The complete MTP family passed `54/0/1` in both Debug and test-enabled Release; the one skip is the declared physical-device smoke because no approved device was configured.
