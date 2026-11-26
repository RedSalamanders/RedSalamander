# I66 Unpack worker and archive dialogs

Stored and deflated ZIP extraction use the File Operations worker. The exact archive case holds each worker immediately before an entry commits, after an exact fixture cleanup approval. The pane retains control. Cancel preserves the ZIP and previous destination, retires the unpublished stage by handle and starts no cleanup Delete. The round trip still verifies overwrite and skip, exact cleanup, and the supported and rejected archive contents.

Both full test-enabled builds at `41211242ce6016c1fba02c339a96055ebe0f05b4` had zero warnings and errors. The Debug receipt was `ca9dd4242953e2b6aad045ff10c3bc04f86a9e53c0d0c31a74fc670b06754e6f`. Every Release entry verified `b193bc1ec813c48455998710934241be21527003bb5290a3e83e08145a55f119`.

| Profile and selection | Governed run | Passed | Failed | Skipped |
| --- | --- | ---: | ---: | ---: |
| x64 Debug archive commands | `20261125T211028Z-91548-88b6e7893a5c40a9a9e97844a8ae03e4` | 1 | 0 | 0 |
| x64 Release archive commands | `20261125T212624Z-98148-fe523f41bf36496e9cc5b52b299c59f8` | 1 | 0 | 0 |
| x64 Release Pack dialog | `20261125T212424Z-65756-b146d49a9ab740de8f08af10d3da60b1` | 1 | 0 | 0 |
| x64 Release Unpack dialog | `20261125T212520Z-84068-e804bd785ec04b17b19b1d5c26c34f14` | 1 | 0 | 0 |
| x64 Release deferred exit | `20261125T212749Z-27324-0854ddf96b184e2f8811b909b118e1e8` | 3 | 0 | 0 |

The dialog pass used the one approved focus window `20261125T212422Z-88992-a883a57548514ebe90d20aa6ac063f3c`. The bounded launcher finished both selections in 102.937 seconds. Both runners closed their processes and released the desktop. The archive command and deferred-exit cases used no desktop activation.

For the first successful stored ZIP on this machine, the [Pack worker baseline](../2026-11-25_220215_pack_worker/README.md) still extracted on the pane thread. The same fixture unpacked 16 entries and 23 bytes. The following individual timings are in microseconds.

| Profile | Baseline pane execution | Candidate pane admission | Candidate worker execution |
| --- | ---: | ---: | ---: |
| x64 Debug | 49,300 | 404 | 45,676 |
| x64 Release | 46,570 | 112 | 59,490 |

The candidate admission and execution metrics came from distinct threads. These focused runs reported repository coverage `NOT_EVALUATED`.
