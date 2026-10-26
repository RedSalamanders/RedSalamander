# Separate cursor and blocking handles — focused pass

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlEntryLookupTruth -FailFast -RunId c0-curl-split-handles-debug-20260905`.
Based on `3dab443c`; pending pool/cursor/test changes. Build 03:45, zero warnings/errors.
Receipt `69086f9e7568b574c6db494e08e9f08fdeed384dfee1eee9d920318a1952c40b`;
workspace `00de4818d995672d276fa3a3a66458c7735d2159074102bd2484bd0f96d8071a`.

**201/201 assertions; 3/3 harness cases, zero skips.** Pure 16-lookup Unix/DOS
controls use one authenticated session; alternating 16 LIST/SIZE controls use
exactly two sessions with all 16 sizes correct. All four use one passive listener
and 15 reuses. Cancellation remains 55,707–59,558 us, with a successful subsequent
lookup in each of eight routes. Complete rows, requested metadata, conservative
mutation receipts and zero unsafe mutations remain checked.

Each existing pool entry now retains its own multi cache and distinct cursor and
blocking easy handles. Cursors detach; both easy handles reset callbacks on return.
All owners remain under the existing four-entry/60-second lazy-expiry policy.
No global connection-sharing pool, retry policy or transfer scheduler was added.

Pure Unix/DOS probes take 399,513/361,741 us, versus 267,142/261,366 us in the
preceding single-easy pool candidate on this machine. Mixed probes take
407,711/337,218 us and have no like-for-like old timing baseline. Do not claim a
speed improvement from these samples. The measured gain is mixed-operation
connection reuse; opaque transport memory and broad throughput remain separate.

Full perf: 150,452 bytes, 408 rows, SHA-256
`81238abac9a2d564c5e321227f40e3972b369b2ec149ac6774acfdd4081afb48`.
Case 8,297 ms; suite 9,797 ms. Disk audit 82 issues, none deleted.
Source/localization 197/197 (17.77 s), spec inventory zero blockers.
The [same-source broad run](../2026-09-05_c0_curl_split_handles_wide_red/README.md)
still fails; this is not C0/Release/Done qualification.
