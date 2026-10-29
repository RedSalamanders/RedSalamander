# Captured-console and writer repetition

Candidate at 7a3a7057669c5cd9d2ae5c629133f6cdebd23875 plus the retained native
input patch, x64 ASan Debug Embedded. Run 20261029T071701Z-63824-12778437662840b2b106a8991cdb9d2e:
30 passed, zero failed/skipped/retries, wrapper exit zero. Each of the original
runner contract, duplicate-instance rejection and resource-scoped writer cases
runs ten times. Native duration: 28518 ms.

All ten console probes report N, all ten second instances exit 1 while the primary
remains alive, and the original output/ownership/deadline assertions remain active.
The trace retains every repetition; last-duplicate-output.txt is the final raw child
output. Build receipt and native input patch identify the uncommitted candidate.
The pending changelog fragment is not a native compiler input. Final PDBs and ASan
instrumentation remain enabled as applicable. This is focused local acceptance;
broader and hosted/native ARM64 qualification remain open. A passing isolated
duplicate retry before the flag repair is retained separately and is not a causal fix.
