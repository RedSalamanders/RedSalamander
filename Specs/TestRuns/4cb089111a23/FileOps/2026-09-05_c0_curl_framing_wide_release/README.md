# Curl framing: broad Release C0 failure retained

Native/runner: 15 passed, 1 failed, 1 not reached; suite 80,984 ms. Copy's 46
scenarios run, but only 131/142 checks pass. Both wide copies publish all 4,097
files with exact contents: parallel 7,800,994 us, serial 23,156,427 us. The broad
run is nevertheless FAILED; the later lookup case is not reached.

The failures are later deep-tree scenarios, not wide throughput:

- Depth 128: all four MoveItem/MoveItems serial/parallel cases fail expected
  success, contents and completion oracles. First returns 0x8007012B, the other
  three return 0x800704C9 (connection refused).
- Depth 129: CopyItem parallel and both CopyItems modes return 0x800704C9 instead
  of required 0x8007006F (traversal bound). CopyItem serial reports the expected bound.

These seven scenario rows account for eleven failed checks. Server status remains
S_OK; that does not establish a healthy client transport. Completion-oracle failure
against expected successful Move is not by itself proof of a falsely reported success.

A subsequent [exact Copy run](../2026-09-05_c0_curl_framing_copy_isolated_release/README.md)
passes 142/142 on the same receipt, without a source change. This leaves an
order-sensitive/transient integration failure requiring diagnosis and repair;
no successful retry replaces this red evidence. The runner's original
UNCLASSIFIED_FAILURE classification is retained unchanged.

The 6,318,103-byte, 17,789-row raw stream remains in the governed run. All 1,176
claim-bearing Curl/C0 rows are preserved in order with the digest-bound selection.

## Reproduction and provenance

Run `c0-curl-framing-wide-release-20260905` on machine `4cb089111a23`.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_ -RunId c0-curl-framing-wide-release-20260905
```

Use a new run ID when repeating. `build-provenance.json` and the unchanged entry
fingerprints bind source/build identities; `file-provenance.json` records raw hashes.
Results and traces are unchanged. If present, `build.txt` only normalizes trailing
horizontal whitespace and UTF-8 encoding; both log hashes are recorded and the raw
log remains under `.build/logs`. Raw runtime data remain in this exact run beneath
`C:\RedSalamander.Perf`. No retained data were deleted to clear disk-audit findings.

See the [step overview](../2026-09-05_c0_curl_framing_lookup_debug/README.md).
This is focused evidence, not final Fresh Full or plan retirement.
