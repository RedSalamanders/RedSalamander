# File Operations P1.3 focused failure evidence

- Time: 2026-08-22 16:32 Europe/Paris
- Build attestation: `4516a88c230540e0930cfdfd58e8ce4b01b0607b6bac08af7aebe92dd872be66`
- Repro: `.build\\x64\\Debug\\RedSalamander.exe --fileops-selftest --selftest-case=FileOps_ProviderCapabilityMatrix`
- Process result: clean exit `1`; no new crash artifact and no retained process.
- Native evidence source before rerun: `%LOCALAPPDATA%\\RedSalamander\\SelfTest\\last_run\\fileops\\trace.txt` and `results.json`.

Failure:

```text
[7080234] FAIL: Host typed admission should reject 7z copy before task creation
(hr=0x80070057 before=0 after=0).
```

Diagnosis: the deterministic 7-Zip admission fixture supplied `state.fs7z` but omitted its source
plugin/profile overrides, so the typed plan inherited the current local-pane plugin ID. The new
local device-envelope validator correctly rejected `/missing.txt` as an invalid local path before
the intended 7-Zip capability rejection. The fixture must qualify the request as 7-Zip; production
admission behavior is not relaxed.

## Qualified-fixture rerun

- Build attestation: `123e25fccb7ceb0ab21af3f0fba04c2f5a35c288aa4a33fd2d7fe1bcf8e8e2e1`
- Process result: clean exit `1`; no new crash artifact and no retained process.

```text
FAIL: Host typed admission should reject 7z copy before task creation
(hr=0x00000000 before=0 after=1).
```

Diagnosis: once the fixture supplied the correct 7-Zip endpoint, `AdmitOperation` used the
real-path capability document to construct the typed endpoint but did not enforce that document's
operation booleans before publishing the task. The fix validates every selected source document
and the destination document (or the cross-provider export/import pair) inside the plan factory.
This remains a top-level admission check; it does not enumerate recursive children.
