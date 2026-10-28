# Continuation Evidence: Compare Compact Preflight Pipe Missing

Captured after the foreground stdout request-budget fix under the required `C:\RSPerf` root. This is not final proof.

Result:
- CompareDirectories: 218 passed / 1 failed / 30 skipped
- Disk audit: 0 unexpected entries

Failure:
- `search_service_compact_request_roundtrip`
- Reason: timed out waiting for `SearchServiceBroker::GetStatus` for `live compact request preflight service`; last HRESULT `0x80070002`.

Working theory:
- The foreground service can expose an initial pipe instance and then exit or fail to recreate the next instance before the compact preflight status request in broad order. The fix must keep the compact request deterministic and must not restore the hidden readiness status request or quarantine the test.
