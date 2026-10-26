# Clipboard OLE persistence baseline

**0/3 passed**, 3 failed, 0 skipped. Preserve each result/reason; this is focused Debug x64 evidence, not Fresh Full, Release or redesign completion.

Raw evidence: `S:/RedSalamander.Perf/evidence/i26-clipboard-png-ole-baseline-20260920`. Source snapshot `19f73280c8c151d0f0621c0352a671f5fb4882c58a909902c7fe572c42b3ccb7`; build receipt `21bfa65979792124a1f239dfc08454dab68c709924248d30ad4ed12a232c84b8`. The existing application test harness owns all application screenshots and directed keyboard interaction, with its desktop lease, warning and focus/cursor restoration. No user clipboard contents are logged.

The baseline reproduces stream-only rich-format OLE flush/readback loss. The diagnostic isolates round zero (`CLIPBRD_E_BAD_DATA`, 0x800401D3), with a 12-byte stream at position zero. An owned-data-only snapshot previously passed: that did not establish native clipboard persistence. A private window-station probe was denied creation (error 5) before clipboard access. The user subsequently copied text and explicitly authorized these directed runs.

The candidate serializes stream bytes to owned HGLOBAL, bounds snapshot bytes to 64 MiB and entries to 256, preserves the source cursor through a clone, checks exact values after each OLE cycle, and verifies every saved format after user-clipboard restoration. Unsupported carriers, device-specific or indexed descriptors fail before mutation rather than being dropped. This preserves admitted byte contents, not arbitrary delayed-render/provider behavior. A deliberately corrupted expected byte must fail comparison. RestoreReadback metrics record format counts/results only. StreamSnapshot records bytes/position only.

API context: [OleFlushClipboard](https://learn.microsoft.com/en-us/windows/win32/api/ole2/nf-ole2-oleflushclipboard) and [OleGetClipboard](https://learn.microsoft.com/en-us/windows/win32/api/ole2/nf-ole2-olegetclipboard) describe persistent native formats and reconstruction. The Shell/Windows stream failure here is measured behavior, not a universal claim about all IDataObject implementations.

The curated runner transcript removes trailing spaces/blank EOF lines only; the original full log remains under `C:/RedSalamander.Perf/evidence/i26-ui`. Result JSON, trace and raw metrics are not rewritten.
