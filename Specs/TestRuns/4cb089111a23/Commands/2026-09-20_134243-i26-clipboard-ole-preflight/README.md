# Clipboard OLE persistence preflight

**6/6 passed**, 0 failed, 0 skipped. Preserve each result/reason; this is focused Debug x64 evidence, not Fresh Full, Release or redesign completion.

Raw evidence: `S:/RedSalamander.Perf/evidence/i26-clipboard-png-ole-preflight-20260920`. Source snapshot `552a4d9939e4bd70987c8170ad173123b31a6eb193a799db36532f0a0f2f7cca`; build receipt `0693ac4571ed0af529cb5eeb6d90f9109535c072b689c40efc6241bf48d7f6b4`. The existing application test harness owns all application screenshots and directed keyboard interaction, with its desktop lease, warning and focus/cursor restoration. No user clipboard contents are logged.

The baseline reproduces stream-only rich-format OLE flush/readback loss. The diagnostic isolates round zero (`CLIPBRD_E_BAD_DATA`, 0x800401D3), with a 12-byte stream at position zero. An owned-data-only snapshot previously passed: that did not establish native clipboard persistence. A private window-station probe was denied creation (error 5) before clipboard access. The user subsequently copied text and explicitly authorized these directed runs.

The candidate serializes stream bytes to owned HGLOBAL, bounds snapshot bytes to 64 MiB and entries to 256, preserves the source cursor through a clone, checks exact values after each OLE cycle, and verifies every saved format after user-clipboard restoration. Unsupported carriers, device-specific or indexed descriptors fail before mutation rather than being dropped. This preserves admitted byte contents, not arbitrary delayed-render/provider behavior. A deliberately corrupted expected byte must fail comparison. RestoreReadback metrics record format counts/results only. StreamSnapshot records bytes/position only.

API context: [OleFlushClipboard](https://learn.microsoft.com/en-us/windows/win32/api/ole2/nf-ole2-oleflushclipboard) and [OleGetClipboard](https://learn.microsoft.com/en-us/windows/win32/api/ole2/nf-ole2-olegetclipboard) describe persistent native formats and reconstruction. The Shell/Windows stream failure here is measured behavior, not a universal claim about all IDataObject implementations.

This final review adds verification of all retained values before clipboard replacement, guarding against Shell format-key coalescing. A failed restore metric reports zero fully verified formats. All six affected cases pass; the 36-case matrix was executed before this final test-only guard.

The curated runner transcript removes trailing spaces/blank EOF lines only; the original full log remains under `C:/RedSalamander.Perf/evidence/i26-ui`. Result JSON, trace and raw metrics are not rewritten.
