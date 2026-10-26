# Owned PNG stream snapshot characterization

**3/3 passed**, zero skips, compiler warnings/errors and sandbox audit issues. This is passing characterization of the existing helper, not evidence that the live PNG restoration failure is repaired.

`cmd_pane_fileops_clipboard_snapshot_medium_contract` now stores an opaque PNG-format byte fixture through IStream alongside its HGLOBAL HTML fixture. It snapshots the data object, invalidates the original stream backing, and requires exact bytes after each of three retained snapshots. It leaves the returned stream at EOF before the next snapshot. This run uses owned data objects only: `clipboardRoundTrip=false`, no user clipboard mutation.

The existing Shell-backed retainer survives the invalidation. Therefore simple shared-backing retention has not reproduced the 34/36 broader run's `CLIPBRD_E_BAD_DATA` for advertised PNG format 49343; do not apply an unproven duplication workaround. The missing boundary remains actual OLE flush/readback and original-format restoration. No unchanged directed clipboard retry or clearing is performed. A future isolated owned-clipboard scenario must characterize that boundary before further live mutation.

API context: [IStream::Clone](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-istream-clone) itself shares storage; this test checks the complete data-object snapshot behavior rather than assuming that interface identity proves independence. [IStream::CopyTo](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-istream-copyto) changes stream positions; any future explicit copy must preserve the original cursor and validate exact counts.

Source snapshot `84a4e0da6a964a8b63987136d56f204e2c94db7119e26feda132e94d291f90d8`, receipt `20ab0644e4916cea47443bf0ddc7fa2ab7d5bf83bbf98b59d25a8a88aee9c17d`. Raw evidence: `S:/RedSalamander.Perf/evidence/i26-clipboard-png-characterization-20260920`. No product renderer, library or dependency change accompanies this test.
