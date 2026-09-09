> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# Data-Integrity Audit — RedSalamander File-System Subsystem

## 1. Verdict

**The subsystem is structurally unsafe for unknown-size streaming transfers, cross-filesystem moves, and in-place overwrites — and the unsafety is silent, reported as success, and concentrated in the exact paths users trust least to be recoverable (cloud/remote backends and overwrite-on-collision).** The single most dangerous theme is **"trust a 0-byte read as EOF when the size is unknown"**: it recurs across the cross-FS bridge COPY path, the 7z spool reader, the S3 ranged reader, the Curl download relay, and the under-specified `IFileReader::Read` contract that licenses all of them — every instance turns a transient short read into a silently truncated file presented as a green success, and for MOVE the original is then deleted. The second theme is **destructive-before-committed**: same-FS overwrite copies and the dummy/MS-Drive overwrite paths truncate or delete the original target before the replacement is durable, so any interruption is net data loss. The third is **delete-without-reverification**: cross-FS directory MOVE and S3/MS-Drive recursive deletes act on a stale enumeration snapshot, destroying source children that were created during the copy window or skipping/over-deleting concurrently-changed trees. Compounding all of this: the LocalSystem search service runs with **no client impersonation** (metadata-disclosure security hole), the conflict-resolution cache **leaks an Overwrite grant from a plain file onto junction/symlink replacement** (catastrophic silent link destruction), and traversal **follows reparse points with no cycle guard** (stack-overflow crash). The good news: the same-FS MOVE path, the cross-FS MOVE size verification, the batch copy/move work-stack, and the multipart S3 path already demonstrate the correct patterns — the fixes are largely about making the unsafe paths symmetric with the safe ones that already exist in the same files.

---

## 2. Findings by severity

### SEVERITY: DATA-LOSS & CORRUPTION

---

**1. Unknown-size cross-FS COPY trusts a single 0-byte read as EOF and promotes a truncated file as success**
`RedSalamander/FolderWindow.FileOperations.State.cpp:7474` (also 7090-7104, 7188, 7312/7383, 7533) — **data-loss / CONFIRMED**
*(Merged: this is the same root-cause/line reported under five unit keys — bridge-crossfs-copy, bridge-move-delete, bridge-samefs-overwrite, bridge-dir-tree, bridge-cancel. Affects every cross-filesystem single-file COPY.)*

When `reader->GetSize()` fails, MOVE bails out with `ERROR_PARTIAL_COPY` and preserves the source (7094-7104), but COPY deliberately continues with `hasKnownFileTotalBytes=false`. Both copy loops (serial 7188, pipelined 7312/7383) treat any `S_OK` read with `bytesRead==0` as definitive EOF. The only completeness guard (`fileCompletedBytes != fileTotalBytes`, 7474) and the post-promote destination re-stat (7533) are **both gated on `hasKnownFileTotalBytes`** (and the re-stat is additionally MOVE-only), so an unknown-size COPY has **zero end-to-end verification**. The truncated temp is `Commit()`ed (7501) and atomically promoted over the destination (7518), and the task returns `S_OK`.

**Failure scenario:** COPY an S3/Curl/cloud object whose `GetSize`/HEAD fails or is denied (HEAD-vs-GET permission split, throttle, network blip). A mid-stream ranged GET returns an empty/short body — `S3RangedFileReader::Read` returns `S_OK, bytesRead=0` (`FileSystemS3.IO.cpp:411-413`, latched at 549-553). The bridge treats it as EOF, promotes the partial over any existing destination, reports success. If the remote was the only copy and the user trusts the green result, the original may later be deleted.

**Deeper fix:** Make COPY symmetric with MOVE. After the loop, when size is unknown, re-stat **both** endpoints (re-open a fresh source reader for `GetSize`, re-stat the committed destination) and require `destinationFinal == bytesStreamed` before promoting; on inability to prove it, return `ERROR_PARTIAL_COPY` and never promote over existing data. Root-cause: tighten `IFileReader::Read` to forbid premature `bytesRead==0` (see #20), and fix `S3RangedFileReader` to distinguish genuine completion (`EnsureSizeKnown` succeeded and `_position>=_sizeBytes`) from an empty/short range response (retryable error, not EOF).

---

**2. Same-FS COPY overwrites the destination in place via `CopyFileExW`; an interrupted overwrite destroys the original and leaves a partial**
`Plugins/FileSystem/FileSystem.FileOps.cpp:3588` (reparse variant: 3722) — **data-loss / CONFIRMED**
*(Merged: two findings at this line, unit keys bridge-samefs-overwrite and fsplugin-copy, describe the identical defect.)*

With overwrite granted, `copyFlags=0` (3587) and `CopyFileExW` opens the live destination with `CREATE_ALWAYS` (truncating) semantics — original bytes are destroyed the instant writing begins. On interruption (Cancel → `PROGRESS_CANCEL`, disk-full, USB/network drop, power loss) the failure branch calls `ShouldRollbackCopiedDestination(destinationExisted=true, false)` which returns **false** by design (2762-2765), so no rollback runs and the truncated partial is left in place. The op returns a clean failure HRESULT while the user's original file is permanently gone. The reparse/symlink copy at 3722 has the identical defect. This diverges from the cross-FS bridge (temp-stage + atomic promote) and from same-volume MOVE (`MOVEFILE_REPLACE_EXISTING`, atomic).

**Failure scenario:** Copy an 8 GB file over a different existing 8 GB file on the same volume, choose Overwrite; at ~50% the user clicks Cancel (or the volume fills). The original 8 GB is destroyed and replaced by a partial; routine-cancel data loss.

**Deeper fix:** Never write onto the live destination. Copy to a `CREATE_NEW` temp sibling (as `MakeTempDestinationPath` does), flush/verify byte count == source size, then atomically replace via `ReplaceFileW` (preserves ACLs/streams) or `MoveFileExW(MOVEFILE_REPLACE_EXISTING)`. On any failure/cancel before the rename, delete the temp and leave the original untouched.

---

**3. `CreateFileWriter(ALLOW_OVERWRITE)` opens `CREATE_ALWAYS` in place; abort path then `DeleteFileW`s the (already-truncated) original**
`Plugins/FileSystem/FileSystem.cpp:1953` (destructor 1782-1794) — **data-loss / PLAUSIBLE**

With `FILESYSTEM_FLAG_ALLOW_OVERWRITE`, the writer opens the **final** destination with `CREATE_ALWAYS` (truncates at open) and offers no staging. `IFileWriter`'s contract says Release-without-Commit is an abort, and `~Win32FileWriter` honors that by `DeleteFileW(_path)` when uncommitted — so an interrupted overwrite both truncates the real file and then deletes it outright. *No in-tree caller currently triggers this* (the bridge masks ALLOW_OVERWRITE and stages to a temp; other callers pass `NONE`→`CREATE_NEW`), but the flag is public, documented as a supported in-place mode, so any external/future caller using it as documented loses data.

**Deeper fix:** Make the writer's overwrite path crash-safe internally — stage to a sibling temp, atomic `ReplaceFileW`/`MoveFileEx(MOVEFILE_REPLACE_EXISTING)` on `Commit()`, abort deletes only the temp; never truncate the user's file until the replacement is flushed. Alternatively refuse `ALLOW_OVERWRITE` and document temp+rename as the caller's responsibility.

---

**4. Cross-FS directory MOVE blind-deletes the source root from a stale enumeration snapshot (TOCTOU); files created during the copy window are deleted without ever being copied**
`RedSalamander/FolderWindow.FileOperations.State.cpp:8535` (serial 9134; gate 8262-8277; enumeration 8001-8132) — **data-loss / CONFIRMED**
*(Merged: three findings, unit keys bridge-move-delete, bridge-dir-tree — same blind recursive-delete root cause.)*

After `bridge.CopyPath()` returns `SUCCEEDED`, the bridge issues an **unbounded** `DeleteItem(sourceText, itemFlags|FILESYSTEM_FLAG_RECURSIVE)` on the source root. The gate `ShouldDeleteMoveSourceAfterBridgeCopy` checks only the copy HRESULT plus reparse/conflict skip counters — never whether each enumerated source entry has a verified destination counterpart. The recursive delete re-walks the **live** tree at delete time, not the materialized set captured at enumeration (8001-8132). Any child created (or appended to) in the source between enumeration and delete is destroyed without being copied. This is in sharp contrast to the same-FS MOVE path (`DeleteCopiedSourceEntryForMove`/`DestinationMatchesSourceFile`, `FileSystem.FileOps.cpp:3241/3276`), which deletes a child bottom-up only after re-confirming its destination counterpart matches, and ends `ERROR_PARTIAL_COPY` (source preserved) otherwise.

**Failure scenario:** Cross-volume MOVE of a download/log/appdata directory another process is actively writing into. The bridge copies the enumerated set, returns `S_OK`; an external write lands a new file in an already-enumerated subdir; the recursive delete removes it. Permanent loss of just-created data, reported as a successful move.

**Deeper fix:** Mirror the same-FS MOVE contract. Re-enumerate the source at delete time and delete bottom-up, per entry, only after re-confirming a destination counterpart with matching size/content; preserve any uncopied/extra child and end the move `ERROR_PARTIAL_COPY`. Do not trust the copy-time snapshot.

---

**5. Move-source delete is identity-bound, not image-bound: data written to the source between the byte-compare and the unlink is lost**
`Plugins/FileSystem/FileSystem.FileOps.cpp:3385` (compare 3241-3268; delete handle 3288) — **data-loss / PLAUSIBLE**

`DeleteCopiedSourceEntryForMove` opens the source for DELETE (3288), then `DestinationMatchesSourceFile` opens a **separate** read handle, confirms identity, and byte-compares. On match it unlinks via the original DELETE handle (3385). Both opens grant `FILE_SHARE_WRITE`, so a concurrent writer (log/DB) can append between `FileContentsEqual` returning true and `MarkHandleForDeletion` executing. The delete is bound to the file **identity**, not the verified byte image, so newly-written bytes the destination never received are destroyed — while the move reports success for that entry. The function's own comments claim copy-window content is preserved; that guarantee does not extend to the post-compare/pre-disposition window.

**Deeper fix:** Hold the delete to the exact verified image — compare through the **same** handle that deletes (open with `FILE_READ_DATA` + oplock/lease, or take a delete-pending disposition before comparing), or re-verify size+last-write-time immediately before `MarkHandleForDeletion` and skip (preserve source, `ERROR_PARTIAL_COPY`) on any change.

---

**6. Cross-volume MOVE destination verification checks only file SIZE, not content, before the irreversible source delete**
`RedSalamander/FolderWindow.FileOperations.State.cpp:7546` (7533-7563) — **corruption / PLAUSIBLE**

The post-promote MOVE check re-opens the destination and compares only `destinationSizeBytes != fileTotalBytes`. A same-size-but-corrupt destination (in-flight bit-flips on an unreliable backend, a sparse/logical-size-reporting backend, or a promote `MoveItem` that silently no-op'd over a stale same-size file) passes, the source is deleted, and the corruption is unrecoverable. No `hash|checksum|crc` exists anywhere in the bridge copy path.

**Deeper fix:** For the irreversible MOVE case, verify content not length — compute/compare a hash of the streamed bytes against a re-read of the destination. If a backend cannot support content verification, treat the MOVE as un-verifiable and preserve the source (`ERROR_PARTIAL_COPY`), consistent with the existing unknown-source-size policy at 7094.

---

**7. Read-only attribute restored onto a corrupted/truncated destination after a failed in-place overwrite**
`Plugins/FileSystem/FileSystem.FileOps.cpp:3509` — **corruption / CONFIRMED**

For a read-only existing destination with replace-readonly granted, READONLY is stripped (3554) and a `scope_exit` (3509-3515) re-applies the original attributes while `restoreDestinationReadonly` stays true. That flag is cleared only **after** `CopyFileExW` succeeds (3606). On a mid-copy failure (device removal, I/O error) the in-place truncation has already destroyed the original (see #2), the failure branch returns without clearing the flag, and the `scope_exit` re-stamps the original READONLY/hidden/system attributes onto the truncated stub — the garbage masquerades as an intact protected file.

**Deeper fix:** Tie attribute handling to the temp-file strategy (#2): apply attributes only to the verified temp before promotion; never restore attributes onto a destination truncated in place.

---

**8. `CreateFileWriter` strips READONLY then never restores it when the retry also fails**
`Plugins/FileSystem/FileSystem.cpp:1984` — **corruption / CONFIRMED**

On first-open `ERROR_ACCESS_DENIED`/`ERROR_SHARING_VIOLATION` against a read-only target, the code unconditionally clears READONLY (1984) and retries. If the retry also fails (another process holds the file open, ACL denial, lock unrelated to the read-only bit), the function returns the error at 1991 with the attribute **already permanently cleared** and nothing written. No `scope_exit` restores it. The user's read-only protection is silently and irrecoverably removed even though the op reported failure.

**Deeper fix:** Capture original attributes and install a `scope_exit` that re-applies READONLY unless the retry succeeds and the handle ownership transfers to a committing writer — make the attribute mutation transactional with the open.

---

**9. Reparse-following metadata copy stamps the LINK TARGET's timestamps/attributes instead of the link's — corrupting an out-of-tree directory**
`Plugins/FileSystem/FileSystem.FileOps.cpp:2682` (and 2715; call site 3941) — **corruption / CONFIRMED**

After `WriteReparsePointData` rebuilds a directory junction/symlink at the destination, `CopyPathBasicInformationBestEffort` opens both ends with `FILE_FLAG_BACKUP_SEMANTICS` but **without** `FILE_FLAG_OPEN_REPARSE_POINT` (2682/2715), so `CreateFileW` follows the link to its TARGET. It reads the source junction's target metadata and writes it onto **whatever the destination junction points at** — mutating an unrelated, possibly out-of-tree directory (e.g. `C:\Windows`). This file uses `FILE_FLAG_OPEN_REPARSE_POINT` everywhere else it touches reparse points (1392, 1429, 2789, 2942, 3148), confirming the omission is a defect.

**Deeper fix:** Add `FILE_FLAG_OPEN_REPARSE_POINT` (with `FILE_FLAG_BACKUP_SEMANTICS`) to both `CreateFileW` calls, or give `CopyReparsePointInternal` a no-follow metadata copy on the reparse handle it already holds.

---

**10. `\\?\` inputs bypass canonicalization — embedded `..`/trailing-dot components reach file ops literally, enabling wrong-target read/write/delete/move**
`Plugins/FileSystem/FileSystem.Path.cpp:58` (and `MakeAbsolutePath` 27-30; `MakePathInfo` 204-211) — **corruption / CONFIRMED**

Both `MakeAbsolutePath` and `ToExtendedPath` short-circuit and return any `\\?\`-prefixed input verbatim, never running it through `GetFullPathNameW`. The OS honors `\\?\` literally, so `..`/`.`/trailing-dot/trailing-space components survive into `info.extended` (used by `CopyFileExW`, `DeleteItem`, `MoveItem`, `CreateDirectoryW`, reparse ops) **and** into the `display`-derived reparse-root containment key (because `MakeAbsolutePath` also short-circuits), so `IsPathWithinRoot`'s literal prefix compare never collapses `..`. A `\\?\C:\data\..\secret\file.txt` op acts on a file outside the directory the containment check approved.

**Deeper fix:** On a `\\?\`(/`\\?\UNC\`) input, strip the prefix, fully canonicalize the remainder (collapse `.`/`..`, normalize trailing dots/spaces consistently with `display` derivation), then re-apply the prefix. Derive `display` and `extended` from one canonical form so they cannot diverge.

---

**11. Apply-to-all Overwrite consent leaks from a plain file collision onto junction/symlink replacement — silent catastrophic loss of the link target**
`RedSalamander/FolderWindow.FileOperations.State.cpp:2957` (buckets 2944-2962; cache 3128-3151; reuse 8706/9306, 8747/9347) — **data-loss / CONFIRMED**

`ClassifyConflictBucket` folds `ERROR_ALREADY_EXISTS`/`ERROR_FILE_EXISTS`, `ERROR_DIR_NOT_EMPTY`, and `ERROR_REPARSE_POINT_ENCOUNTERED` into one `ConflictBucket::Exists`. The decision cache is keyed **only** by bucket index — no path/type/error discriminator. A user who resolves a plain-file collision with Overwrite + Apply-to-all caches Overwrite for `Exists`; a later item whose destination is a junction/symlink classifies to the same bucket, the cached Overwrite is auto-applied with no dedicated prompt, and the engine `RemovePathForOverwrite` no-follow-deletes the link itself (engine 3993-4002/4485-4495 → `RemovePathForOverwrite` 3104-3106). The in-code comments at 2946-2956 explicitly say reparse/non-empty-dir replacement "require their own explicit consent" — the conflation defeats that. *(Note: the original finding's "deletes a non-empty real directory's contents" sub-claim is refuted — the operation-wide grant only `RemoveDirectoryW`s an empty dir; recursive removal needs the engine's per-path one-shot grant. The junction/symlink leak is the real catastrophe.)*

**Failure scenario:** Item 1 is `foo.txt` (exists) → Overwrite + Apply-to-all. Item N's destination `photos` is a junction to `C:\Users\me\OneDrive\Photos`. The cached Overwrite fires with no prompt; the engine deletes the junction and replaces it. Silent, irreversible loss of the link target the user never saw or consented to.

**Deeper fix:** Introduce distinct `ConflictBucket` values for reparse-replace and non-empty-dir-replace (independent cache slots, each demanding its own apply-to-all consent), or key the cached decision on `(bucket, destination-kind)` and force a fresh prompt whenever **this** item's concrete error is `ERROR_REPARSE_POINT_ENCOUNTERED`/`ERROR_DIR_NOT_EMPTY`. The reparse prompt must state that a link will be destroyed.

---

**12. Recursive DeleteNode destroys read-only descendants without `ALLOW_REPLACE_READONLY` (contract under-specification inherited by real plugins)**
`Plugins/FileSystemDummy/FileSystemDummy.cpp:4485` — **contract (data-loss-grade pattern) / CONFIRMED**

`DeleteNode` checks `FILE_ATTRIBUTE_READONLY` only on the top-level target; the recursive branch `ExtractChild`s the whole subtree in one move with no per-descendant re-check. A recursive delete with `RECURSIVE` but without `ALLOW_REPLACE_READONLY` silently removes read-only files buried in the tree. `FileSystem.h` `DeleteItem`/`DeleteItems` (449/478) carry **no** normative text requiring descendant read-only enforcement, so conformant production plugins inherit the gap. *(This is the in-memory reference plugin — no real user data here — but the header under-specification is the load-bearing defect.)*

**Deeper fix:** Walk the subtree on recursive delete, enforce read-only policy per descendant (raise `FileSystemIssue`/ReplaceReadOnly, or skip/fail when not granted), and document in `FileSystem.h` that delete MUST honor descendant `FILE_ATTRIBUTE_READONLY`.

---

**13. Single-item MoveItem/RenameItem aborts a directory merge on the first child collision AND leaves the source half-emptied (contract + data-loss)**
`Plugins/FileSystemDummy/FileSystemDummy.cpp:4284` (MoveNode 4432/4452; CopyItem variant 4247) — **data-loss / CONFIRMED**

`MergeMoveDirectory` moves children one at a time (each physically `ExtractChild`ed from source before being added to destination) and `return hr` on the **first** failed child (4285-4288). A mid-merge file-vs-file collision (`ERROR_ALREADY_EXISTS`, no overwrite flag) leaves children `0..N-1` already gone from the source while `b.txt`/`c.txt` remain — the function returns a hard error, the source is half-emptied, and `MoveItem` reports failure. This violates the `FileSystem.h:419-431` contract ("never abort a whole directory transfer on the first collision"; skipped children keep their source; end `ERROR_PARTIAL_COPY`). The batch `MoveItems` work-stack does honor the contract; the single-item path uses the abort-on-first-error helper. The same abort pattern exists in `MergeCopyDirectory` (#13b below, contract-only since Copy does not empty the source).

**Deeper fix:** Make the merge helpers per-child-continuation: raise each collision via `FileSystemIssue`, honor Skip/Overwrite/Retry, keep skipped children in source, finish `ERROR_PARTIAL_COPY`. Unify single-item Copy/Move/Rename onto the batch work-stack contract.

---

**14. MoveNode/CopyNode overwrite path destroys the existing destination before the replacement is committed; `AddChild` can throw after the destroy**
`Plugins/FileSystemDummy/FileSystemDummy.cpp:4452` (CopyNode 4358; CreateDirectoryClone 4196) — **data-loss / PLAUSIBLE**

`ExtractChild(&destinationParent, existing)` (4452) discards the return value, freeing the existing destination, then the source is extracted and a non-`noexcept` `AddChild` (3682) can throw on reallocation — leaving the destination gone with no replacement. *(Reference plugin; OOM-only trigger; the null-source-extract loss path is effectively unreachable.)* The destroy-before-commit ordering is the anti-pattern any real plugin copying this structure would inherit.

**Deeper fix:** Stage the replacement (insert new node / rename source into place) before destroying the original; only extract/free the previous destination after commit. Make `AddChild` `noexcept` (reserve up front) or restore the extracted `existing` on its failure.

---

**15. In-memory 7z spool `Read()` reports `S_OK`/0 bytes on a short extraction — silent truncation of an archive member**
`Plugins/FileSystem7z/FileSystem7z.cpp:2988` (and `EnsureExtractUntil` 3931) — **data-loss / CONFIRMED**
*(Merged: the `2988` and `3931` findings are the same short-extraction-as-EOF root cause.)*

In the in-memory spool branch, when 7-zip sets `SetOperationResult(kOK)` yet writes fewer bytes than the index-declared `kpidSize` (corrupted/inconsistent header, truncated member, size-field disagreement), `EnsureExtractUntil` exits its wait loop on `_extractFinished` with `_spooledBytes < endExclusive` and returns `S_OK` (3931). `Read` then sees `canTake==0` and returns `S_OK` with `*bytesRead=0` while `_positionBytes < _fileSizeBytes` (2988). The asymmetry is provable: the streaming pipe path (3644-3646) and the itemStream sub-path (4036-4038) correctly return `ERROR_HANDLE_EOF` on a drained-short read. Any copy/viewer/hash consumer treats `S_OK`/0 as clean EOF and silently truncates.

**Deeper fix:** Make the spool path symmetric — when `canTake==0` and `_positionBytes < _fileSizeBytes`, return `FAILED(_extractStatus) ? _extractStatus : ERROR_HANDLE_EOF`. Better: have `EnsureExtractUntil`/`EnsureSpooledUntil` return `ERROR_HANDLE_EOF` whenever extraction finishes with `_spooledBytes < endExclusive` and the declared size was non-zero.

---

**16. Curl download-to-temp never verifies size against the known source size; a short download (curl returns `CURLE_OK`) is promoted as the complete destination, then the move deletes the source**
`Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp:1211` (verify 1276-1298; download 3382-3384; move-delete 2392/3181) — **data-loss / PLAUSIBLE**
*(Merged: three findings — download truncation, SIZE-less-dialect existence-only verification, and the dir-merge amplifier — all flow from the same "verify the staged upload against the downloaded temp size, not against `expectedSizeBytes`" root cause.)*

`CopyFileViaTemp` takes `GetFileSizeBytes(tempFile)` as ground truth (1211) and verifies the staged upload **against itself** (1278/1293), never against the caller's `expectedSizeBytes` (used only for progress, 1151/1179). `CurlDownloadToFile` returns `S_OK` purely on `CURLE_OK` with no byte-count check. On SFTP/SCP/FTP dialects where a clean early stream-close yields `CURLE_OK`, or a SIZE-less FTP server where verification degrades to existence-only, a truncated transfer passes self-consistent verification, is promoted, and — for MOVE — the source is then `RemoteDeleteFile`d. The truncated copy is the only surviving copy. *(For a directory merge, a continue-on-error skip is correctly caught as `ERROR_PARTIAL_COPY` — the original "partial merge then delete" prong is refuted — but a silently-truncated child still amplifies into a bulk source-tree delete.)*

**Deeper fix:** Thread `expectedSizeBytes` into `CopyFileViaTemp`; require `GetFileSizeBytes(tempFile) == expectedSizeBytes` after download (fail `ERROR_PARTIAL_COPY`, delete temp, otherwise). Cross-check `CURLINFO_SIZE_DOWNLOAD_T` against `CURLOPT_INFILESIZE`. For MOVE on SIZE-less dialects, do not delete the source without a positive size/hash match — preserve and return `ERROR_PARTIAL_COPY`.

---

**17. S3 download-to-temp treats a short/interrupted body as success; relay copy+delete then loses the only copy**
`Plugins/FileSystemS3/FileSystemS3.S3.cpp:460` (relay 849-862; move-delete `Directory.cpp:1683`) — **data-loss / PLAUSIBLE**

`DownloadS3ObjectToTempFile`'s loop reads `while (stream.good())` and never compares bytes written to `result.GetContentLength()`, never distinguishing a clean EOF from `stream.bad()/fail()`. A mid-body network drop flips the AWS CRT IOStream to not-good without throwing → short temp file returned `S_OK`. On an S3→S3 MOVE where server-side `CopyObject` failed and `CopyS3ObjectWithFallback` falls back to download+upload, a <64 MiB object uses `PutS3ObjectFromHandle` with the original `ContentLength`, and after upload the source is deleted (`Directory.cpp:1683`). *(≥64 MiB uses the multipart path whose `ReadExactBytesFromFile` returns `ERROR_HANDLE_EOF` on a short file — guarded. The catastrophic loss is confined to the <64 MiB PutObject path and depends on whether the CRT commits a short-vs-ContentLength body.)*

**Deeper fix:** Capture `GetContentLength()` before the loop; after it, fail with `ERROR_HANDLE_EOF`/`ERROR_PARTIAL_COPY` unless bytes written == content length, and treat `stream.bad()/fail()` as error. Never return `S_OK` from a download whose byte count cannot be proven.

---

**18. `S3RangedFileReader` latches a short range-GET as EOF when size is unknown, permanently truncating the logical file**
`Plugins/FileSystemS3/FileSystemS3.IO.cpp:549` — **corruption / CONFIRMED**

When constructed without a prior `GetSize`/HeadObject (`_sizeKnown=false`), and a range GET returns `total < maxBytes` due to a transient short read (3 MiB of an 8 MiB range after a network hiccup), `FillBufferFrom` concludes EOF and bakes in `_sizeBytes = _bufferStart + total; _sizeKnown = true` (549-553), comparing against nothing in the GetObject response. All subsequent `Read`s short-circuit to clean EOF with no error. A viewer/search/indexer streaming from offset 0 without a prior `GetSize` presents/indexes/copies the object as a 3 MiB file. *(When `_sizeKnown==true` the same short read self-heals via re-request — defect is specific to the unknown-size path.)*

**Deeper fix:** Establish authoritative size via HeadObject before/around the first fill; on the unknown-size path treat `total<maxBytes` as possibly-incomplete (re-issue from `_bufferStart+total`, conclude EOF only when a non-empty remaining range returns zero), or validate against `Content-Range`/`Content-Length`.

---

**19. `PutS3ObjectFromHandle` relies on S3 to reject a short body; a truncated file can be committed as a complete object on permissive endpoints**
`Plugins/FileSystemS3/FileSystemS3.S3.cpp:1043` — **corruption / PLAUSIBLE**

The small-object path sets `ContentLength` to the stat'd `sizeBytes` and streams from `HandleReadIStream`. A premature EOF (`underflow` returns eof without setting `_readErrorHr`, 565-568) leaves `GetReadError()==S_OK`; there is no streamed-byte-count check. If the file shrank between stat and read (the #17 truncated relay, or concurrent truncation), the body is short. Canonical S3 rejects with `IncompleteBody`, but permissive S3-compatible endpoints (MinIO/Ceph/gateways) may store the short object and return success — and for MOVE the source is then deleted. *(Multipart path is guarded by `ReadExactBytesFromFile`.)*

**Deeper fix:** Count bytes streamed (in `underflow`) and fail unless `bytesStreamed == sizeBytes`; treat premature EOF before `sizeBytes` as `ERROR_HANDLE_EOF`/`ERROR_PARTIAL_COPY` rather than trusting the server.

---

**20. S3 server-side `CopyObject` for MOVE is not verified (size/ETag) before the source delete**
`Plugins/FileSystemS3/FileSystemS3.S3.cpp:868` (multipart 924; delete `Directory.cpp:1683`) — **data-loss / PLAUSIBLE**

Both branches of `CopyS3ObjectServerSide` return `S_OK` purely on `outcome.IsSuccess()` / `CompleteMultipartUpload` success, with **no** HeadObject comparing destination `ContentLength`/ETag against `sourceSizeBytes`. S3's documented "200 OK with `<Error>` in body" `CopyObject` case, if the CRT surfaces it as success, leaves the destination absent/short while `ExecuteCopyOrMove` deletes every transferred object's source (1667-1691). *(Whether the S3Crt SDK ever reports that case as `IsSuccess()==true` is version/endpoint-dependent.)*

**Deeper fix:** After `CopyObject` (and after `CompleteMultipartUpload`) HeadObject the destination, verify `ContentLength == sourceSizeBytes` (and ETag where available) before allowing the move's source-delete pass.

---

**21. IMAP single-message delete falls back to mailbox-wide `EXPUNGE`, destroying every other `\Deleted`-flagged message**
`Plugins/FileSystemCurl/FileSystemCurl.Imap.cpp:3033` — **data-loss / CONFIRMED**

`ImapDeleteMessage` flags one UID `\Deleted` (3021), tries `UID EXPUNGE {uid}` (3027, requires UIDPLUS RFC 4315), and on failure falls back to bare `EXPUNGE` (3033) — which per RFC 3501 permanently purges **every** `\Deleted`-flagged message in the mailbox. No `UIDPLUS`/`CAPABILITY` detection exists; the fallback is unconditional. Many clients defer expunge to mailbox close, so other `\Deleted` messages commonly pre-exist.

**Failure scenario:** A non-UIDPLUS server (older Dovecot/Courier/some Exchange-IMAP) plus pre-existing `\Deleted` messages; deleting one message purges all of them. Multi-select compounds it (every call repeats the full-mailbox EXPUNGE).

**Deeper fix:** Detect UIDPLUS up front; use only `UID EXPUNGE`. If absent, never issue unscoped `EXPUNGE` — surface a "server does not support per-message expunge" error (leaving only the `\Deleted` flag) or implement a safe re-flag-around-expunge equivalent.

---

**22. IMAP fetch/delete operate purely by UID with no UIDVALIDITY check — wrong-message delete/fetch after a UIDVALIDITY change**
`Plugins/FileSystemCurl/FileSystemCurl.Imap.cpp:2990` — **data-loss / PLAUSIBLE**

The leaf name encodes only the UID; `ImapFetchMessageToFile`/`ImapDeleteMessage` act via `UID FETCH`/`UID STORE`/`UID EXPUNGE` without validating the mailbox's current UIDVALIDITY against the listing's epoch. UIDs are stable only within a UIDVALIDITY epoch (RFC 3501 §2.3.1.1). After a mailbox delete+recreate (or a server that bumps UIDVALIDITY) between a cached listing and the action, UID 42 refers to a different message — irrecoverable wrong-message expunge or wrong-file download under the original name. UIDVALIDITY is parsed but used only as a STATUS display field.

**Deeper fix:** Capture UIDVALIDITY at listing time; before any UID FETCH/STORE/EXPUNGE re-read STATUS/SELECT UIDVALIDITY and abort with "mailbox changed, please refresh" if it differs.

---

**23. `ImapDeleteMessage` leaves the message permanently `\Deleted` when EXPUNGE fails — later surprise loss**
`Plugins/FileSystemCurl/FileSystemCurl.Imap.cpp:3021` — **data-loss / CONFIRMED**

STORE `\Deleted` and EXPUNGE are two independent curl performs. If STORE succeeds but every EXPUNGE attempt fails (network drop, ACL denial, server error), the function returns failure but the `\Deleted` flag is committed server-side and never rolled back (no `-FLAGS` compensation). The user is told the delete failed and assumes the message is intact; a later auto-expunge on mailbox close — or any other client's `EXPUNGE` — silently purges it.

**Deeper fix:** On any EXPUNGE failure, compensate with `UID STORE {uid} -FLAGS.SILENT (\Deleted)` before returning the error — treat STORE+EXPUNGE as a transaction.

---

**24. Microsoft Drive non-recursive folder delete uses incomplete-enumeration emptiness check, then issues a RECURSIVE id-delete that destroys hidden children**
`Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:5352` — **data-loss / CONFIRMED**

The non-recursive branch lists the folder via `ListDirectory(... entries)` **without** the `incompleteDueToInvalidChildNameOut` out-param, and refuses only when `entries` is non-empty (5357). `ParseChildren` silently skips non-object array elements and children with empty/missing `name`, signaling only through that discarded flag — so a folder containing a single malformed/empty-name child enumerates as empty, the `ERROR_DIR_NOT_EMPTY` guard is skipped, and `DeleteItemById` issues a Graph DELETE that recursively removes the folder and all descendants. The merge-move code (4058-4083) explicitly captures `reListIncomplete` and refuses to delete a folder it cannot prove empty — `DeleteItem` violates that established contract.

**Failure scenario:** Plain delete of a folder the user believes empty that contains one empty-name/malformed child (abnormal Graph response or eventual-consistency artifact). The recursive id-delete destroys the hidden child.

**Deeper fix:** Pass `incompleteDueToInvalidChildNameOut` into the emptiness list and treat incomplete enumeration (or any non-empty result) as "cannot prove empty" → `ERROR_DIR_NOT_EMPTY`. Mirror the merge-move guard; ideally use an etag/if-empty-conditioned delete so the server enforces emptiness atomically.

---

**25. Snapshot index save is a non-atomic in-place truncate-then-write; a crash or disk-full mid-write destroys the prior valid snapshot**
`Common/LocalSearchIndexCore.cpp:988` — **corruption / CONFIRMED**

`SnapshotVolumeStore::Save` opens the live snapshot with `CREATE_ALWAYS` (truncates the prior valid file to zero), then writes header + entries incrementally, returning on the first failed/partial write. No temp-file + atomic rename, no `FlushFileBuffers` (grep confirms no `MoveFileEx|ReplaceFile|FlushFileBuffers|.tmp` other than line 988). An interruption (crash, power loss, `ERROR_DISK_FULL`) between truncation and final write destroys the old snapshot and leaves a partial. The next load detects corruption (`ERROR_INVALID_DATA`) and forces an expensive whole-volume re-enumeration. *(The index is a derived cache, so this is recoverable-but-degraded rather than user-data loss — kept here because it is a destructive in-place write, the same anti-pattern as #2.)*

**Deeper fix:** Write to a sibling temp, `FlushFileBuffers`, then atomically `MoveFileExW(MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)` (or `ReplaceFileW`); delete temp only on failure.

---

**26. `EnsureSchemaVersion1` runs `BEGIN IMMEDIATE..COMMIT` as one `exec` with no `ROLLBACK` on partial failure**
`Common/SqliteIndexStore.cpp:1121` — **corruption-risk / PLAUSIBLE (current consequence: clean failed-bootstrap)**

The whole v1 schema script is one multi-statement `sqlite3_exec`; a mid-script failure (SQLITE_FULL during CREATE INDEX, SQLITE_BUSY) aborts before `COMMIT` leaving the write transaction open, with no explicit `ROLLBACK` (unlike `EnsureSchemaVersion2`, which uses a `scope_exit` ROLLBACK). *(Current callers use a local connection and return on failure, so `sqlite3_close` auto-rolls-back — no on-disk corruption today. The hazard is a future caller reusing the connection after a soft retry.)*

**Deeper fix:** Wrap v1 like v2 — separate `BEGIN IMMEDIATE`, a `scope_exit` ROLLBACK gated on a commit flag, individual statements, then `COMMIT`. Don't rely on connection teardown to roll back open DDL.

---

**27. `FindVolumeId`/`ReadVolumeRow` match `root_path` with case-sensitive BINARY collation while the in-memory layer keys on a case-folded key — duplicate volume rows and shadowed index data**
`Common/SqliteIndexStore.cpp:1583` (1081, 753; folded layer `LocalSearchIndexCore.cpp:4111`) — **corruption / CONFIRMED**

`volumes.root_path` is `TEXT UNIQUE` (BINARY), and every store lookup binds the raw, case-preserving path. But the in-memory dedup and `Load` use `FoldPathKey(...)`. The layers disagree: the persisted casing is the first writer's; a later session/process/service with differently-cased components (`C:\Users\Bob\Docs` vs `...\docs`, which `GetFullPathNameW` preserves) misses `FindVolumeId`, `InsertVolume`s a second row, and `ApplyJournalDelta` reseeds an empty volume — orphaning the original index so `EnumerateVolume` on the original casing returns stale data and the new casing a half-built index. New file changes silently never appear in search.

**Deeper fix:** Store and look up a canonical case-folded `root_path_folded` (UNIQUE) bound from `FoldText(rootPath)` in `FindVolumeId`/`ReadVolumeRow`/`InsertVolume`, guaranteeing one row per physical root regardless of caller casing.

---

### SEVERITY: CRASH / HANG

---

**28. Directory traversal/live-scan follows reparse points with no reparse guard or cycle detection — stack-overflow crash or whole-drive scan**
`Common/LocalSearchIndexCore.cpp:2438` (live-scan 3867; handle open 1542-1557) — **crash / CONFIRMED**

`HydrateDirectorySubtree` recurses into every directory child (`IsDirectoryAttributes` checks only the directory bit, no `FILE_ATTRIBUTE_REPARSE_POINT` exclusion, no visited-id/depth guard), and `GetPathNodeId` opens children **without** `FILE_FLAG_OPEN_REPARSE_POINT`, resolving through the target. `EnumerateLiveFileSystem` pushes directory children with no guard either. `SeedTraversalIndex` is the production fallback for any **non-elevated** user (`SeedNtfsIndex` falls back on `ERROR_ACCESS_DENIED`/`ERROR_PRIVILEGE_NOT_HELD`, 2529-2531). A self/ancestor-referential junction (`C:\Root\link -> C:\Root`) recurses unboundedly → stack overflow / process crash; a `link -> C:\` makes a small-folder search scan the entire drive. The `visited` set in `RebuildDerivedState` runs post-build and cannot bound the recursion.

**Deeper fix:** Skip `(fileAttributes & FILE_ATTRIBUTE_REPARSE_POINT)` children before recursing/pushing; if link-following is intended, add a visited-NodeId set + max-depth bound and open children with `FILE_FLAG_OPEN_REPARSE_POINT`. Apply to both `HydrateDirectorySubtree` and `EnumerateLiveFileSystem`.

---

**29. Reentrant nested jobs on the single bounded scheduler pool deadlock cross-FS directory copy/move (operation hangs forever)**
`RedSalamander/FolderWindow.FileOperations.State.cpp:8146` (outer dispatch 8866/8881; nested 7931) — **crash (hang) / PLAUSIBLE**

There is one process-wide `PerItemTaskScheduler` (function-local static). The outer per-item dispatch blocks a worker in `WaitJob`; for a cross-FS **directory** item the worker calls `CopyDirectoryParallel`, which submits a **nested** job on the **same** scheduler and blocks that worker in a second `WaitJob` (8146). `WaitJob` has no work-stealing. *(The starvation guard `effectiveMaxConcurrencyLocked` reserves `workerCount-1` for multi-item jobs, refuting the multi-core deadlock; the genuine residual hang is a **single-worker pool** — `hardware_concurrency==1` — with a plugin advertising `copyMoveMax>1`, where the reserve degenerates to 0 and the lone worker blocks on a child with no worker to drain it.)* For a MOVE this leaves the source intact but the operation wedged; user must kill the app.

**Deeper fix:** Reentrant work submitted from inside a scheduler worker must never block on the same bounded pool: give the bridge its own disjoint scheduler, or make `WaitJob` work-steal, or run the bridge directory consumers inline on the producer thread when the call originates from a worker (detect via a thread-local "inside worker" flag).

---

**30. Scheduler synchronous fallback runs a blocking consumer before its producer, deadlocking under thread-creation failure**
`RedSalamander/FolderWindow.FileOperations.State.cpp:3429` (empty-workers 3629-3643; producer 7973-8144) — **crash (hang) / CONFIRMED**

`StartJob` has a synchronous fallback: if `_workers.empty()` it runs `processIndex` inline on the calling thread and returns only after completion. `_workers` is permanently empty when every `std::jthread` creation throws `std::system_error` (handle/thread exhaustion) — the catch breaks, `_initialized=true`. `CopyDirectoryParallel` relies on `StartJob` returning **before** the producer loop runs; in the fallback it instead invokes `workerProc(0)` inline, which parks in `workCv.wait_for` waiting for items while `producerDone` is false — and `StartJob` never returns, so the producer never runs. `CopyDirectoryParallel` only diverts to sequential when `withinFolderBudget <= 1`; for budget>1 it wedges with no error (MOVE leaves source intact).

**Deeper fix:** Don't use the synchronous fallback for producer/consumer jobs. Have `StartJob` return a failed/null job when no workers exist so `CopyDirectoryParallel` routes to `CopyDirectorySequential`; or enqueue all work and set `producerDone` **before** `StartJob` so an inline consumer terminates.

---

**31. Arena path builder copies `FileName` for `FileNameSize` bytes without bounding `FileName` against the record extent — OOB read / wrong-target**
`Common/PlugInterfaces/FileSystem.h:1039` — **crash / CONFIRMED**

`BuildFileSystemPathListArenaFromFilesInformation` validates `NextEntryOffset` against the buffer but never validates that `FileName`'s payload (`FileNameSize` bytes at `offsetof(FileInfo,FileName)`) fits inside the entry's record. For the last entry (`NextEntryOffset==0`) there is **no** record upper bound at all; the sizing pass checks only `FileNameSize % sizeof(wchar_t)==0`. A hostile/buggy `IFilesInformation` plugin returning a final entry with `FileNameSize = 0x40000000` and a few committed bytes makes `CopyMemory` read ~1 GiB past the buffer → access violation, or copies adjacent heap into a user-visible "source path" handed to Delete/Move (wrong target).

**Deeper fix:** Validate `FileNameSize` against per-entry available payload in **both** passes (`NextEntryOffset - offsetof(FileInfo,FileName)` for non-last; `bufferSize - offset - offsetof` for the tail), reject with `ERROR_BAD_LENGTH` on overflow, and document in `FileInfo` that `FileName` MUST be contained within `NextEntryOffset`/remaining buffer.

---

### SEVERITY: INCOMPLETE-RESULTS

---

**32. Parallel content-scan worker `return`s on `MaterializeEntryPaths` failure, silently dropping all remaining matches in the chunk and reporting `S_OK`**
`Plugins/FileSystem/FileSystem.Search.cpp:2155` (and 2198) — **incomplete-results / CONFIRMED**

The two `return;` statements exit the **entire** 128-entry chunk loop (not just the bad entry) with `result.status` defaulting to `S_OK`, so `ParallelEvaluateEntries` reports overall success. `MaterializeEntryPaths` returns false for a zero-length-name entry (corrupt FS / network/virtual FS metadata). The sequential path handles the identical failure per-entry (`return S_OK;` skips one entry) — the parallel path is strictly worse. Up to 127 later files in a chunk (≥500-entry directories) are silently abandoned, including files containing the search term.

**Deeper fix:** Replace both `return;` with `continue;` (skip only the malformed entry) and set a warning flag in `result.warningFlags` so incompleteness surfaces, mirroring `EvaluateEntry`.

---

**33. LITERAL name-search SQL prefilter forces exact equality, silently dropping all substring matches**
`Common/SqliteIndexStore.cpp:671` (SQL 903; matcher `LocalSearchIndexCore.cpp:2236`) — **incomplete-results / CONFIRMED**

`TryBuildNamePrefilter` maps `FILESYSTEM_SEARCH_NAME_LITERAL` to `ExactFolded`, emitting `AND name_folded = ?` — but the authoritative `MatchLiteral` is a **substring** match (`find(...) != npos`). Substring-only rows are filtered out in SQL and never reach the post-filter. An indexed literal search for `report` returns zero rows against `Report_2024.txt`/`annual-report.pdf`/`reporting.xlsx` and reports `S_OK`. Self-tests mask it by always passing the full file name (exact == substring).

**Deeper fix:** Disable the prefilter for LITERAL (fall to `NamePrefilterKind::None`, let `MatchName` do the work) — a leading-`%` LIKE defeats the index anyway, so honest correctness lives in the post-filter.

---

**34. Wildcard `*.<multi-part-extension>` prefilter compares against the last-segment-only extension column, dropping all matches**
`Common/SqliteIndexStore.cpp:685` (SQL 913; column `GetExtensionFolded` 161-165) — **incomplete-results / CONFIRMED**

`*.tar.gz` sets `ExtensionFolded = FoldText(".tar.gz")` and emits `AND extension_folded = ?`, but the indexed `extension_folded` is `path.extension()` = `.gz` only. The equality can never match; `*.tar.gz`/`*.d.ts`/`*.min.js`/`*.blade.php` all return empty `S_OK` while the authoritative `MatchWildcard` returns true.

**Deeper fix:** Only use `ExtensionFolded` when the portion after `*.` has no further `.` (single segment matching `std::filesystem` extension semantics); otherwise fall through to `None` and rely on the post-filter.

---

**35. Hardlinked files collapse to one index entry keyed by file id, dropping all but one path**
`Common/LocalSearchIndexCore.cpp:2764` (seed 2599) — **incomplete-results / CONFIRMED**

Every entry is keyed solely by `NodeId` (the NTFS FileReferenceNumber); hardlinks share one reference but have distinct `(parentId, name)`. The map holds one entry per id, so the last-seen link overwrites the others, and `USN_REASON_HARD_LINK_CHANGE` flows through the same single-entry upsert (creating a hardlink replaces an existing one's path). A name/wildcard search matching both `C:\Root\report.docx` and a hardlink elsewhere returns only one — order-dependent.

**Deeper fix:** Key the index (and store) by `(fileId, parentId, name)`, or maintain a per-id list of `(parentId, name)` links so each hardlink yields its own queryable path.

---

**36. `GetDirectorySize` uses non-extended (MAX_PATH-limited) paths, silently under-counting deep trees**
`Plugins/FileSystem/FileSystem.DirectoryOps.cpp:708` (753, 832-837) — **incomplete-results / CONFIRMED**

Unlike every sibling op, `GetDirectorySize` never calls `ToExtendedPath`; it seeds and builds child paths from the raw, non-`\\?\` string. Once cumulative recursion crosses 260 chars (`node_modules`, build trees, a tree near an already-deep root), `FindFirstFileExW` fails `ERROR_PATH_NOT_FOUND` — not in the swallow list, so `status` is set to failure, but all shallower bytes are still returned and every subtree below the boundary is skipped. "Calculate size" silently reports far less than reality (or errors on a fully `\\?\`-enumerable tree).

**Deeper fix:** `ToExtendedPath(path)` once at entry; build all search/child paths from the extended form (as `StartEnumerationWin32` does); use extended paths for the initial attribute probes.

---

**37. Access-denied / vanished subdirectories are swallowed, so `GetDirectorySize` returns `S_OK` with an under-counted total**
`Plugins/FileSystem/FileSystem.DirectoryOps.cpp:712` — **incomplete-results / CONFIRMED**

`pushDirectory` excludes `ERROR_FILE_NOT_FOUND` and `ERROR_ACCESS_DENIED` from `result->status` and returns, abandoning the subtree — but the directory was already counted (811). The result struct has no incomplete/partial flag, so an under-counted total comes back `S_OK` and the UI presents it as authoritative ("will this fit on the target?" decisions on a silently-low figure). Root-level inaccessibility is correctly surfaced; the gap is subtrees (ACL-protected, or TOCTOU-deleted between enumeration and descent).

**Deeper fix:** Track an "incomplete" flag / skipped-entry count and surface it (`S_FALSE`/dedicated HRESULT and/or progress-callback report) so a partial total is never presented as exact.

---

**38. Diagnostics/issues plumbing: short writes, dedup, and midnight-crossing mis-attribution lose or corrupt the failure record**
*(Grouped — all `unit key bridge-diagnostics`, all about the post-incident record of what failed; none affect user file data but each makes the audit trail unreliable.)*

- **`ExportTaskIssuesReport` ignores the `WriteFile` byte count → truncated report presented as success.** `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp:180` (126, 150) — **incomplete-results / CONFIRMED.** Three `WriteFile` calls test only the BOOL; on a near-full volume a short write (`written < bytesToWrite`) leaves a half-written UTF-16 (odd-byte) line, the function returns true, and `ShellExecute` opens a garbled "complete" record. `FlushDiagnostics:797` does it correctly. **Fix:** require `written == bytesToWrite`; on mismatch delete the partial and return false.
- **`FlushDiagnostics` partial write produces a corrupt half-line + a duplicated full line in the JSONL log.** Same file `:797` — **corruption / PLAUSIBLE.** The `!WriteFile || written != bytesToWrite` branch requeues the entry without distinguishing a hard failure (0 bytes committed) from a partial success (partial bytes already appended), so the next flush rewrites the line in full after a truncated fragment — breaking line-based JSONL parsing for everything after. *(Trigger needs `WriteFile`==TRUE with a short count on a synchronous disk handle — documented mainly for pipes/redirectors.)* **Fix:** build the batch into one buffer and write atomically; dedupe on requeue.
- **Issues-pane dedup key uses only `wMilliseconds`, collapsing distinct same-path/same-message failures.** `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp:423` — **incomplete-results / PLAUSIBLE.** The key truncates the timestamp to the sub-second field; two repeat records for one path/error sharing `wMilliseconds` (1-in-1000) drop the second. *(Path and message are in the key, so genuinely different files are not collapsed — narrower than the original claim.)* **Fix:** include the full timestamp or a per-entry monotonic sequence id.
- **Per-task diagnostics log path stamped at completion mis-attributes tasks crossing midnight.** `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.Part.cpp:956` — **incomplete-results / CONFIRMED.** Diagnostics flush continuously into per-day files (~5s interval); the summary stores only the completion-day path, so an overnight task's pre-midnight warnings (the bulk) sit in yesterday's file, unreachable from the task UI. **Fix:** record the set of day-files the task wrote to (or aggregate by taskId across day files when opening).

---

**39. Shared-drive listing proceeds with an empty `driveId` when UTF-8 conversion yields empty, querying the wrong corpus**
`Plugins/FileSystemGoogleDrive/FileSystemGoogleDrive.cpp:1792` — **robustness (downgraded from incomplete-results) / PLAUSIBLE**

The `sharedDrive` branch emits `driveId=Utf8FromUtf16(connection.sharedDriveId)` without the empty-check the sibling `parentId` path has (1774-1778). On conversion failure (unpaired surrogate / oversized) it emits `driveId=` with no value and `corpora=drive`, producing an undefined query. *(JSON-sourced IDs normally round-trip; concrete trigger requires corrupt/edge config.)* **Fix:** convert once up front and bail `ERROR_NO_UNICODE_TRANSLATION` when empty-while-non-empty.

---

### SEVERITY: SECURITY

---

**40. LocalSystem search service performs no client impersonation or access check — any interactive user enumerates metadata of paths they cannot read**
`Common/SearchServiceBroker.cpp:2107` (SDDL 2287; service account `Main.cpp:2592-2593`) — **security (information disclosure / privilege) / CONFIRMED**

The service runs as LocalSystem (no account arg to `CreateServiceW`); the pipe SDDL grants `GRGW` to all Interactive Users (`(A;;GRGW;;;IU)`). `HandleQueryRequest` takes the client-supplied `rootPath` verbatim and enumerates it through the repository under **SYSTEM's** token, streaming names/sizes/timestamps/paths back. Grep confirms **zero** `Impersonate`/`AccessCheck`/`CheckTokenMembership` and no per-caller path allowlist or index partitioning. A standard user querying `C:\Users\<other>` or protected system dirs receives listings they have no rights to.

**Deeper fix:** Before acting on any caller path, `ImpersonateNamedPipeClient`, perform the enumeration (or at minimum an `AccessCheck`/`CreateFile` probe of the root) under the caller's token, `RevertToSelf` via `scope_exit`, and reject on impersonation failure or denied access. Partition/access-filter the cache per caller so the index cannot leak metadata across security boundaries.

---

**41. Server frame I/O uses synchronous `ReadFile`/`WriteFile` (NULL OVERLAPPED) on a `FILE_FLAG_OVERLAPPED` handle — UB, possible torn frames**
`Common/SearchServiceBroker.cpp:796` (819; pipe 2305) — **corruption / CONFIRMED**

`CreateServerPipe` opens with `FILE_FLAG_OVERLAPPED`, but `ReadExact`/`WriteExact` call `ReadFile`/`WriteFile` with `lpOverlapped=nullptr`. Per the Win32 contract this is erroneous for an overlapped handle: the byte-count out-param may be unreliable and the call can return before completion. The framing loops trust `chunkRead`/`chunkWritten`, so a wrong count advances past/short of frame boundaries → torn frames, mis-parsed `payloadBytes`, or a header mixing two frames. The client path uses overlapped correctly (`ClientIoExact`), confirming the server-side mix is the defect. Triggered by any frame larger than one buffer fill.

**Deeper fix:** Make server I/O overlapped consistently (route through a `ClientIoExact`-style helper with its own `OVERLAPPED`+event), or create the server pipe **without** `FILE_FLAG_OVERLAPPED` and use a separate overlapped handle only for the `ConnectNamedPipe` wait.

---

### SEVERITY: ROBUSTNESS / CONTRACT / EFFICIENCY (terse)

- **Crash/power-loss orphans `.rs_tmp_*` staging files with no recovery sweep.** `FolderWindow.FileOperations.State.cpp:7114` — robustness/CONFIRMED. Add a startup/destination sweep keyed on owning PID/start-tick embedded in the temp name.
- **Overall progress clamped to a stale pre-calc total, freezing below 100%.** `...State.cpp:6641` — robustness/CONFIRMED. Raise the working total to `max(total, callCompleted)` instead of clamping down; reserve clamping for authoritative per-file `GetSize`.
- **Apply-to-all has no end-of-task audit and one `Exists` bucket spans file/dir/reparse.** `...State.cpp:8716` — robustness/PLAUSIBLE (designed behavior; the real gaps are the missing overwrite summary and cross-kind re-confirmation — see #11 for the data-loss edge).
- **Recursive sequential delete returns `DIR_NOT_EMPTY` instead of `PARTIAL_COPY` under continue-on-error.** `FileSystem.FileOps.cpp:6930` — contract/CONFIRMED. Gate the parent disposition on `hadFailure` and return `ERROR_PARTIAL_COPY`, mirroring the parallel driver.
- **Resume-skip identical-file compare is a TOCTOU snapshot (shared-write handle).** `FileSystem.FileOps.cpp:3523` — robustness/CONFIRMED. Open the equality-check handle `FILE_SHARE_READ` only, or re-verify under an exclusive handle before declaring complete.
- **Reparse file-symlink move strips source READONLY but never restores it when the source delete fails.** `FileSystem.FileOps.cpp:5229` — robustness/CONFIRMED. Wrap in a `scope_exit` re-applying READONLY, disarm only on delete success (mirror `DeleteOpenedPathNoFollow`).
- **`OpenPathForDeleteNoFollow` unconditionally requires `FILE_WRITE_ATTRIBUTES`, failing deletes for files the user can delete but not modify.** `FileSystem.FileOps.cpp:2785` — robustness/CONFIRMED. Open `DELETE | FILE_READ_ATTRIBUTES`; request `FILE_WRITE_ATTRIBUTES` lazily only in the readonly-clear branch.
- **`_directorySizeDelayMs` read unsynchronized while written under mutex (debug pacing hook).** `FileSystem.DirectoryOps.cpp:796` — robustness/CONFIRMED (data race, `_DEBUG` only). Make it `std::atomic` or snapshot under the mutex.
- **Writer open lacks `FILE_FLAG_OPEN_REPARSE_POINT`; `CREATE_ALWAYS` follows symlinks and the abort `DeleteFileW` removes the link.** `FileSystem.cpp:1958` — robustness/PLAUSIBLE (no in-tree caller passes a symlink dest to ALLOW_OVERWRITE; latent in the public primitive). Decide reparse policy explicitly; only ever delete the writer's own temp on abort.
- **`MakeAbsolutePath` mishandles `GetFullPathNameW`'s buffer-too-small return on the second call (CWD race) → NUL-filled path.** `FileSystem.Path.cpp:45` — robustness/PLAUSIBLE. Loop while `written >= bufferLen`; accept only `0 < written < bufferLen`.
- **Directory timestamp/attribute restore fails: `FILE_ATTRIBUTE_DIRECTORY` passed to `SetFileInformationByHandle` aborts before `SetFileTime`.** `FileSystem.FileOps.cpp:2728` — contract/CONFIRMED. Mask non-settable bits before the call; attempt timestamps independently of the attribute write.
- **Redundant `SetFileTime` can clobber timestamps to 1601 when a source field is 0.** `FileSystem.FileOps.cpp:2736` — corruption/PLAUSIBLE (needs a backend reporting a 0 time field). Drop the redundant `SetFileTime`, or pass `nullptr` for any field whose source QuadPart is 0.
- **7z `NormalizeArchiveEntryKey` does not collapse `..` segments (latent zip-slip).** `FileSystem7z.cpp:1794` — robustness/PLAUSIBLE (read-only keyed map today; no consumer joins keys to a real path). Reject/canonicalize `..` in key derivation.
- **7z entry whose name collides with a directory key makes nested members unreachable.** `FileSystem7z.cpp:4416` — incomplete-results/CONFIRMED. On collision, never demote a directory-with-children to a file; surface `ERROR_INVALID_DATA`.
- **`MultipartS3FileWriter::Commit` can leak an orphaned multipart upload when the single abort attempt also fails.** `FileSystemS3.IO.cpp:818` — robustness/PLAUSIBLE (uploadId already logged at Error; SDK retries internally). Clear session identity only on abort success; persist uploadId for a later sweep.
- **S3 recursive prefix delete lists-then-deletes with no completeness re-verify (eventual-consistency TOCTOU).** `FileSystemS3.Directory.cpp:1252` — incomplete-results/PLAUSIBLE. Re-list (`MaxKeys=1`) after the batch; loop bounded passes; only return `S_OK` when the prefix confirms empty.
- **S3 list pagination treats an empty `NextContinuationToken` on a truncated page as a restart (infinite loop / dup enumeration).** `FileSystemS3.Directory.cpp:756` (`DirectoryOps.cpp:337`, `S3.cpp:419`) — robustness/PLAUSIBLE (non-conformant endpoints). Abort `ERROR_INVALID_DATA` when truncated with empty token (the `S3Table.cpp:152` guard already does this).
- **S3 overwrite-with-backup MOVE reports `ERROR_PARTIAL_COPY` when only post-success backup cleanup fails, orphaning `.rs-bak-*`.** `FileSystemS3.Directory.cpp:1694` — robustness/CONFIRMED. Return `S_OK` (or a warning) once user-visible objects are correct; retry/schedule backup deletion separately; filter `.rs-` from listings.
- **MS-Drive chunked upload ignores Graph `nextExpectedRanges`; offset advances blindly + no completion verification + no resume.** `FileSystemMicrosoftDrive.cpp:6290` — robustness/PLAUSIBLE (silent desync needs a partial-acceptance 202). Parse `nextExpectedRanges`, re-seek from the server's offset, verify committed size/hash, GET the session URL to resume on transient failure.
- **MS-Drive ranged download accepts a 206 without validating `Content-Range` matches the requested offset.** `FileSystemMicrosoftDrive.cpp:4561` — corruption/PLAUSIBLE (needs a non-conformant CDN 206). Parse/verify `Content-Range` start==offset and total==`_sizeBytes`; refresh URL + retry on mismatch (the `HttpResponse` struct lacks a `Content-Range` field — add one).
- **MS-Drive overwrite-with-backup MOVE reports failure when only the post-success backup delete fails; orphans `.redsalamander-rollback-*` and a retry then fails not-found.** `FileSystemMicrosoftDrive.cpp:4246` — robustness/CONFIRMED. Treat backup cleanup as best-effort: return `S_OK` after the move succeeds, retry orphan cleanup separately.
- **Curl `DeleteItems` folds a global concurrency from heterogeneous connections inside the loop; one serial host collapses the whole batch.** `FileSystemCurl.CopyMove.cpp:3368` — efficiency/CONFIRMED. Compute scheduler width from settings and let the per-connection `ConnectionConcurrencyLimiter` enforce caps; remove the in-loop `min`.
- **Curl IMAP error mapping classifies auth/permission text only for `CURLE_QUOTE_ERROR`.** `FileSystemCurl.Imap.cpp:178` — robustness/PLAUSIBLE (the dominant `CURLE_LOGIN_DENIED` path is already mapped; narrow residual gap for `WEIRD_SERVER_REPLY`/`RECV_ERROR`/`TIMEDOUT`). Apply the server-line text classification regardless of CURLcode when a tagged NO/BAD was captured.
- **Google Drive duplicate-name navigation matches file IDs case-insensitively (`EqualsNoCase`).** `FileSystemGoogleDrive.cpp:1943` — robustness/PLAUSIBLE (Drive IDs are case-sensitive but never minted as case-variants, so wrong-file is practically unreachable). Use ordinal case-sensitive `Equals`; match purely on the unique id; don't `break` on a fuzzy hit.
- **Google Drive files literally named `… [id:…]` or unnamed items are listed but unreachable (round-trip broken).** `FileSystemGoogleDrive.cpp:1935` — robustness/CONFIRMED. Only attempt synthetic-id parsing for segments the plugin emitted; try exact plain-name match first; use a non-collidable marker and an id-only synthetic name for unnamed items.
- **WAL checkpoint `TRUNCATE` returning `SQLITE_BUSY` aborts maintenance whenever a reader is attached.** `SqliteIndexStore.cpp:586` — robustness/CONFIRMED. Treat `SQLITE_BUSY` on TRUNCATE/RESTART as non-fatal (retry with backoff or fall back to `PASSIVE`) and continue; enable extended result codes.
- **No `integrity_check`/`quick_check` (or extended codes) before VACUUM rewrites the whole DB; `SQLITE_CORRUPT` collapses to generic `E_FAIL` with no rebuild.** `SqliteIndexStore.cpp:2628` (auto 2772) — robustness/CONFIRMED. Run `quick_check` first; enable extended codes; on `SQLITE_CORRUPT`/`NOTADB` trigger a controlled reseed (it's a derived cache) and surface a distinct status.
- **`ApplyJournalDelta` reports `deleted/upsertedEntryCount` for delta ops it skipped on the seed path.** `SqliteIndexStore.cpp:2477` — contract/CONFIRMED (latent; no current caller trusts the counts). Set both counts to 0 on the seed path.
- **`GetEnvironmentValue` lacks the `written >= required` truncation guard its sibling has — racing env-var change yields a truncated pipe name (no default fallback).** `SearchServiceBroker.cpp:1320` — robustness/CONFIRMED. Mirror `ReadEnvironmentVariableTrimmed`; consolidate the two helpers.
- **No-batch-callback `QueryComplete` failure leaves partial candidates in `outCandidates` (inconsistent with overflow/catch paths that clear).** `SearchServiceBroker.cpp:2952` — contract/CONFIRMED (test/accumulation path; production passes a callback). Clear on any FAILED return, or set a partial-results flag.
- **No-batch-callback overflow discards the entire result set instead of returning the capped prefix.** `SearchServiceBroker.cpp:2885` — incomplete-results/CONFIRMED (buffered API path; production avoids via callback). Return the truncated prefix + a truncation flag rather than clearing and failing.
- **`WaitForClientIo` recomputes the per-frame deadline every wakeup → a trickling peer is bounded only by the 10-min operation deadline (slow-loris).** `SearchServiceBroker.cpp:936` — robustness/CONFIRMED. Track the frame deadline once at frame start; enforce a real per-frame ceiling.
- **Single-threaded server loop with an untimed blocking `ReadExact` lets one local client hang the whole service.** `SearchServiceBroker.cpp:793` — robustness/CONFIRMED (local authenticated users only). Give server reads the overlapped/deadline/stop-event path the client has; dispatch clients to a bounded pool; cancel via `CancelIoEx`.
- **Connected client handle never `DisconnectNamedPipe`d on a hard `HandleClient` failure; the whole server loop exits on one client-induced error.** `SearchServiceBroker.cpp:3347` — robustness/CONFIRMED. `Disconnect`/`Flush` in a `scope_exit`, log, and `continue`; reserve termination for genuinely fatal errors.
- **Service shutdown cannot interrupt an in-flight client request → SCM-forced kill mid-operation (can interrupt WAL checkpoint).** `RedSalamanderSearchService/Main.cpp:3210` — robustness/CONFIRMED. Thread the stop event + cancel callback into `HandleClient`/read/query; `CancelIoEx` on stop; make maintenance checkpoint-safe and refresh the wait hint while it runs.
- **Plugin Refresh drops mount/connection context: pane still shows archive/remote path but the new `IFileSystem` is never re-`Initialize`d.** `FolderWindow.FileSystem.cpp:4288` — corruption (downgraded from data-loss)/CONFIRMED. `ReleaseFileSystemPluginsForRefresh` leaves `currentPath`/`instanceContext`, `EnsurePaneFileSystem` clears `instanceContext` and never `Initialize`s, no `SetFolderPath` re-issued → unbound backend admits operations against a default/empty root. Re-issue `SetFolderPath(state.currentPath)` after reload, or preserve `instanceContext` and re-`Initialize`; until rebound, neutralize `currentPath`. *(7z errors out on empty `_archivePath`; the silent wrong-target delete depends on a remote backend defaulting an empty context.)*
- **7z shorthand `7z:<relative>` resolves the archive path only when the current pane is local.** `FolderWindow.FileSystem.cpp:4921` — robustness/CONFIRMED. From a non-file VFS, the bare relative name reaches `FileSystem7z::Initialize` and is anchored to the process CWD (`GetFullPathNameW`) → wrong/empty mount. Resolve against a well-defined local base independent of the active plugin; reject if none.
- **`ReloadFileSystemPlugins` returns `S_OK` even when both panes fail and no fallback exists.** `FolderWindow.FileSystem.cpp:4299` — contract/CONFIRMED. Panes left with null `fileSystem` but stale `currentPath`, result discarded by the caller; no UI signal of the dead pane. Clear/neutralize `currentPath`, surface an error, return a failing HRESULT.
- **`SetConnectionSecret` mutates the in-memory session cache before persisting; a persist failure returns the error but the session silently holds the new/cleared secret.** `RedSalamander/HostServices.cpp:2262` — contract/CONFIRMED. Persist first; mutate the (highest-priority) session cache only on success; restore the prior entry via `scope_exit` on failure.
- **Connection-secret host callbacks run their UI-thread bodies on a worker thread when the host window is null — racing the session-secret maps and `g_settings`.** `HostServices.cpp:694` — corruption/CONFIRMED. Four entry points fall through to the `*OnUiThread` body when `GetInitializedHostWindow()` is null (startup/teardown), racing the destructor's `clearAllSessionSecrets` on unlocked `unordered_map`s (UB). Return `ERROR_INVALID_WINDOW_HANDLE` when the window is null; guard the three `_session*` maps with a dedicated mutex.
- **Cross-thread `SendMessage` failure (window torn down) misreported as "secret not found"/`E_FAIL` instead of transient.** `HostServices.cpp:666` (`GetConnectionJsonUtf8` 624) — contract/PLAUSIBLE. A failed/aborted send returns 0 (indistinguishable from real "no secret"); `GetLastError` never consulted, so a credentialed connection silently downgrades to anonymous. Check `GetLastError` after a 0 return, or have the handler always write an explicit status field.
- **`IFileReader::Read` contract omits short-read semantics — consumers that don't loop silently truncate.** `Common/PlugInterfaces/FileSystem.h:538` — contract/CONFIRMED (header documentation gap; the root licensor of #1/#15/#16/#17/#18; all live consumers currently loop correctly after the ViewerText hex fix). Document normatively: `S_OK` may return `*bytesRead` in `[0, bytesToRead]`; only `*bytesRead==0` is EOF; callers MUST loop. Ship a shared `ReadFully()` helper so consumers don't re-derive the loop.
- **`FileSystemDirectoryChange`/`FileSystemDirectoryChangeNotification` string ownership/lifetime under-specified for size-prefixed non-NUL buffers.** `Common/PlugInterfaces/FileSystem.h:668` — contract/PLAUSIBLE (the search-struct half is refuted — line 159 already specifies it). Add an explicit "plugin-owned, valid only until `FileSystemDirectoryChanged` returns; host MUST copy honoring `*Size`" clause, since the notes direct cross-thread `PostMessage` delivery.

---

## 3. Cross-cutting themes

1. **Unknown-size / short-read truncation treated as EOF (the dominant theme).** The same defect recurs across the cross-FS bridge COPY (#1), 7z spool (#15), Curl download (#16), S3 download + ranged reader + PutObject (#17/#18/#19), all flowing from the **under-specified `IFileReader::Read` contract** (#20). Every instance turns a transient `S_OK`/short-or-0-byte read into a silently truncated file reported as success. The author already recognized the hazard (MOVE bails on unknown size at 7094; multipart S3 and the streaming-pipe 7z path guard correctly) — COPY and the spool path were simply left unprotected. **Systemic fix:** tighten the Read contract, ship `ReadFully`, and make every transfer require a proven byte-count match before treating a 0-byte read as completion.

2. **Destructive-before-committed in-place overwrite.** `CopyFileExW`/`CREATE_ALWAYS` on the live destination (#2, #3, #7), dummy/MS-Drive destroy-before-replace (#14, and the rollback-cleanup conflations), and the non-atomic snapshot save (#25) all truncate or delete the original before the replacement is durable. The cross-FS bridge's temp-stage + atomic-promote pattern and same-volume MOVE's `MOVEFILE_REPLACE_EXISTING` are the correct templates that the same-FS COPY and the writer never adopted.

3. **Delete from a stale enumeration snapshot / eventual-consistency TOCTOU.** Cross-FS directory MOVE blind-deletes the source root (#4), move-source delete is identity-bound not image-bound (#5), MS-Drive non-recursive delete trusts an incomplete enumeration (#24), and S3 recursive prefix delete never re-verifies (S3 bullet). The same-FS MOVE path's re-confirm-then-delete-bottom-up is the contract the others should inherit.

4. **Reparse-point following with no guard.** Metadata copy follows links to the target (#9), traversal/live-scan recurses through junctions with no cycle/depth bound (#28, a crash), the writer's `CREATE_ALWAYS` follows symlinks, and the directory watch (`FileSystem.Watch.cpp:127`) watches the target not the link. The plugin uses `FILE_FLAG_OPEN_REPARSE_POINT` correctly in five other places — the omissions are inconsistencies, not policy.

5. **Swallowed / mis-mapped errors hiding incompleteness.** `GetDirectorySize` swallows access-denied subtrees and MAX_PATH failures (#36/#37), the parallel scanner drops a whole chunk on one bad entry (#32), `ReadDirectoryChangesW` zero-byte-overflow and re-arm failure are silently dropped (`FileSystem.Watch.cpp:388`/`414`), the SQL prefilters silently narrow LITERAL/multi-dot-wildcard semantics (#33/#34), and generic `E_FAIL` masks `SQLITE_CORRUPT`/auth failures. A recurring sub-pattern: **the result struct/HRESULT has no "partial/incomplete" channel**, so under-counts and dropped results are presented as authoritative.

6. **Apply-to-all / cached-grant consent leaking across semantic categories** (#11) and **completion-reported-as-failure on post-success cleanup** (S3 `.rs-bak-*`, MS-Drive `.redsalamander-rollback-*`) — both conflate distinct states behind one signal, producing either silent over-destruction or misleading failures that drive retries into not-found errors.

---

## 4. Coverage & residual risk

**Audited (code-level, with line evidence):** the file-operations bridge (cross-FS copy/move/delete, conflict resolution, staging/promote, progress, diagnostics, scheduler/queue, cancel/pause); the local `FileSystem` plugin (CopyFile engine, move/rename, delete, DirectoryOps, reader/writer, path normalization, attributes/timestamps, directory watch); all six virtual FS plugins (7z, Curl/FTP/IMAP, S3, GoogleDrive, MicrosoftDrive, Dummy reference); the SQLite index store (schema/migration, volume rows, journal delta, maintenance/VACUUM/checkpoint, query building); `LocalSearchIndexCore` (seed/traversal/USN, snapshot persistence, derived state, live-scan); the search service (broker IPC, host process, impersonation, pipe security); and the `PlugInterfaces` contracts (`IFileReader`, `FileInfo` arena, change/match string ownership). Each kept finding cites concrete `file:line` evidence; REFUTED claims and over-broad sub-claims were trimmed during verification (noted inline).

**What only an at-rest / integration / fuzz test can confirm — a human should manually verify:**

- **Backend-dependent truncation triggers (the PLAUSIBLE data-loss core).** #16/#17/#19/#20 all hinge on whether a specific backend (SFTP/SCP/SIZE-less FTP; MinIO/Ceph/gateway S3; the AWS S3Crt "200-with-error-body" case) actually returns success on a short/interrupted body. Drive real truncation against those backends (kill the connection mid-transfer; a server that closes the data channel cleanly after a short RETR/STOR) and confirm whether the destination is short and the source is deleted. **Highest-value manual test.**
- **Cross-FS directory MOVE TOCTOU (#4) and move-source compare/delete race (#5).** Run a cross-volume directory MOVE while an external process writes new files into / appends to already-enumerated source subdirs; verify whether the new data survives. Requires a controlled concurrent writer and timing — not visible statically.
- **Conflict-cache reparse leak (#11).** End-to-end: a multi-item copy with item 1 = plain-file Overwrite + Apply-to-all, item N = a junction/symlink destination; confirm the link is destroyed with no second prompt. Needs a real FolderView interaction.
- **Reparse-cycle crash (#28).** As a **non-elevated** user, create `C:\Root\link → C:\Root` and run an indexed/live search; confirm stack-overflow vs graceful skip. The elevation path (MFT seed) is safe, so this needs the non-admin fallback specifically.
- **IMAP destructive paths (#21/#22/#23).** Against a **non-UIDPLUS** server with pre-existing `\Deleted` messages, delete one message and confirm whether others are purged; force a UIDVALIDITY change between listing and action; drop the connection after STORE but before EXPUNGE. Pure server-behavior dependence.
- **Crash/power-loss safety (#2/#3/#7/#25, and orphaned `.rs_tmp_*`).** Only a kill-the-process / pull-power harness mid-write confirms the original-file and snapshot survival claims and the absence of a recovery sweep.
- **Search-service concurrency & security (#29/#30/#40/#41, untimed read, slow-loris, shutdown-kill).** The single-worker scheduler deadlock (#29) needs a 1-logical-core host (or forced budget) with a `copyMoveMax>1` plugin; the impersonation hole (#40) needs a standard user querying another user's tree; the torn-frame UB (#41) and slow-loris need a crafted client. Run the broker under a multi-client / hostile-client / `sc stop`-mid-query harness.
- **Volume-row case collision (#27).** Index a root from the app, then from the background service (or a differently-cased path), and confirm whether new file changes appear in search — requires exercising the two FoldPathKey boundaries across processes.
- **At-rest verification absent throughout.** No copy/move path computes a content hash; "verified" means size-only. A human should decide whether content verification (or `MOVEFILE_WRITE_THROUGH` + post-read compare) is warranted for the irreversible MOVE case (#6) on unreliable backends.

Static analysis cannot establish any of the above because each depends on runtime timing, external concurrent mutation, abnormal termination, or third-party server/SDK behavior that the in-tree code does not deterministically produce.
