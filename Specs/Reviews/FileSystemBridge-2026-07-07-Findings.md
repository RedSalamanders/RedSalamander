> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# Cross-File-System Bridge — Review Findings (2026-07-07)

Method: multi-agent review (4 mappers, 6 review dimensions, adversarial verification — 61 agents). 40 findings raised, 38 confirmed, 2 refuted. Severities below are post-verification.

Scope: the host cross-filesystem copy/move engine (`RedSalamander/FolderWindow.FileOperations.State.cpp` `CrossFileSystemBridge`), the `IFileReader`/`IFileWriter` plugin contract, and all seven filesystem plugins.

## Remediation status (2026-07-13)

Operation Causeway resolved the 38 confirmed bridge/provider findings in this historical review. The implementation and lasting contracts are recorded in `Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md` and `Specs/Core/Core_FileSystemBridge.md`.

- Findings 1–2 and 15: host-side structural/destination component validation and sibling collision rejection before every path join.
- Findings 3, 19, 25, 33: bridge copy-time FNV-1a, final COPY size/hash verification, destination-only bridge MOVE cleanup hashing, and zero-content-reread native cross-volume MOVE proof from stable source/destination snapshots around the successful copy.
- Findings 4, 20–21, 28–29, 38: Curl targeted size probe/unknown-size contract, aligned 1 MiB read hint/fill, EOF-tail correctness, and success-with-warning after post-replace cleanup failure.
- Findings 5, 23, 32, 34–35: scheduler help-execution, pause-aware pipeline teardown, real first-failure propagation, atomic perf counters, and callback-mutex option snapshots.
- Findings 6–8, 10, 12, 14, 17–18: MicrosoftDrive streaming expected-size writer and 8 MiB range request; S3/Graph revision pins; S3 probe caching, async multipart overlap, and explicit atomic final-key writer.
- Findings 11, 13, 30, 36–37: sleep-outside-lock throttling, high-metadata pre-scan suppression, full transfer-hint consumption, normalized transient HRESULT vocabulary, a process-wide 256 MiB host bridge-buffer budget, and a separate 256 MiB S3 multipart payload budget (512 MiB combined application-owned ceiling).
- Findings 16, 24, 26–27, 31: file-reparse policy parity, persistent 7z terminal errors including final-byte CRC synchronization, provider count clamps, and deterministic end-to-end Causeway fault coverage.
- Finding 9 was routed to the Compare owner; the normative decision that I/O failure is distinct from `Different` is now recorded in `Specs/Core/Core_CompareDirectories.md`.

The detailed text below is retained as the pre-remediation evidence and failure analysis; line numbers describe the reviewed pre-Causeway source, not current code.


## CRITICAL

### 1. Zip-Slip / path traversal: bridge tree-copy concatenates unsanitized plugin-supplied entry names into the local destination path

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8729`
- **Category:** path-traversal (security)
- **Verdict:** CONFIRMED
- **Defect:** In both CopyDirectorySequential (name join at :8729, isDot guard at :8723) and CopyDirectoryParallel (:9089, guard :9083) the bridge takes the raw entry name returned by the SOURCE plugin's ReadDirectoryInfo (via TryGetValidatedFileInfoName, which only bounds-checks the buffer, never the name content) and builds childDest = JoinFolderAndLeaf(destinationPath, name). The ONLY name filtering is `name.empty()` and exact-match `name==L"."||name==L".."` at :8723/:9083. Embedded path separators and multi-segment relative paths are not rejected. JoinFolderAndLeaf (:2407) simply appends the name after a separator. The destination temp/final path is then handed to the destination plugin; the local FileSystem plugin's ToExtendedPath (Plugins/FileSystem/FileSystem.Path.cpp:270) calls MakeAbsolutePath, which explicitly COLLAPSES `..` dot-segments (proven by its own self-test at FileSystem.Path.cpp:340-341: `child\..\leaf` -> `leaf`). Remote/archive backends supply attacker-controlled names that survive the guard: an S3 object key containing backslashes (FileSystemS3.S3.cpp:406 only strips a forward-slash '/', never '\' or '..'), e.g. key `..\..\..\Users\victim\evil.exe`, is returned verbatim as the leaf name; the host guard passes (name != exactly `..`), JoinFolderAndLeaf yields `<dest>\..\..\..\Users\victim\evil.exe`, MakeTempDestinationPath appends the CSPRNG suffix, the writer stages there and PromoteTempToFinalPath (:6990) MoveItems it to the traversed final path, which MakeAbsolutePath resolves outside the chosen destination folder. Directory (common-prefix) names take the same path and can create escaping directories via stack.emplace_back at :9117.
- **Failure scenario:** User copies/downloads an attacker-controlled S3 bucket (or crafted 7z/FTP/Drive listing whose entry name contains backslashes or '..' components) to a local folder. The bridge writes/overwrites files at arbitrary locations the user can write to (e.g. %USERPROFILE%, Startup folder, an existing binary), enabling remote code execution / file overwrite outside the destination directory. The host, which is the trust boundary between an untrusted remote namespace and the local disk, performs no separator/traversal sanitization.

### 2. Cross-FS directory copy joins untrusted source child names into destination paths without rejecting path separators or '..' (arbitrary file write / traversal)

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8728`
- **Category:** security (edge-cases)
- **Verdict:** CONFIRMED
- **Defect:** In CopyDirectorySequential the child name returned by the source plugin's directory enumeration (TryGetValidatedFileInfoName, which only bounds-checks the buffer, not the name content) is filtered solely against exactly L"." and L".." (line 8723) and then fed verbatim into JoinFolderAndLeaf(destinationPath, name) at 8728-8729 (and identically in CopyDirectoryParallel at 9083/9088-9089). Names containing '\', '/', or multi-segment '..\..\' sequences are NOT rejected. The joined destination path is passed to destinationIo.CreateFileWriter for the local plugin, where ToExtendedPath -> MakeAbsolutePath (Plugins/FileSystem/FileSystem.Path.cpp:288) COLLAPSES dot segments (confirmed by its own selftest at :341), so '..' escapes are actually resolved rather than blocked. Source plugins that expose attacker-controlled names include Curl FTP/HTTP: TryParseUnixListLine (Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp:869/893) sets entry.name to the raw remainder of the LIST line with no separator/traversal filtering.
- **Failure scenario:** User copies a directory from a malicious/compromised FTP or HTTP server (a supported cross-FS bridge source) to a local folder. The server returns a directory listing whose filename is '..\..\..\Users\<u>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\evil.bat'. The bridge stages and promotes attacker-controlled file content to that resolved absolute path outside the chosen destination, yielding arbitrary local file write and code execution on next logon. No prompt or error is surfaced for the traversal.

## HIGH

### 3. Cross-FS MOVE re-downloads every destination byte (and re-reads every source byte) for cleanup verify, serially

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:7491`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** Before deleting each moved source file, CopiedFileStillMatchesDestination (:7491) opens fresh readers on BOTH endpoints and runs ReadersHaveEqualContent (:7396) — a full byte-for-byte memcmp of the whole file. For a MOVE local→cloud this downloads the entire just-uploaded object back from the destination; for cloud→local it downloads the entire source a second time. The cleanup pass (DeleteCopiedSourceForMove :7818 → DeleteCopiedSourceEntryForMove :7569) is a single-threaded recursion per top-level item that shares the bridge's one primary buffer (:7442 reads into this->buffer), so verify runs at 1 stream even when the copy ran 8-way parallel, and each chunk issues the source Read and destination Read sequentially (:7442/:7459) with zero overlap. Verified in code; distinct from the known KS-1 (verify inconsistency) and Floodgate re-stat items, which are about correctness, not cost.
- **Failure scenario:** MOVE of a 100 GB tree from local disk to S3/OneDrive: 100 GB uploaded, then 100 GB downloaded again (cloud egress billed) through a single non-overlapped stream — total wall time is upload + a strictly serial dual-read compare, roughly 2-3x the copy time; users see the operation 'hang' after 100% copy progress.

### 4. Curl CreateFileReader ignores sizeKnown: GetSize returns S_OK+0 for sizeless-LIST servers, breaking every bridge copy/move from FTP dialects without a size column

- **Location:** `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp:1339`
- **Category:** contract-mismatch (contract-conformance)
- **Verdict:** CONFIRMED _(seen in a prior review)_
- **Defect:** FileSystemCurl::CreateFileReader constructs CurlStreamingReader(conn, path, entry.sizeBytes) from a GetEntryInfo LIST probe, discarding entry.sizeKnown (FileSystemCurl.h:50, set false at Shared.cpp:892 when the listing dialect has no parseable size). CurlStreamingReader::GetSize (DirectoryOps.cpp:416-425) then unconditionally returns S_OK with the cached value, i.e. S_OK+0 for an unknown size. The host reserves GetSize FAILURE as the only unknown-size signal (State.cpp:8098-8114 refuses to copy) and treats S_OK+0 as an authoritative 0-byte file. The 2026-06-20 per-file SIZE probe (CurlProbeRemoteFileSize, Shared.cpp:2879) was wired only into the native staged-copy path (CopyMove.cpp:1275,1293), never into the IFileSystemIO reader path the bridge uses. Curl capabilities allow full bridge import/export (FileSystemCurl.h:246-249), so this is a mainline path.
- **Failure scenario:** Bridge copy or move from an FTP/FTPS server whose LIST output has no parseable size: reader GetSize returns 0, the pump streams N>0 bytes, host integrity gate at State.cpp:8485 fires ERROR_PARTIAL_COPY (bridge.integrity.sizeMismatch) and deletes the temp — every non-empty file fails. Same defect on the destination side falsifies the MOVE post-promote re-stat (State.cpp:8547-8557): destinationSizeBytes=0 != fileTotalBytes, so cross-FS MOVE onto such a server always ends ERROR_PARTIAL_COPY with source preserved. Fail-safe direction but a hard functional break of the FTP bridge.

### 5. Nested scheduler job + blocking WaitJob can deadlock all file operations process-wide

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:9165`
- **Category:** deadlock (concurrency)
- **Verdict:** CONFIRMED
- **Defect:** CopyDirectoryParallel runs its producer inside a per-item processIndex lambda that itself executes on a PerItemTaskScheduler worker (StartJob :9897), then starts a NESTED job for its file-copy workers (StartJob :8950) and blocks in scheduler.WaitJob(job) (:9165). WaitJob (:3882-3891) is a bare condition_variable wait on job->done with no cancellation predicate and no help-execute/work-stealing. Workers are global and capped at min(hardware_concurrency,16) (:4039-4046). The fairness cap in effectiveMaxConcurrencyLocked (:3972-4029) only limits NEW dequeues; it does not prevent already-running producers from occupying every worker. If N concurrently active per-item producers (across tasks; e.g. 4 concurrent cross-FS directory copy tasks on a 4-core machine, or 2 tasks x 2 directory items) occupy all workers, their nested jobs can never be dequeued: producers block forever in enqueueWork backpressure (:8970-8974, wakes on cancel but then falls through to WaitJob) or directly in WaitJob. Cancel does not recover: cancelled jobs are only finished by cleanupJobsLocked (:4098-4133), which runs only on a free worker's dequeue path — and none are free. Even Shutdown() joins workers BEFORE finishing jobs (:3928-3930), so app exit hangs too.
- **Failure scenario:** On a 4-core machine (4 scheduler workers), user starts 4 concurrent cross-FS copy/move tasks each on one large directory (or 2 tasks with 2 top-level directories each). All 4 workers run producers; each producer starts a nested worker job and blocks in WaitJob/enqueueWork. Nested jobs never get a worker. All file operations in the process hang permanently; Cancel does not release them; exiting the app deadlocks in scheduler Shutdown. Requires process kill.

### 6. OneDrive writer stages the whole file to local %TEMP% and uploads only at Commit — zero read/upload overlap

- **Location:** `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:4632`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** MicrosoftDriveFileWriter::Write is just WriteFile to a local delete-on-close temp file (:4653-4666); nothing touches Graph until Commit (:6279) which then does the simple PUT (≤250 MiB) or chunked upload session. The bridge's pipelined pump and 2-slot overlap are completely defeated for OneDrive destinations: the entire source is first pumped into the local temp, THEN the upload starts, so wall time = full source read + full upload with no overlap, plus a full extra disk write + read of every byte and transient free-disk requirement equal to the file size. Contrast with S3's writer which at least starts flushing 64 MiB parts during Write.
- **Failure scenario:** Copying a 20 GB file FTP→OneDrive: 20 GB downloaded to %TEMP% (fails outright if the system drive has <20 GB free), and only after that does the multi-hour Graph upload begin — total time is the sum of both legs instead of max(), nearly doubling wall time; progress reaches 100% while the upload has not started a single byte.

### 7. OneDrive reader caps every HTTPS range GET at 1 MiB, issued synchronously back-to-back

- **Location:** `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:4488`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** MicrosoftDriveRangedFileReader::FetchRange clamps chunkBytes to at most 1 MiB (:4487-4488) regardless of how much the caller requested, and Read (:4391-4441) loops issuing these fetches synchronously — each one a full SendHttpRequest that buffers the entire response body and then memcpy's it into _cache (:4549-4553). An 8 MiB bridge Read (the hint-resolved buffer size MSDrive itself advertises at :5741) therefore becomes 8 sequential request/response cycles with no prefetch, no pipelining, and no parallel ranges; the bridge's reader jthread gains nothing because each Read call is internally serial.
- **Failure scenario:** Downloading a 2 GB file from OneDrive over a 100 ms-RTT link: ~2048 sequential ranged GETs, each paying RTT+TTFB before any body bytes flow — effective throughput collapses to roughly 1 MiB per (RTT + server time), often <10 MB/s on a gigabit line, and the plugin's own preferredBufferBytes=8MiB hint is unachievable.

### 8. Small-file cloud pathology: 10-15+ HTTP round trips per file on an S3 destination, no batching

- **Location:** `Plugins/FileSystemS3/FileSystemS3.IO.cpp:833`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** Per bridged file to S3 the fixed overhead is: host ValidateDestinationOverwritePolicy → destinationIo.GetAttributes (State.cpp:7989, a HEAD/LIST); CreateFileWriter(temp) → EnsureWritableS3Target (IO.cpp:627-698) = TryGetS3ObjectSummary for the key + one summary probe PER ancestor path component (:654-669) + a ListObjectsV2 prefix probe (:683); Commit RE-RUNS the entire EnsureWritableS3Target (:833) before PutObject; then host promote (PromoteTempToFinalPath State.cpp:6990) calls S3 MoveItem = server-side CopyObject + delete of the temp key; MOVE adds a post-promote CreateFileReader+GetSize HEAD (State.cpp:8547-8551). A small file at depth 5 costs ~15+ S3 API calls where a plain PutObject is 1. There is no batch path — the bridge is strictly per-file CopyFileWithBuffer, and concurrency is capped at min(source,dest) preferred = 8.
- **Failure scenario:** Copying 50,000 small files (node_modules-like tree) local→S3: ~750k sequentialized-per-worker HTTP requests at WAN latency (~50 ms each / 8 workers) ≈ >1 hour dominated by request overhead, plus 3-5x S3 request billing versus a direct uploader; transfer of the actual bytes is a rounding error.

## MEDIUM

### 9. CompareDirectoriesEngine maps reader-create and Read failures to 'Different', contradicting the bridge's fail-hard contract stance

- **Location:** `RedSalamander/CompareDirectoriesEngine.cpp:3010`
- **Category:** contract-mismatch (contract-conformance)
- **Verdict:** CONFIRMED
- **Defect:** CompareFileContent returns FileContentCompareResult::Different when CreateFileReader fails on either side (:2930-2933, :2937-2940) and when IFileReader::Read returns any FAILED hr mid-pump (tryRead :3010-3013 returns false, mapped to Different at :3039-3046). No warning or error is surfaced. The bridge treats the identical conditions as operation failures (State.cpp:8071, :8194-8197), and SearchFallbackEngine degrades to explicit warnings — three consumers of the same 23-line IFileReader contract with three contradictory failure semantics, proving the contract ambiguity. Curl maps transient send/recv errors to ERROR_CONNECTION_ABORTED and S3 to ERROR_UNEXP_NET_ERR, both of which compare renders as a content difference.
- **Failure scenario:** User compares a local tree against an S3/FTP tree during a brief network blip: files that are byte-identical are reported 'Different'. Acting on that verdict via the compare-driven sync/copy actions triggers unnecessary overwrites of identical destination files (and with the previously-reviewed sync-move semantics, source deletion decisions) based on a transient read fault presented as a definitive content difference.

### 10. S3 and MSDrive ranged readers pin no object version across chunked GETs; a concurrent same-size replacement is committed as a silently mixed-version file

- **Location:** `Plugins/FileSystemS3/FileSystemS3.IO.cpp:601`
- **Category:** data-integrity (contract-conformance)
- **Verdict:** CONFIRMED
- **Defect:** S3RangedFileReader::FillBufferFrom (IO.cpp:530-601) issues an independent Range GET per 8 MiB chunk with no If-Match/ETag/versionId; validation (ValidateS3RangeResponseLength :601, :74) is length-only. MicrosoftDriveRangedFileReader has the same shape (FetchRange, FileSystemMicrosoftDrive.cpp:4491-4554) and additionally RefreshDownloadUrl on a mid-stream 401/403 (:4518-4521) re-resolves metadata and OVERWRITES _sizeBytes/_downloadUrl (:4465-4466), so later chunks come from the NEW item content while earlier chunks came from the old. The IFileReader contract (Common/PlugInterfaces/FileSystem.h:530-539) says nothing about GetSize stability or content-version pinning, and the host's only integrity gate is total byte count vs the start-of-copy GetSize snapshot (State.cpp:8485).
- **Failure scenario:** Object at the cloud source is replaced with different same-size content while the bridge is mid-file (large file, many chunks): the pump stitches old-version and new-version chunks, byte count matches fileTotalBytes, Commit+promote succeed — a corrupted mixed-version file is silently committed on plain COPY. On MOVE the cleanup byte-compare (State.cpp:7491-7547) detects the mismatch and preserves the source, but the corrupt promoted destination file remains in place.

### 11. Bandwidth throttle sleeps while holding throttleMutex, serializing all pump workers behind the sleeper

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:7075`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** ThrottleThreadSafe (:7068-7077) takes scoped_lock(throttleMutex) and then calls Throttle, which sleeps in 50 ms slices for the full pacing duration (:7057-7064 via SleepResponsive :7012) with the mutex held. When a speed limit is active, every other parallel worker's write loop blocks inside ThrottleThreadSafe for the whole sleep rather than computing its own residual, so per-stream progress updates stall in lockstep and pacing degenerates to convoy behavior. The pacing formula is also cumulative-average based (bytesSoFar*1000/limit vs elapsed since bridge start, :7042-7055), so raising the limit mid-transfer releases a large instantaneous burst and lowering it causes a long full stall until the average catches down.
- **Failure scenario:** 8 workers with a 10 MB/s cap: one worker sleeps 400 ms holding the lock while 7 others sit blocked on the mutex instead of overlapping their reads; user raises the cap mid-copy and the transfer bursts unthrottled for many seconds (or lowers it and all streams freeze) because the cumulative average must re-converge.

### 12. S3 writer uploads 64 MiB parts synchronously inside Write, one part at a time, stalling the pump

- **Location:** `Plugins/FileSystemS3/FileSystemS3.IO.cpp:947`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** MultipartS3FileWriter::Write appends to a RAM vector and FlushBufferedParts (:947-965) synchronously calls UploadPart for each accumulated 64 MiB (kMultipartMinPartSizeBytes) part before returning. There is no concurrent part upload despite S3 multipart being designed for parallel parts and the plugin using the s3-crt client. Because the host pipeline can only buffer bufferBytes (8-16 MiB) ahead in its 2 slots, the source reader is stalled for essentially the entire duration of every 64 MiB part PUT; effective upload throughput = single-stream part-at-a-time, typically a small fraction of available bandwidth on high-BDP links.
- **Failure scenario:** Uploading a 50 GB file local→S3 over a 1 Gbps/60 ms link: ~800 strictly sequential 64 MiB PUTs, each preceded by a pump stall; achieved throughput settles near single-TCP-stream limits (~10-30 MB/s) versus 100+ MB/s achievable with 4-8 concurrent parts — hours instead of minutes.

### 13. Source tree enumerated 2x for cross-FS COPY and 3x for MOVE on high-metadata-cost cloud backends

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:5116`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** RunPreCalculation (:5112-5132) runs for every COPY/MOVE when the setting is on, calling the source plugin's GetDirectorySize (full recursive LIST — implemented by S3/Curl/MSDrive per their DirectoryOps files) plus its own per-directory ReadDirectoryInfo child enumeration for work distribution (:5281-5347). The bridge then re-enumerates the identical tree via ReadDirectoryInfo in CopyDirectorySequential/Parallel (:8649/:8807). A MOVE re-lists every directory a third time during cleanup (DeleteCopiedSourceEntryForMove directory branch). The plugins all set FILESYSTEM_TRANSFER_FLAG_HIGH_METADATA_COST in their hints, but no code path reads that flag to skip or reuse the pre-calc enumeration; the '5F early admission' overlap also means pre-calc LIST traffic competes with the copy's LIST/GET traffic on the same connection budget.
- **Failure scenario:** MOVE of a 100k-entry directory tree off an FTP or S3 source: ~300k LIST-class round trips instead of ~100k; on a 50 ms-latency backend the redundant enumerations alone add tens of minutes and can dominate total job time for small-file trees.

### 14. S3 promote performs a full server-side copy of every uploaded object (temp key → final key) plus delete

- **Location:** `Plugins/FileSystemS3/FileSystemS3.S3.cpp:845`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** The bridge always stages to a .rs_tmp_ key and promotes via destinationFs.MoveItem (State.cpp:6990). S3 has no rename: MoveItem → ExecuteCopyOrMove → CopyS3ObjectServerSide (:845) which is CopyObject for <64 MiB or a full multipart server-side copy session (:884+) for larger objects, followed by deleting the temp object. Every byte uploaded to S3 is therefore written twice server-side (PUT storage ops billed twice, transiently double storage), and per-file completion is delayed by the copy — for multi-GB objects a multipart copy adds substantial per-file latency after the upload already finished. Graph/FTP promotes are metadata renames; this cost is S3-specific and unavoidable only because the staging happens remotely rather than committing directly to the final key with S3's already-atomic PUT/CompleteMultipartUpload semantics (the destination object is invisible until Complete anyway, per EnsureWritableS3Target + MPU design, making the temp-key hop redundant for S3).
- **Failure scenario:** Uploading a 200 GB file to S3: after the multipart upload completes, the file sits at the temp key while a full 200 GB server-side multipart copy runs (minutes, additional request costs), then a delete; a task of 10k small files pays 2 extra mutation requests per file on top of finding #2.

### 15. Names crossing from case-sensitive/foreign namespaces are not sanitized for NTFS-illegal characters, causing ADS injection or silent case-collision overwrite

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8729`
- **Category:** correctness (edge-cases)
- **Verdict:** CONFIRMED
- **Defect:** The bridge builds childDest = JoinFolderAndLeaf(destinationPath, name) with no validation of NTFS-illegal characters (':','*','?','"','<','>','|'), trailing dots/spaces, reserved device names, or case. A source key containing ':' (legal on S3, 7z, Linux FTP) becomes an NTFS alternate-data-stream write (e.g. 'file.txt:hidden' writes into stream of 'file.txt') because ToExtendedPath prefixes \\?\ and passes the colon through. Separately, two distinct source entries differing only in case (e.g. 'README' and 'readme' from a case-sensitive backend) both map to the same NTFS destination: the second file's ValidateDestinationOverwritePolicy (State.cpp:7980) sees the first as an existing file, and under an Overwrite/overwrite-all grant the two source files collapse into one destination file.
- **Failure scenario:** Copying a case-sensitive directory (S3 bucket, tar/7z listing, or Linux FTP) that contains both 'Data' and 'data' to an NTFS folder with overwrite-all selected silently discards one of the two distinct files (data loss); a source name like 'report.txt:$secret' writes attacker content into an alternate data stream instead of the intended visible file.

### 16. Reparse-point FILES bypass ReparsePointPolicy in the bridge tree walk

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8758`
- **Category:** correctness (edge-cases)
- **Verdict:** CONFIRMED
- **Defect:** In CopyDirectorySequential the reparse flag `isReparse` (computed at 8727) is only consulted inside the `if (isDirectory)` branch (8733-8752); for non-directory entries the code unconditionally falls through to CopyFile at 8760, ignoring reparsePointPolicy entirely. CopyDirectoryParallel has the same structure (isReparse at 9087 used only under isDirectory at 9093; files enqueued at 9122 regardless). Thus file symlinks / reparse-point files are always followed and their target content copied even when the policy is ReparsePointPolicy::Skip. For a cross-FS MOVE, ShouldDeleteMoveSourceAfterBridgeCopy is not blocked by a file-reparse skip (only directory-reparse skips increment skippedDirectoryReparseCount), so the source reparse-point file is subsequently deleted after its target's bytes were copied to the destination as a plain file.
- **Failure scenario:** User moves a tree containing a file symlink with ReparsePointPolicy=Skip expecting links to be left untouched; instead the bridge copies the link target's content to the destination as a regular file and deletes the source symlink, changing on-disk semantics (link replaced by a materialized file / broken relationship) without any skip diagnostic.

### 17. MSDrive ranged reader mixes file versions across chunks (no ETag pinning, stale size gates EOF)

- **Location:** `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:4410`
- **Category:** silent-corruption (data-integrity)
- **Verdict:** CONFIRMED
- **Defect:** MicrosoftDriveRangedFileReader reads until _position < _sizeBytes where _sizeBytes is the metadata snapshot at reader creation, refreshed only on HTTP 401/403 (:4518, RefreshDownloadUrl :4447). FetchRange (:4477-4554) issues independent ranged GETs against the cached pre-authenticated downloadUrl with no ETag/If-Match/If-Range validation and no cross-chunk version check. If the remote file is replaced mid-copy, later ranges serve the new version: the bridge receives a byte stream that is a splice of two versions, or the new (longer) file truncated at the old _sizeBytes. The stream length still equals the stale GetSize value, so the host integrity gate (State.cpp:8485) passes, Commit and promote succeed, and plain COPY commits a corrupt destination with no diagnostic.
- **Failure scenario:** User bridge-copies a OneDrive file to local disk while another client saves a new revision of the same file; chunk N is fetched pre-save and chunk N+1 post-save -> destination is a torn mix of both revisions, reported as a fully successful copy. MOVE is protected (post-promote re-stat + byte-compare at delete), COPY is not.

### 18. S3 ranged reader issues per-chunk GetObject without IfMatch/versionId (torn copy under concurrent overwrite)

- **Location:** `Plugins/FileSystemS3/FileSystemS3.IO.cpp:567`
- **Category:** silent-corruption (data-integrity)
- **Verdict:** CONFIRMED
- **Defect:** S3RangedFileReader::FillBufferFrom builds GetObjectRequest with only Bucket/Key/Range (:567-572); no IfMatch on the create-time ETag and no versionId pin. Size is HEAD-cached once in EnsureSizeKnown (:504-528). ValidateS3RangeResponseLength (:74) only enforces that each individual range response matches its requested length, which catches size-changing replacement (leading to a fail-closed ERROR_PARTIAL_COPY) but NOT a same-size replacement or a replacement occurring between 8 MiB chunk fetches while remaining ranges are still in-bounds. Successive chunks can then come from different object versions.
- **Failure scenario:** Bridge COPY from a shared S3 bucket while another writer PutObject-replaces the same key with equal-size content -> chunks 0..k from the old version, k+1.. from the new -> host size gate passes, torn file committed as a successful copy. As with MSDrive, MOVE is caught by the cleanup byte-compare; COPY is silent.

### 19. Plain COPY still has no post-promote re-stat or content verification (Floodgate item, partially addressed)

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8544`
- **Category:** verification-gap (data-integrity)
- **Verdict:** CONFIRMED _(seen in a prior review)_
- **Defect:** The post-promote destination re-stat added since the Floodgate review is gated on task._operation == FILESYSTEM_MOVE (:8544-8576); the full byte-compare (:7396/:7491) runs only in the MOVE source-delete pass. Bridge COPY trusts: writer Write() return counts, writer Commit(), and destination MoveItem promote — with zero independent destination verification. This is the enabler that turns the Curl-Commit-truncation and the two ranged-reader torn-copy findings into SILENT corruption instead of detected failures. Verified still present; MOVE-side verification is now solid (re-stat + manifest + basic-info + byte-for-byte compare before any source delete), so the old KS-1 'S3 MOVE verifies nothing before source delete' concern is closed in the bridge path.
- **Failure scenario:** Any destination plugin that acknowledges writes/Commit but persists fewer or different bytes (see findings 1-3) produces a destination that differs from the source while the COPY task reports full success.

### 20. Curl sizeless-LIST hole NOT closed for the bridge reader path: GetSize returns S_OK with wrong/zero size

- **Location:** `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp:1339`
- **Category:** correctness (concurrency)
- **Verdict:** CONFIRMED _(seen in a prior review)_
- **Defect:** FileSystemCurl::CreateFileReader passes entry.sizeBytes into CurlStreamingReader ignoring the entry.sizeKnown flag (FileSystemCurl.h:50), and CurlStreamingReader::GetSize (:416-425) unconditionally returns S_OK with that cached value (0 when the LIST dialect reported no size). The per-file SIZE/stat probe added 2026-06-20 (CurlProbeRemoteFileSize, Shared.cpp:2879) is wired only into the plugin's internal staged-upload verify (CopyMove.cpp:1275) — not into CreateFileReader. The host bridge treats a successful GetSize as authoritative (State.cpp:8098-8102) and enforces fileCompletedBytes==fileTotalBytes (State.cpp:8485), so a 0 (or stale) reported size makes every non-empty file fail with ERROR_PARTIAL_COPY after the bytes were fully transferred, and the temp is deleted.
- **Failure scenario:** Cross-FS copy or move FROM an FTP/SFTP server whose listing dialect omits file sizes: each file streams completely, then fails 'bridge.integrity.sizeMismatch' (expected 0, wrote N); the whole tree ends ERROR_PARTIAL_COPY. Safe direction (MOVE preserves source, no data loss) but FTP-source bridge transfers are entirely broken on such servers, wasting full-transfer bandwidth per file.

### 21. Curl sizeless/stale-LIST hole not closed for the bridge reader path (SIZE probe not wired into CreateFileReader)

- **Location:** `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp:1339`
- **Category:** reliability (data-integrity)
- **Verdict:** CONFIRMED _(seen in a prior review)_
- **Defect:** CreateFileReader sizes CurlStreamingReader from GetEntryInfo (:1326-1339), which stats the file by parsing the parent directory LIST (Imap.cpp:3644-3690). The per-file SIZE probe added 2026-06-20 (CurlProbeRemoteFileSize, Shared.cpp:2879, with an explicit sizeKnownOut flag) is only used by the plugin's native CopyFileViaTemp staged-upload verification (CopyMove.cpp:1275) — the bridge reader never benefits. Reader GetSize (:416-425) unconditionally returns S_OK with the LIST-derived value, so (a) on servers whose LIST omits sizes the bridge sees fileTotalBytes=0 and every non-empty file fails the :8485 size gate with ERROR_PARTIAL_COPY (fail-closed: no silent loss, but cross-FS copy from such FTP servers is 100% broken), and (b) the host's bridge.integrity.sourceSizeUnknown fail-fast (:8098-8114) can never trigger for Curl because GetSize cannot fail, producing a misleading late sizeMismatch diagnostic instead.
- **Failure scenario:** Bridge COPY/MOVE from an FTP server whose LIST lines lack a size field (or report stale sizes): every file copy uploads fully, then is discarded with 'File copy size mismatch: expected 0 bytes...'; MOVE preserves the source. Availability/diagnostic defect, deliberately fail-closed.

### 22. Cross-FS COPY can silently truncate when a cloud reader's cached GetSize is stale (OneDrive)

- **Location:** `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp:4410`
- **Category:** correctness (edge-cases)
- **Verdict:** CONFIRMED _(seen in a prior review)_
- **Defect:** MicrosoftDriveRangedFileReader::Read loops `while (totalRead < bytesToRead && _position < _sizeBytes)` where _sizeBytes is the object size captured from Graph metadata at reader creation and only refreshed on a 401/403. The bridge's ONLY plain-COPY integrity check (State.cpp:8485) compares fileCompletedBytes to fileTotalBytes, but fileTotalBytes comes from reader->GetSize (State.cpp:8098) which returns the same cached _sizeBytes (GetSize at :4350 just echoes _sizeBytes). Because both the read bound and the integrity target derive from the identical stale value, a source object that grows after the reader is opened is copied only up to the old size, the byte-count gate passes, and Commit/promote succeed with no error. Plain COPY performs no post-commit content or re-stat verification (only MOVE does, at :8544).
- **Failure scenario:** A OneDrive file is replaced/extended by another client between when the bridge opens the reader (metadata fetch) and when streaming completes. The cross-FS COPY writes only the original _sizeBytes prefix, reports success, and the destination is a silently truncated copy of the now-larger source.

## LOW

### 23. Pipelined pump shutdown join can be blocked indefinitely by a paused/prompting reader thread or a long plugin Read

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8472`
- **Category:** liveness (concurrency)
- **Verdict:** CONFIRMED
- **Defect:** After a write failure or cancellation the writer sets pipelineStop and joins the reader jthread (:8470-8472). The reader may be blocked in task.WaitWhilePaused() (:8296) — which wakes only on unpause/cancel, not on pipelineStop — or inside a plugin Read with no host-side cancellation (CurlStreamingReader::Read blocks on its ring CV until data/error; worst case bounded only by curl's 1B/s-60s low-speed abort; 7z extract-pipe Read blocks until the extract thread produces data). A write-fault failure occurring while the task is paused, or while another worker owns a conflict prompt, cannot surface until the pause/prompt clears; cancel latency is bounded by one full plugin Read.
- **Failure scenario:** User pauses a cross-FS transfer; the destination writer then fails (disk full on temp). The writer thread sits in join() because the reader is parked in WaitWhilePaused; the per-file failure, temp cleanup and error reporting are all deferred until the user resumes. No deadlock (unpause/cancel releases it), but the operation appears hung and error handling is delayed arbitrarily.

### 24. 7z reader swallows a mid-item decoder failure after reporting it once: subsequent Reads at EOF position return clean S_OK+0

- **Location:** `Plugins/FileSystem7z/FileSystem7z.cpp:2991`
- **Category:** contract-mismatch (contract-conformance)
- **Verdict:** CONFIRMED
- **Defect:** SevenZipItemFileReader::Read defers a decoder/extract failure into _terminalReadStatus and reports it exactly once when _positionBytes >= _fileSizeBytes (:2989-2996, guarded by _terminalStatusReported); every subsequent Read at that position returns S_OK with 0 bytes — indistinguishable from a clean EOF. Additionally the spool path can return S_OK+0 mid-file when the spool lags the requested position (canTake==0 at :3013-3016), violating the header's only normative Read rule (FileSystem.h:534: S_OK+0 only at EOF). IFileReader has no documented post-failure state, so any consumer that retries a failed Read (FileSystemIssueAction::Retry semantics are undefined for streams) observes a clean truncated file.
- **Failure scenario:** A consumer that re-reads after a reported failure (retry logic, or a second verification pass over the same reader instance) sees S_OK+0 and treats the truncated decode as a complete file. Current bridge/compare consumers fail closed on the first error, so today's impact is contained; the hazard is latent for any retry-style consumer.

### 25. Plain COPY has no post-promote destination verification: re-stat is gated to MOVE only, leaving the Floodgate eventual-consistency item open for COPY

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8544`
- **Category:** verification-gap (contract-conformance)
- **Verdict:** CONFIRMED _(seen in a prior review)_
- **Defect:** The post-promote destination re-stat (fresh CreateFileReader+GetSize vs fileTotalBytes) executes only when task._operation == FILESYSTEM_MOVE (:8544-8576). A cross-FS COPY declares per-file success from pumped-byte count (:8485), writer Commit (:8512) and PromoteTempToFinalPath returning S_OK (:8529) — the destination is never read back. Success therefore rests entirely on each destination plugin's Commit/MoveItem truthfulness: for Curl the promote is an FTP RNFR/RNTO quote command sequence whose success is taken at protocol face value; nothing detects a server that reports rename success but stored a truncated STOR payload, since the Curl writer's Commit (DirectoryOps.cpp:1004-1032) also verifies no byte count against the server.
- **Failure scenario:** COPY of a tree to a cloud/FTP backend where the writer's Commit or the promote rename succeeds per-protocol but the stored object is short/absent (proxy or server bug, eventual-consistency window): the host reports success, the user may then delete the source manually trusting the copy. MOVE is protected by the re-stat + byte-compare; COPY has zero read-back.

### 26. Host bridge pump never clamps plugin-reported bytesRead to the buffer size, allowing out-of-bounds heap over-read copied into the destination file

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8192`
- **Category:** memory-safety (security)
- **Verdict:** CONFIRMED
- **Defect:** Both pump paths trust the plugin-controlled `bytesRead` out-param from IFileReader::Read without validating `bytesRead <= bufferBytesIn`. Serial pump: Read(bufferIn, bufferBytesIn, &bytesRead) at :8192, then the write loop iterates `while (offset < bytesRead)` writing from `bufferIn + offset` (:8205-8216). Pipeline pump: Read(slots[readIndex].buffer, bufferBytesIn, &bytesRead) at :8316, writer loops `while (offset < slotBytesRead)` from `slots[writeIndex].buffer + offset` (:8400-8412). If a plugin returns bytesRead greater than bufferBytesIn (the contract in Common/PlugInterfaces/FileSystem.h:534 pins only the bytesRead==0 EOF case and never forbids bytesRead>bytesToRead), the host reads past the end of its heap-allocated buffer (bufferIn is sized bufferBytesIn at ctor; secondaryBuffer is `new std::byte[bufferBytesIn]` at :8260) and streams that adjacent heap memory into the destination file via writer->Write. The overflow-accounting checks at :8229/:8427 guard only the uint64 completed-bytes counters, not the buffer bound.
- **Failure scenario:** A malicious or buggy third-party IFileSystem/IFileReader plugin (the plugin ABI is COM-style and third-party-loadable) returns bytesRead > bufferBytesIn from Read. The host over-reads its own heap and copies leaked heap contents (potentially credentials/tokens held by other plugins in-process) into the copied file, and/or crashes. No host-side clamp or rejection exists.

### 27. 7z reader can return S_OK with bytesRead==0 before EOF, violating the Read contract

- **Location:** `Plugins/FileSystem7z/FileSystem7z.cpp:3013`
- **Category:** contract-violation (data-integrity)
- **Verdict:** CONFIRMED
- **Defect:** In the in-memory spool path, after EnsureSpooledUntil succeeds, if the spooled watermark still does not cover the current position (canTake == 0), Read returns S_OK with *bytesRead == 0 while _positionBytes < _fileSizeBytes (:3011-3016). FileSystem.h:534 reserves S_OK+0 strictly for EOF/zero-length requests. Both bridge pumps (State.cpp:8199, :8323) treat 0 as EOF, break, and then fail the file at the :8485 size gate. Reachable when the background extract worker terminates early without setting a failure status for the requested range.
- **Failure scenario:** Bridge COPY out of a 7z archive hits the spool short-return -> pump ends early -> ERROR_PARTIAL_COPY, temp deleted, source kept. Fail-closed (spurious failure, no data loss), but any future Read consumer that trusts the documented EOF rule less defensively would silently truncate.

### 28. Curl advertises 8 MiB preferredBufferBytes but its reader ring holds only 1 MiB

- **Location:** `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp:354`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** CurlStreamingReader allocates a fixed 1 MiB ring (:354-360) that the background curl worker fills, while the plugin's GetTransferHints (Shared.cpp:4474) advertises preferredBufferBytes = 8 MiB — which the host dutifully adopts as the bridge buffer via the max-merge (State.cpp:2083). Every 8 MiB bridge Read can return at most the ~1 MiB currently buffered (short reads are routine), so the 8 MiB host buffers (x2 slots x workers) are 87% dead weight, prefetch depth against FTP stalls is capped at 1 MiB, and the pump makes 8x more Read iterations (plus progress/throttle checks) than the hint implies.
- **Failure scenario:** FTP→local copy of large files over a jittery WAN: the 1 MiB ring drains during any server hiccup and the pump idles despite 16 MiB of allocated-but-unfillable host buffer; per-task RAM is inflated ~8x for zero prefetch benefit.

### 29. Overwrite-promote reports failure after the destination was already replaced when backup cleanup fails (Curl)

- **Location:** `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp:1061`
- **Category:** false-failure (data-integrity)
- **Verdict:** CONFIRMED
- **Defect:** RenameWithOverwriteRollback (used for the bridge's temp->final promote via Curl MoveItem :2309) renames the pre-existing destination to a .rs-rollback sibling, renames temp into place, then FinalizeOverwriteTarget deletes the backup (:1061 -> :1032-1040). If only that final backup DELE fails, the whole promote returns failure although the rename already succeeded: the host marks the file failed (MOVE preserves the source -> duplicate), the destination actually holds the new content, and an orphan rollback file is left next to it. The equivalent S3 path has the same shape (CleanupBackupObjects failure -> ERROR_PARTIAL_COPY after a fully successful transfer, FileSystemS3.Directory.cpp:1689-1693). No data loss — old content survives in the backup — but the reported result contradicts the destination state and leaks staging artifacts.
- **Failure scenario:** Bridge COPY with granted overwrite onto FTP; DELE of the rollback backup times out -> file reported failed, user retries, gets a fresh overwrite prompt, and a stray .rs-...-rollback file remains on the server.

### 30. Transfer hints dead except preferredBufferBytes, and any non-default user buffer setting silently disables adaptation

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:2058`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** ResolveAdaptiveCrossFsBridgeBufferBytes returns the configured value verbatim whenever it differs from the 4096 KB default (:2058) — a user who once set 512 KB (or an old default migrated forward) permanently pins every cloud transfer to that size with no indication, since the check is exact-equality against the default rather than an explicit 'auto' sentinel. Of the hint struct (FileSystem.h:103-111) only preferredBufferBytes is ever read (:2070); latencyClass, PREFERS_LARGE_BUFFERS/PREFERS_SEQUENTIAL_IO/HIGH_METADATA_COST flags, and preferredProgressPeriodMs are never consulted anywhere in the engine — progress cadence is hardcoded 200 ms (ProgressIntervalMs :6800) and the pipeline/enumeration decisions ignore latency class (see findings 5 and 7). Hints are also resolved once per top-level-item bridge ctor from root paths only.
- **Failure scenario:** User sets crossFsBridgeBufferSizeKB=512 while tuning local copies, later copies to S3/OneDrive: every cloud chunk is 512 KB (16x more round-trip-paced chunks than the plugins' 8 MiB request), throughput drops several-fold with no warning; plugins' carefully chosen hints (2 MiB local / 8 MiB cloud) are ignored.

### 31. No end-to-end self-test exercises the bridge size-integrity gate via short reads or writer under-consumption

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8485`
- **Category:** test-coverage (data-integrity)
- **Verdict:** CONFIRMED
- **Defect:** The only bridge IO fault injector, SelfTestBridgeFileReader (:307-509), can fail GetSize; Read/Write are pure pass-throughs. Nothing injects a premature-EOF/short source reader or a destination writer that accepts bytes but persists fewer, so the fileCompletedBytes != fileTotalBytes gate (:8485) and the writer-truncation class (finding 1) are never exercised end-to-end, and Commit()/PromoteTempToFinalPath() failure paths have no injection hooks either. CompareDirectories already has the needed pattern (ShortReadFileReader, CompareDirectoriesEngine.SelfTest.cpp:1756). Note: new fileops cases must be registered in kFileOpsFamilyDefinitions or Full runs skip them silently.
- **Failure scenario:** A regression that drops or reorders the size gate, or a writer that lies about bytesWritten, would ship undetected by the Full suite.

### 32. Genuine worker failure can be misreported as ERROR_CANCELLED in CopyDirectoryParallel

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8880`
- **Category:** race (concurrency)
- **Verdict:** CONFIRMED
- **Defect:** recordFailure stores stopRequested=true (:8880) BEFORE CAS-ing the real failure HRESULT into firstFailure (:8881-8882). A producer blocked in enqueueWork observes stopRequested and returns ERROR_CANCELLED (:8976-8978); recordProducerFailure then wins the firstFailure CAS with ERROR_CANCELLED (:8952-8963) before the worker's CAS of the real error executes. CopyDirectoryParallel maps firstFailure==ERROR_CANCELLED to ERROR_CANCELLED (:9184-9187) even though CancelRequested() is false, and the per-item loop returns ERROR_CANCELLED for the whole task (:9695-9698).
- **Failure scenario:** During a parallel cross-FS directory copy, a worker hits disk-full/access-denied while the producer is parked in admission-queue backpressure. The task silently ends as 'Cancelled': the real error is never classified into a conflict bucket, no retry/skip prompt is shown, and diagnostics record a user cancel that never happened. For MOVE the source is preserved (no data loss), but the failure cause is lost.

### 33. Plain COPY still has no post-Commit/post-promote destination verification (MOVE-only re-stat)

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8544`
- **Category:** verification-gap (concurrency)
- **Verdict:** CONFIRMED _(seen in a prior review)_
- **Defect:** The post-promote destination re-stat (fresh CreateFileReader + GetSize, :8544-8576) and the byte-for-byte compare before source delete (:7491/:7396) run only for FILESYSTEM_MOVE. For COPY the only integrity gates are the pre-copy GetSize (:8098), the written-bytes==size check (:8485) and writer Commit (:8512) — nothing re-checks the promoted destination object. Note: the previously-open KS-1 (S3 MOVE destination verified nothing before source delete) now appears closed by the MOVE re-stat + ReadersHaveEqualContent gate; only the COPY half of the Floodgate re-stat item remains.
- **Failure scenario:** Cross-FS COPY to a cloud backend where Commit/promote succeeds but the final object is truncated or replaced by a concurrent writer between promote and completion: the task reports success with no re-stat. Because it is a COPY the source still exists, so impact is a false success indication rather than data loss.

### 34. Data race on plain uint64_t Task::_perf counters updated from concurrent bridge threads

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:6366`
- **Category:** data-race (concurrency)
- **Verdict:** CONFIRMED
- **Defect:** Task::WaitWhilePaused does `_perf.pauseWaitUs += PerfElapsedUs(...)` (:6366) on a plain uint64_t (FolderWindow.FileOperationsInternal.h:267, PerfStats — only the bridge* members :243-247 are atomic). WaitWhilePaused is called concurrently by the pipeline reader jthread (:8296), the pipeline writer (:8350), parallel copy workers (:8889) and producers (:8994) of the same task. Similarly CopyDirectoryParallel does `task._perf.schedulerWaitUs += ...` and four sibling `+=` updates (:9172-9180) from multiple concurrently-running per-item processIndex threads. Concurrent unsynchronized read-modify-writes are UB per the C++ memory model.
- **Failure scenario:** Two workers of a parallel bridge copy pause simultaneously; both do the non-atomic += and one update is lost (or, formally, UB). Observable effect is corrupted FileOps.* perf telemetry for the task; no data-plane impact.

### 35. PromptDestinationCollision reads shared FileSystemOptions without the callbackMutex that writers hold

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:7155`
- **Category:** data-race (concurrency)
- **Verdict:** CONFIRMED
- **Defect:** PromptDestinationCollision copies the bridge-shared `options` struct (`FileSystemOptions issueOptions = options;` :7155) without acquiring callbackMutex, while concurrent workers' ReportProgress mutates options.bandwidthLimitBytesPerSecond under callbackMutex (:7131, :7144). With per-item parallel workers (>1 in-flight file per bridge), a conflict prompt on one worker can race a progress write on another — a torn/UB read of the live bandwidth limit handed into task.FileSystemIssue.
- **Failure scenario:** Worker A prompts for a destination collision while worker B's progress callback rewrites the bandwidth limit: the issue prompt observes a half-updated options struct. On x64 a 64-bit tear is unlikely in practice, so the realistic effect is a stale bandwidth value in the prompt path (and formal UB), not corruption.

### 36. Host transient/circuit-breaker HRESULT vocabulary does not match what Curl and S3 actually emit for network failures

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:2962`
- **Category:** contract-mismatch (contract-conformance)
- **Verdict:** CONFIRMED
- **Defect:** IsCircuitBreakerTransientError counts only ERROR_CONNECTION_ABORTED, ERROR_SEM_TIMEOUT and ERROR_TIMEOUT (:2960-2966). S3 maps NETWORK_CONNECTION errors to ERROR_UNEXP_NET_ERR (FileSystemS3.Internal.h:232-234) and Curl maps connect/DNS failures to ERROR_CONNECTION_REFUSED / ERROR_BAD_NET_NAME (FileSystemCurl.Shared.cpp:2228-2230), none of which are counted, so the per-connection circuit breaker (open after 5 failures/30 s) never trips for the most common outage modes (host unreachable, DNS dead, AWS network drop). Conversely, IsTransientCleanupHr (:7352-7358) drives the MOVE-cleanup 3x retry but omits ERROR_CONNECTION_ABORTED — Curl's mapping for CURLE_SEND_ERROR/CURLE_RECV_ERROR (Shared.cpp:2234-2235) — so the retry never fires for Curl's most common transient. No canonical HRESULT vocabulary exists in the contract (FileSystem.h documents none for stream/network failures).
- **Failure scenario:** An S3 endpoint goes offline mid-batch: every per-file operation slow-fails with ERROR_UNEXP_NET_ERR, the breaker never opens, and a 500-file task grinds through 500 full network timeouts instead of failing fast after 5. A brief FTP connection reset during MOVE source cleanup skips the retry path and immediately marks the item ERROR_PARTIAL_COPY (source kept — safe, but needless manual rework).

### 37. Unbounded-by-design memory footprint: workers x 2 buffers + S3 writer RAM parts can exceed 1 GiB

- **Location:** `RedSalamander/FolderWindow.FileOperations.State.cpp:8858`
- **Category:** performance (performance)
- **Verdict:** CONFIRMED
- **Defect:** Each parallel worker allocates its own bufferBytes buffer (:8858) and each pipelined file allocates a second bufferBytes buffer (:8257-8260): with the 16 MiB clamp ceiling and the 16-worker cap that is up to 512 MiB of pump buffers per task. On top of that, every active MultipartS3FileWriter holds up to 64 MiB (+ one incoming chunk) in a std::vector (IO.cpp:795-802) — 8 concurrent S3 uploads add ~512+ MiB — and every sub-64 MiB S3 file is fully RAM-buffered until Commit. MicrosoftDrive/Curl-IMAP writers spool to disk instead but consume equivalent temp-disk. Nothing ties bufferBytes x slots x workers x concurrent-tasks to a global budget; two simultaneous cross-FS tasks double everything. Buffers are also allocated/freed per file (secondary buffer + reader jthread per pipelined file) rather than pooled.
- **Failure scenario:** Two concurrent 8-worker copy tasks into S3 with 16 MiB hint-resolved buffers: ~2 GiB of transient RAM (pump buffers + part vectors) on a 8 GB machine causes paging that throttles the very pumps holding the memory; allocation failure mid-tree silently degrades workers (E_OUTOFMEMORY fails files) rather than shrinking buffers.

### 38. Curl promote reports failure after the point of no return: backup-delete failure surfaces as a failed promote though the destination was already replaced

- **Location:** `Plugins/FileSystemCurl/FileSystemCurl.CopyMove.cpp:1061`
- **Category:** correctness (contract-conformance)
- **Verdict:** CONFIRMED
- **Defect:** RenameWithOverwriteRollback (the path taken by the host's PromoteTempToFinalPath via FileSystemCurl::MoveItem with an overwrite grant): after PrepareOverwriteTargetForRename renamed the pre-existing destination to a .rollback sibling (:992) and RemoteRename put the temp at the final path (:1054), a failure in FinalizeOverwriteTarget's backup RemoteDeleteFile (:1061 -> :1032-1040) is returned as the promote result. The host treats any promote failure as file-failed (State.cpp:8530-8541): cleanupTemp targets the old temp name (already renamed away, delete tolerated as missing), the file is counted failed, and on MOVE the source is preserved by ShouldDeleteMoveSourceAfterBridgeCopy.
- **Failure scenario:** Overwrite-granted bridge copy/move to FTP where the final DELE of the rollback file hits a transient error: the destination file HAS been replaced with the new content, yet the task reports the file failed; the old destination content is stranded as a visible '<name>.rollback...' sibling on the server, and on MOVE the user ends up with duplicate source+destination plus a misleading failure. No data loss, but the host's assumption that promote-failure implies destination-unchanged is violated.

## Refuted (investigated, not real)

- **Curl streaming writer Commit never verifies the upload consumed all written bytes** (`Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp:1004`)
- **Pipelined pump only unlocks for files larger than the buffer; most cloud files run the fully serial one-chunk-in-flight loop, and even the pipeline allows only 1 outstanding read** (`RedSalamander/FolderWindow.FileOperations.State.cpp:7092`)

