# File Operations: S3 path-style server-side copy and two harness residuals

## Status and ownership

- Status: ACTIVE. Three residuals routed from I24 (`Specs/Plans/Done/FileOperations_S3MarkerAndResiduals_2026-09-15.md`)
  on 2026-09-15, each established with evidence there and none fixed. None is a data-safety defect.
- Index owner: I25.
- Planned at: `359e4419`, 2026-09-15, on `4cb089111a23`. Recheck every anchor's drift before editing.
- This plan is self-contained and non-normative. Everything needed is in this file and in git; there
  is no dependency on the originating session, its scratchpad, or this machine.
- Authoritative contracts: `Specs/FileSystem/FileSystem_S3.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`,
  `Specs/Testing/Testing_SelfTests.md`.

## How to resume

Build one plugin with `.\build.ps1 -ProjectName <Name>`; everything with `.\build.ps1 -Rebuild`. Run a
plugin's debug self-test aggregator standalone with a 30-line loader (`LoadLibraryExW` with the
`.build\x64\Debug` directory on `PATH`, `GetProcAddress` of the export, call
`HRESULT __stdcall (unsigned int*, unsigned int*)`, read stderr); `PluginContractTests.exe` runs the
same aggregators. The governed runner is `.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter <case>
-AllowAlternateVolumeTestRoot`; a Commands `-CaseFilter` must be an exact case name. A killed governed
run leaves `.build\artifact-operations\*.contaminated.json`; clear it with `.\build.ps1 -Rebuild`.

---

## R1 — Path-style S3 endpoints never get small-object server-side copy

### Established (I24 Task 1, `359e4419`)

The S3 CRT client (`aws-c-s3` v0.13.2, `source/s3_copy_object.c` and
`source/s3_request_messages.c::aws_s3_get_source_object_size_message_new`) implements `CopyObject`
as a meta-request that first HEADs the copy source for its size. It builds that HEAD from the
`x-amz-copy-source` header in virtual-hosted style, bucket in `Host` and key alone in the path; the
library's own comment reads "this only works for virtual host endpoints and has a slew of other
issues". Against a path-style endpoint the HEAD names the wrong bucket, the meta-request fails
without sending the `PUT`, and `CopyS3ObjectWithFallback` (`Plugins/FileSystemS3/FileSystemS3.Directory.cpp`)
falls back to the bounded relay: download, then upload.

Witness, from the loopback fake's request log for a prefix Copy of three small objects:

```
[HEAD /x/src/tree/a.bin -> 404]  [GET /r0f-bucket/x/src/tree/a.bin -> 206]  [PUT /r0f-bucket/x/copy-dst/tree/a.bin -> 200]
[HEAD /x/src/tree/b.bin -> 404]  [GET /r0f-bucket/x/src/tree/b.bin -> 206]  [PUT /r0f-bucket/x/copy-dst/tree/b.bin -> 200]
```

Consequences:

- **Product.** Against any path-style S3-compatible endpoint (MinIO, Ceph RGW and similar in their
  default configuration), every small-object same-bucket Copy, and the copy half of every Move,
  moves the bytes through the client twice. Correct, but slow and network-heavy, and the relay's
  temporary-file staging costs local disk for every object.
- **Coverage.** The fake supports `CopyObject`, but its handler is never reached, so the
  "conditioned server-side copy" half of the pinned-revision route in `FileSystem_S3.md` is proven
  only by the gated live profile. `FileSystem_S3.md` records this since `359e4419`.

Objects at or above `kMultipartMinPartSizeBytes` take `UploadPartCopy` instead of `CopyObject`;
whether that path survives a path-style endpoint was not established and is the first thing to
measure.

### Options, undecided

1. Send `CopyObject` through the classic `Aws::S3::S3Client`, which issues a plain `PUT` with
   `x-amz-copy-source` and no meta-request, keeping the CRT client for everything else. Costs a second
   client type and its lifecycle in `GetS3Client`.
2. Check whether the C++ SDK's `S3CrtClient` exposes the CRT's `copy_source_uri` option (it is the
   branch that takes precedence in `aws_s3_get_source_object_size_message_new`); if it does, an
   explicit path-style source URI fixes the HEAD without a second client.
3. Configure the client virtual-hosted and make the fake honour `Host`. This fixes only the fake; real
   path-style endpoints stay slow, so it is not sufficient on its own.

Decide on measured evidence, then make the fake witness a server-side copy (its request log must
show a `PUT` with `x-amz-copy-source` and no `GET` of the source) and record the decision here and in
`FileSystem_S3.md`.

### Checklist

- [ ] Measure whether `UploadPartCopy` (large objects) survives the path-style fake.
- [ ] Choose among the options above with a measured cost; record it.
- [ ] Server-side copy witnessed on the fake; the `FileSystem_S3.md` limit sentence updated.

---

## R2 — One in-process plugin self-test failure times out every later FileOps family

### Established

On `LT-PF5VDAGE` without `SeCreateSymbolicLinkPrivilege`, `FileSystem.dll`'s object-binding self-test
failed at its first symlink fixture and returned early (`passed=105 failed=1`; I24 Task 3 now names
that check and its cause on stderr, `f5ab7fe0`). `Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker`
(`RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Fairstream.cpp`) runs those
self-tests in-process and failed as designed, and then the first case of **every later family** timed
out (I22's record lists them; 12 failures and 81 "not reached" skips). The Riptide case's own failure
path is clean: the module is released by RAII and the case returns `true`. The mechanism is therefore
elsewhere and could not be reproduced on `4cb089111a23`, whose token lacks the privilege too but has
Developer Mode on, which is what lets the FSCTL path succeed.

A second trigger, privilege-independent, was recorded on `4cb089111a23` during I24's first closing Fresh Full
(run `20260915T211116Z-48752-87739ba450f44ef18a76815c9764e343`): `Phase7_RecycleBinBatchDeleteMultiBatch`
hung in the shell's recycle operation and timed out at 240 s (it had passed in ~10 s in both earlier gates
on the machine and passed 9/9 immediately afterwards), the rest of its family was aborted, and the next
family's `Setup` then failed with "Failed to create temp root directory for self-test" after 65 s: 2
failures and 36 "not reached" skips from one shell hang. The mechanism is therefore not tied to the
symlink privilege; it is that a family which dies mid-case leaves its temp root held, and the next
family's Setup has no way to reclaim it.

### To do

- [ ] Make the harness reclaim or side-step a held temp root when a family aborts mid-case (a fresh
      per-family root, or a Setup that retries on a new name), so one hung case costs one case.
- [ ] Reproduce the symlink-privilege trigger on a machine, or a token, without the privilege and
      without Developer Mode; do not turn Developer Mode off on a shared machine to get there.
- [ ] Find what the early return leaves behind (fixture files or handles under the temp root, a
      scheduler worker, a global) that makes the next family's Setup block, and make the harness keep
      running later families after an in-process plugin self-test failure.

---

## R3 — A crashed self-test process does not exit

### Established (I24 Task 2)

With the palette use-after-free reintroduced deliberately, the Commands self-test process crashed at
the expected access violation, the crash handler wrote
`RedSalamander-20260915-220715-p66660.txt` and its `.dmp` within four seconds, and the process then
stayed alive, not responding, with 580 threads, until it was terminated by hand. The governed runner
only reaped it at its case timeout. This is what let one crash take a whole Commands suite down on the
laptop rather than one case. The crash report itself was complete and correct.

### To do

- [ ] Establish where the crashed process waits after the report is written
      (`CrashHandler.cpp` / `CrashQuarantine.cpp`; a wait for a debugger, WER, or a quiet-point join).
- [ ] Make a self-test process exit promptly after its report, or have the runner treat a written
      crash report as terminal for the case without waiting for the case timeout.

## Completion checklist

- [ ] R1 decided and witnessed, R2 reproduced and fixed, R3 fixed; each with its own commit and evidence.
- [ ] Durable outcomes merged into the authoritative specs named above.
- [ ] `git diff --check` and the spec inventory pass; the completed plan moves to `Specs/Plans/Done/`
      and I25 leaves the active index.
