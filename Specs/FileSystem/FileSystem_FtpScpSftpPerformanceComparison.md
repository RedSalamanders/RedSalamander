# FTP / SCP / SFTP Backend Comparison — Retired Research

Status: RETIRED historical note
Retired: 2026-08-04
Last research baseline: `6aecfde8e`

## Why this document is retired

This file previously mixed a snapshot of the shipped libcurl/libssh2 implementation with speculative recommendations to tune it, replace it with libssh, or add a hybrid backend. Those recommendations were not accepted product requirements and contradicted the current provider contract when read as normative work.

The current authority for FTP, SFTP, SCP, and IMAP behavior is `FileSystem_FtpSftpScp.md`. The implementation is `Plugins/FileSystemCurl/`; dependency identity is owned by the repository vcpkg manifests and lock data. Performance requirements and evidence rules are owned by `Testing/Testing_PerformanceValidation.md`.

## Current decision

RedSalamander currently uses libcurl, with libssh2 underneath curl for SSH protocols. The transfer backend is not a UI or plugin-ABI promise. No alternative backend project is approved or queued by this retired comparison.

A future backend change requires a new scoped WIP plan with representative protocol/workload scenarios, instrumentation, deterministic regression coverage, before/after archived evidence, packaging impact, and an explicit dependency/security review. It must update `FileSystem_FtpSftpScp.md` when the shipped contract actually changes.

## Historical recovery

The pre-retirement research can be inspected without restoring it into the normative tree:

```powershell
git show 6aecfde8e:Specs/FileSystem/FileSystem_FtpScpSftpPerformanceComparison.md
```
