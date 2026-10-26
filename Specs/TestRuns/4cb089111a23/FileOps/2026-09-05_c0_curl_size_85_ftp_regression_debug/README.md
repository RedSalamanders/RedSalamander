# C0 Curl directory-size repair — ordinary FTP regression

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -FailFast -RunId c0-curl-size-85-ftp-regression-debug-20260905`.

Build: 131,336 ms, zero warnings/errors. Receipt
`879c3009a9f77c80447161ce49c1f8147c80fb48c110ab399d8a377f522ef454`;
source `3663194d725a23644d1ff88b3eeb10eb4d3955600ed513571a41e0a791c190df`.
Host: 3/3 passes, zero failures/skips; 2,563 ms suite duration.

This is the ordinary read/write/create/delete FTP workflow plus setup/cleanup,
not three independent provider qualification scenarios. It passes with the same
final source identity as Debug C0 and test-enabled Release C0. The run protects
ordinary behavior around the size callback and return-to-pool option reset.

All five raw files are unchanged SHA-256-verified copies from the named owned run
beneath `C:/RedSalamander.Perf`, including the complete 710,183-byte perf stream
(1,956 rows; SHA-256 `b2075b36c74c84d4688f65cf9508148983d30668ab40c60ce75f644a5c4d6f90`).

The three final runs bind the same source snapshot. Subsequent edits only finalize
spec/evidence reporting. Source/localization checks pass 197/197. The intermediate
shared z.dll Clean warning remains a C6 issue despite these warning-free builds;
its exact build log is archived with that intermediate run. Retained disk-audit
findings were not deleted. This closes the size-repair substep only: C0's other
routes, C1–C6, inherited I17 and exact Fresh Full still prevent moving the plan to Done.
