# Corrected consumer ARM64 build qualification

Frozen RedSalamander `a20f6680dd5f82835a549f4c231e5666e4047eda`, DxUi `13788e95`,
builds locally in all three ARM64 configurations after the ViewerText Release diagnostic
correction. The process removes any compiler-host environment override, exercising the
checked-in x64-host selection and its propagation through restore and project references.

| Configuration | Build time | Warnings | Errors | Verified modules |
|---|---:|---:|---:|---:|
| Debug | 04:00 | 44 | 0 | 14 |
| Release | 08:19 | 60 | 0 | 14 |
| ASan Debug | 03:43 | 46 | 0 | 14 |

All 42 recorded module hashes were rechecked against actual bytes. Each profile uses its
matching pin/configuration and x64 compiler host. Original warnings remain in the logs.
No hosted CI ran. Cross-builds do not establish native ARM64 runtime or sanitizer detection;
that gate remains blocked by the unavailable native host in the I19 checklist.

`retained-files.json` hashes the five original receipts. Runtime Full and final portable
qualification are recorded separately; these build receipts do not replace them.
