# Final RedSalamander ARM64 cross-builds

Source `ad7f7735591c3cd18bd43bddce0a7f06cdd3318b` builds the complete solution
against DxUi `13788e95f1c9a16210a87988d742d9dfc6f66eef` with the corrected
consumer compiler-host propagation. No PreferredToolArchitecture environment
override was supplied; the checked-in first-party policy selected x64.

| Configuration | Build time | Warnings | Errors | Verified DxUi modules |
| --- | --- | --- | --- | --- |
| Debug | 3m51s | 44 | 0 | 14 |
| Release | 4m58s | 60 | 0 | 14 |
| ASan Debug | 3m40s | 46 | 0 | 14 |

All 42 module records match the actual module bytes, exact pin, configuration,
architecture and x64 compiler host. Archive hashes are retained per module.
Warnings remain in the original build logs. The earlier rejected host mismatch
and incorrect probe-path attempt remain in the adjacent qualification record.

These are sequential cross-builds on Windows x64. They do not establish native
ARM64 execution or native ARM64 sanitizer detection. The unchanged sanitizer
compile guards remain enabled. No hosted CI ran.

retained-files.json records the exact copied bytes of the logs and identity record.
