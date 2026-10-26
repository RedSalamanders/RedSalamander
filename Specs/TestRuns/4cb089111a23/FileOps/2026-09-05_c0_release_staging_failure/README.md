# C0 Release invocation — no accepted build receipt

Run attempt `c0-receipt-prefix-release-20260905` used a focused Release invocation
without explicitly setting `RSBuildEnableTests=true`. It cannot qualify the test
matrix. MSBuild compiled with zero warnings/errors, but `build.ps1` then failed:
the required profile-root `.build/x64/Release/sqlite3.dll` was absent when collecting
the build receipt. No receipt or runtime test result was produced.

`build.txt` is the unchanged MSBuild log. It records the root SQLite copy earlier
in the graph; subsequent read-only inspection found the vcpkg source and
`Plugins/sqlite3.dll` present but the profile-root DLL missing. The cause of its
disappearance is not yet proved. Runtime dependencies must be repaired/validated
through their canonical manifest, not by manually copying a DLL around attestation.

The corrected, explicitly test-enabled run is
`c0-receipt-prefix-enabled-release-20260905`. Its outcome is recorded separately.
A successful test-enabled run does not by itself close this non-test Release
staging observation. Keep it in the final build qualification review.
