# C0 IMAP UID SEARCH candidate — Release broad lane

- Run ID: c0-imap-uids-candidate-release-20260906-r3
- Build: test-enabled x64 Release; receipt ecd372340dc4af1b6efca7cf7faa982498643f482e0adf131e69ad43b403f919
- Broad File Operations result: 18/18 C0_ cases passed, 0 failed, 0 skipped.
- C0_CurlImapListingTruth passed inside the broad lane (357 assertions).
- FileSystemCurlTests.exe: passed (12 helper functions).
- Source/resource policy: 197/197 assertions passed in c0-imap-uids-policy-20260906.
- MSBuild log msbuild-20260906_101106_919-pid123376-95a23ec4.log: zero warnings and zero errors.

This archive preserves the broad C0 regression result after the UID-search repair, including Google Drive, Curl mutation/traversal, IMAP, and lookup cases. No live connection or certificate setting changed.
