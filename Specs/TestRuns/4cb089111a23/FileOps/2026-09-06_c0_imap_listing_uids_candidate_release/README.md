# C0 IMAP UID SEARCH candidate — Release focused

- Run ID: c0-imap-uids-candidate-release-focused-20260906
- Build: test-enabled x64 Release; receipt ecd372340dc4af1b6efca7cf7faa982498643f482e0adf131e69ad43b403f919
- Focused File Operations result: 3/3 runner cases passed (Setup, C0_CurlImapListingTruth, cleanup); the native case includes 357 assertions.
- FileSystemCurlTests.exe: passed (12 helper functions).
- Source/resource policy: 197/197 assertions passed in c0-imap-uids-policy-20260906.
- MSBuild log msbuild-20260906_101106_919-pid123376-95a23ec4.log: zero warnings and zero errors.

This archive is the explicit Release candidate proof for strict case-insensitive UID SEARCH parsing, malformed/missing response failure, duplicate/multi-row aggregation, opaque framing, ESEARCH unsupported behavior, cancellation checkpoints, and no partial/FETCH publication on failure. No live connection or certificate setting changed.
