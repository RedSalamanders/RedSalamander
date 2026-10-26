# C0 IMAP UID SEARCH candidate — Debug

- Run ID: c0-imap-uids-candidate-debug-20260906
- Build: test-enabled x64 Debug; receipt 80d5eb554d77a2474d8476bc638688393786a0343a8f8428e3f1c074ee413e53
- Focused File Operations case: C0_CurlImapListingTruth — 3/3 runner cases passed (Setup, focused case, cleanup); the native case includes 357 assertions.
- Standalone FileSystemCurlTests.exe: passed (12 helper functions).
- Source/resource policy: 197/197 assertions passed in run c0-imap-uids-policy-20260906.
- MSBuild log msbuild-20260906_051407_863-pid140988-871c065a.log: zero warnings and zero errors.

The candidate validates complete case-insensitive SEARCH framing, strict nonzero uint32 UID tokens, duplicate/multi-row aggregation, borrowed opaque framing, explicit ESEARCH unsupported behavior, cancellation/control checkpoints, and no partial publication or FETCH on failure. The focused and broad Release C0 qualifications were subsequently completed and are archived alongside this Debug candidate. No guard bypass was used. The baseline is archived separately.
