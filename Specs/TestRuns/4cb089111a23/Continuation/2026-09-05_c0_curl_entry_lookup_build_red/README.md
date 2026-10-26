# Entry lookup fixture — initial compile failure

Run command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlEntryLookupTruth -FailFast -RunId c0-curl-entry-lookup-baseline-debug-20260905`.
Build stopped after 01:55 with C2065 (unqualified `Settings`) and C4061
(unhandled Create Directory enum). Fixed by qualifying `FileSystemCurl::Settings`
and explicitly handling the non-fixture operation. No runtime tests or valid receipt
were produced. The complete build log is retained; production was unchanged.
One whitespace-only line is trimmed for patch hygiene; diagnostic text and line
positions are unchanged. Original log SHA-256:
`041873f4a712ecd98c80e6810ea00a038d9ff8a85c5a87b76bae0f6245188226`.
