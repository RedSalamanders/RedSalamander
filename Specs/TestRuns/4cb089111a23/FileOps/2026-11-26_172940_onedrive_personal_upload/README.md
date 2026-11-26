# I66 OneDrive Personal upload-session origin

The approved live x64 Release case reached an Access denied decision on its first Copy and recorded its native timeout in `20261126T154910Z-20984-3353ab60891947a7a4717549c0f2eb15`. An instrumented x64 Release run on `6011e7ff243c87feba99c17804b4adae0d4bf83c` then reported the actionable `0x80070005` decision in 9.296 seconds in `20261126T161407Z-51312-27e22f568cae4a9e98ca2ebc3c69a9f5`. Its build receipt was `0088f8cd8bf42e75d2e6379fae27edc894e8bc406600ca079cab32c0d412f40b`.

The archived performance rows show that the authenticated Graph session request returned HTTP 200 with a secure URL on the Microsoft consumer content origin. URL validation rejected it before the upload request. The saved launcher initially tried to reduce the unsigned HRESULT as a signed integer. The raw archived values establish the result.

| Original row | Metric | Recorded result |
| --- | --- | --- |
| 1404 | `FileOps.MicrosoftDrive.RejectedUploadOrigin` | Exact `my.microsoftpersonalcontent.com`, HTTPS/default port with valid URL parts, `0x80070005` |
| 1405 | `FileOps.MicrosoftDrive.UploadSessionFailure` | Session URL validation, HTTP 200, `0x80070005` |

The original `perf_metrics.jsonl` has 2,173 rows and 806,927 bytes, SHA-256 `318019218b3fac1da7b6ab1c5560d81b2dae69593f73f9e10bc69b2c6a8b8080`. The raw artifact is retained under the marked sandbox. This record contains only the stage and redacted transport facts.

The fix accepts the exact observed host for OneDrive Personal. The streaming and buffered session paths keep the Graph and preauthenticated credential audiences separate. Existing TLS, port, userinfo, fragment and redirect rules apply. The validation fixtures exercise the previously allowed hosts, the exact personal host, its wrong-mode and lookalike refusals, the wrong-port refusal, streaming completion and query-secret redaction.

All corrected runs used test-enabled x64 Release head `fb132832da2d31dcff4c332c1cb331d2f3e3a7cc`, built with zero warnings and errors. The verified receipt was `e193a8ed5128d86522d80c6bba0e014b7bc03e072701be855eb0f7a205ecc4e8`.

| Exercise | Run | Result |
| --- | --- | --- |
| Live OneDrive Personal Copy, same-drive Move and Recycle | `20261126T162543Z-72592-86e623ea7ba0477f99504c3487776173` | 3 passed, zero failed or skipped. The case took 13,063 ms. |
| Host Graph loopback | `20261126T162634Z-68340-fd9cf18109644c1e9e86c5710e7b0f10` | 3 passed, zero failed or skipped. |
| Standalone plugin contracts | `20261126T162840Z-55964-11f73e45c8ce498482662ad8b6b5b9a4` | Exit 0, including all 270 Microsoft Drive assertions. |

The live case verified the upload and download bytes, source absence after the Move and Recycle of the moved file. The fresh test child was absent from sync after cleanup. Its exact case result JSON has SHA-256 `b408b61f5abf33edaf4c04ff43012162dfc102f6fffc74b0a89290c40e59d8ba`. The run used the previously approved isolated profile, retired its settings and verified the installed settings digest. The earlier retained Upload directory still needs exact recovery evidence. Business and SharePoint live routes and broader native qualification remain open.
