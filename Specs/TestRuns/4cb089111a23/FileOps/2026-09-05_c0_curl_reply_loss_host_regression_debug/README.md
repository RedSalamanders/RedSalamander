# FTP host workflow after the lost-reply repair

- Source: `58f22d0e` plus the C0 working changes; same source as the
  [candidate Debug C0 run](../2026-09-05_c0_curl_reply_loss_debug/README.md).
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter R0fCurl_FakeFtpReadWriteCreateDelete -FailFast -RunId c0-curl-reply-loss-host-regression-debug-20260905`.
- Accepted build receipt: `a57f9c20a938940c6cffecd424d642d58edc154d0d85c11fe75cb00395401231`;
  zero warnings/errors; build 2:42.
- Harness: 3/3 passed, zero skipped. This exercises host Copy into and out of FTP,
  byte verification, replacement, Create Directory, Rename, and recursive Delete.
  Ordinary mutations remain usable after removing quote replay and incidental NLST.

The evidence files are copied intact from
`C:\RedSalamander.Perf\runs\c0-curl-reply-loss-host-regression-debug-20260905\artifacts\selftest\last_run`.
The runner reported 35 retained-run disk-audit issues. No old runtime data was
removed. This loopback regression is not real-server, deep/wide native-walker,
production non-test Release, or final Fresh Full qualification.
