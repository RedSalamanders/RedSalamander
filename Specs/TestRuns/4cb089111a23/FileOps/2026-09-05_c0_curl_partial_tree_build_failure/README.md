# C0 Curl partial-tree candidate — build failure

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-partial-tree-candidate-debug-20260905`.

Source: `58f22d0e` plus the C0 working changes and first partial-tree repair/test
expansion. The build stopped after 1:44 with four compiler errors in two new
fixture blocks: `FakeFtpEndpoint` requires `UploadRetention::Complete`, and the
existing `CreateConfiguredFtpSelfTestFileSystem` takes the output COM owner, not
a port argument. The corrective change uses those existing APIs.

The full build log also retains one MSB3061 cleanup warning for the Debug
`Plugins/yyjson.dll` path. An exact-path module-holder inspection afterward found
no process with that DLL loaded; this does not identify the transient access
denial's cause. No process was killed and no user data was removed. This build is
not warning-free evidence. No receipt was published and no runtime test ran.

`build.txt` is the intact 26,035-byte MSBuild log. The pre-fix runtime regression
is separately preserved in `2026-09-05_c0_curl_partial_tree_red_debug`.
