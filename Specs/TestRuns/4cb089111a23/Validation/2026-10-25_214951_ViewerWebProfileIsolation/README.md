# ViewerWeb profile isolation baseline

Source cfe43d8ddc92973fd47955876cb80ea0b061509e passed all six builds and the
37 focused checks (including setup/cleanup and repeats), with no focused skips.
ARM64 retained its 37 existing C4324 alignment warnings; none was hidden.

Fresh ASan Full 20261025T214951Z-85840-7d7ea33292d048f0b1762d02c6552550
failed ViewerPE noninteractive: the real private HTML origin was not observed,
and both PDF navigation completion/success/source predicates failed. Ten entries
were promoted; six remained pending and Compare was interrupted. The original
running-state file remains unchanged and is not complete Full qualification.
The already-failed attempt was interrupted before further browser entries reused
an independently running application's profile. The failed entry's full native
output, evidence/terminal records, plan and build receipt are preserved.

The ViewerWeb environment used the ordinary LocalAppData RedSalamander browser
profile, also in use by the separately launched Release executable under Z:/src.
That instance remained open. No production browser process or data was removed.
Three unchanged-binary security-case probes used separate authorized E: sandbox
profiles. All passed in about three seconds each, retaining every original HTML/PDF,
scripting, byte-bound and twelve-second navigation assertion. Exact executable and
plugin digests and complete probe outputs are retained. These controls support
fixture isolation; they do not promote the failed/incomplete Full or establish
every possible runtime failure cause.

The repair assigns each ViewerPE process an owned profile before browser startup.
The strengthened security fixture reads the real WebView2 environment profile back
and verifies exact authorized ownership. Candidate builds and all required Fresh
Full profiles remain pending.

The SDK documents the override and profile-sharing failure modes in
[CreateCoreWebView2EnvironmentWithOptions](https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/webview2-idl?view=webview2-1.0.3912.50#createcorewebview2environmentwithoptions).
