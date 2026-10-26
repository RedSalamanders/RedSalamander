# Hosted ARM64 browser-profile failure

At b5ea0d679, the native ARM64 Debug build and 1008 tooling checks passed (two
existing optional tooling skips). Capture and 200k-file terminal-drain checks
passed. ViewerPE retained three failing readbacks because the actual WebView2
profile was normal AppData instead of the harness's owned sandbox directory.
The complete failed job log remains unmodified and hashed. It is not promoted
by local x64 passes.

The next candidate explicitly passes a validated marked sandbox profile in test
builds. Invalid test overrides fail closed and actual browser profile readbacks
remain mandatory. Hosted qualification remains pending for that repair. Full
runtime reruns remain omitted under the user's accepted scope.
