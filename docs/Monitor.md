# Monitor diagnostics

RedMonitor is an independent optional diagnostics viewer. Install its standalone Release ZIP with
Install-RedMonitor.ps1. The default location is LocalAppData/Programs/RedMonitor; LocalMachine uses
ProgramFiles/RedMonitor. RedSalamander builds and packages share only the pinned client header/license.

Start the helper yourself or run RedSalamander.exe --redmonitor. A bare Debug startup may also queue an
optional asynchronous launch if a production installed helper is present. Release, help and test startup
never auto-launch it. Missing helpers do not block RedSalamander; an explicit request returns failure.

Messages carry the RedSalamander label, severity, Unicode text and original process/thread/time metadata.
Existing debugger, file and crash diagnostics remain independent. ETW collection may require Performance
Log Users membership or elevation. Preferences continues to edit the Release monitor 7.0 settings profile
in every configuration, including Debug; standalone Debug viewer profiles are separate development tools.

Developers use Tools/Update-RedMonitorClient.ps1 -Commit FULL_SHA to update only Client.h, its MIT notice
and the immutable lock. Ordinary builds run the local verifier and never clone or compile the helper.