# Crash handling

This contract owns local crash reports, restart recovery and isolated crash probes for RedSalamander.
The startup contract sets their order relative to the rest of the application.
[Startup and bootstrap](Core_StartupBootstrap.md)

## Installation and fatal paths

An ordinary start pins its fixed local crash directory and initializes the completion event and writer
buffers before startup services, COM, windows or workers. DbgHelp is linked at process load. Startup
returns a failure when the crash state cannot be initialized. Command help and harness startup leave
the product crash store alone.

The top-level SEH filter, unhandled exception filter, `std::terminate`, CRT pure virtual call, invalid
parameter and `SIGABRT` join one fatal writer. CRT callbacks capture an exception context and
assign a distinct fatal code. The elected thread creates and flushes one local minidump, then
atomically publishes its marker from an owned temporary file. A second faulting thread waits for
that attempt to finish. Recursion on the writer thread ends the process immediately. The process
terminates with the original exception code or the assigned CRT failure code.

The application uses fixed buffers and local file operations in the fault writer. It records the stack
and exception context in the minidump. Symbolization runs in a healthy analysis process. The dump
and marker stay on the local machine. Existing text stack reports remain available for forensics.

## Local artifacts and restart

The local store is `%LOCALAPPDATA%\RedSalamander\Crashes`. The `last_crash.txt` marker
contains the completed dump path in UTF-16. New dumps use
`RedSalamander-report-v1-<20-digit UTC file time>-p<pid>.dmp`. An associated text sidecar
uses the same stem and a `.txt` extension. Existing dump names and sidecars remain readable.

On the next healthy interactive start, claim the marker through its owned handle before presenting
a single localized dialog. Show the local dump path when it is available. The dialog offers the
crash folder and, when the previous session recorded active file-system plugins, offers to disable
the implicated plugins before plugin initialization. The folder and plugin decisions are independent.
A headless or self-test start leaves the pending notice for the next interactive start.

Healthy interactive startup keeps the newest ten owned dumps with their sidecars. It preserves an
older pending dump through its restart notice and keeps legacy files, unrelated files and reparse
points. Cleanup is confined to the fixed local store. A later healthy start retries incomplete
cleanup.

## Qualification

- `--crash-probe` and its `--crash-test` alias use the governed test run's artifact directory.
  `--crash-probe-worker` exercises two overlapping faults. Both routes precede normal startup.
- The ordinary probe raises a non-continuable access violation. ASan Debug uses a distinct
  application-defined exception code. Verify real sanitizer detection separately.
- `CrashHandlingTests` covers child-process SEH, worker overlap and abort. Verify each exit code,
  one nonempty dump and the marker's exact path. Its healthy fixtures cover zero, ten and eleven
  reports, a pending older dump, legacy evidence and a headless start followed by two claims.
- On the grouped qualification source, verify the single interactive dialog, its crash-folder
  action and the implicated-plugin decision through the governed desktop route. Preserve the
  native x64 and ARM64 evidence with the build receipts.
