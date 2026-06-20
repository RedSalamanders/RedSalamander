# RedSalamanderMonitor

`RedSalamanderMonitor.exe` is a high-throughput, append-only log viewer designed for real-time streaming (ETW) as well as reviewing captured logs.

![Monitor](res/monitor.png)

## Getting access (ETW)

On some machines, starting an ETW listener requires extra privileges.

One-time setup (adds you to **Performance Log Users**):

```powershell
.\init-etw-trace.ps1
```

Then sign out/in (or reboot) and start the monitor:

- `.build\x64\Release\RedSalamanderMonitor.exe` (Release)
- `.build\x64\Debug\RedSalamanderMonitor.exe` (Debug)

To undo:

```powershell
.\init-etw-trace.ps1 -Remove
```

## Key features

- **Auto-scroll (tail mode)** for live logs (toggle: **Option → Auto Scroll**, shortcut `End`)
- **Filters** by message type (Text/Error/Warning/Info/Perf/Debug) with presets
- **Find** (`Ctrl+F`, then `F3` / `Shift+F3`)
- **Always on top** option
- Theming consistent with the main app (View → Theme)

## A typical session: capture → filter → compare

The Monitor is most useful as a live window onto whatever RedSalamander (or the Monitor itself) is doing right now. A common workflow:

1. **Capture.** Start `RedSalamanderMonitor.exe` and leave it running. It opens an ETW session automatically and begins appending events as they arrive — there is no "record" button to press. With **Auto Scroll** on (the default), the view stays pinned to the newest line, so you watch activity stream by in real time. Now reproduce whatever you want to observe in RedSalamander.
2. **Filter.** A live feed is noisy. Narrow it down with the type toggles under **Option → Active Filter** (see below) — for example, switch to **Preset: Errors Only** to spot failures, or turn off **Info** and **Debug** to cut routine chatter. Filtering only changes what is *displayed*; nothing is dropped from the underlying capture, so you can widen the filter again at any time and the hidden lines reappear.
3. **Compare.** To study what happened rather than what is happening, scroll up. Scrolling away from the bottom pauses tail mode, freezing the view so lines stop moving under you while you read. Use **Find** (`Ctrl+F`, then `F3` / `Shift+F3`) to jump between occurrences of a string, and the line-number gutter (**View → Line Numbers**) to keep your place. Press `End` (or re-enable **Option → Auto Scroll**) to snap back to the live tail.

The status bar at the bottom tracks the session: the active filter description, the **visible** line count, the **total** line count (visible + filtered), and how many ETW events have been received.

## Filtering by message type

Every log line carries a type, and the Monitor shows or hides lines by type using six independent toggles under **Option → Active Filter**:

| Toggle | Hides/shows |
|--------|-------------|
| **Show Text** | Plain text lines with no severity |
| **Show Errors** | Error-level diagnostics |
| **Show Warnings** | Warning-level diagnostics |
| **Show Info** | Informational flow detail |
| **Show Perf** | Performance scopes and counters |
| **Show Debug** | Verbose debug traces |

Each toggle has a checkmark; clicking it flips that one type on or off and leaves the others alone. All six are checked by default, so a fresh capture shows everything.

Below the toggles are four one-click presets that set the whole combination at once:

- **Preset: Errors Only**
- **Preset: Errors+Warnings**
- **Preset: Errors+Perf+Debug**
- **Preset: All Types** (the default — re-checks every toggle)

Choosing a preset updates the individual checkmarks to match. Toggling any single type afterward simply moves you to a custom combination. Your last filter choice is remembered between runs.

The screenshot at the top of this page shows this menu open: **Option → Active Filter** is expanded, with the six **Show …** type toggles at the top (each with its checkmark) and the four **Preset: …** entries below the separator. The surrounding **Option** menu also holds **Always On Top**, **Auto Scroll** (`End`), and **Show Process ID/Thread ID**, which is why those choices live one level up from the filter submenu.

For the underlying log types, ETW event schema, and how each type maps to a filter bit, see [dev/Diagnostics.md](dev/Diagnostics.md).

## Opening and saving logs

Beyond the live feed, the Monitor can review captured text:

- **File → Open…** (`Ctrl+O`) loads a `.txt` or `.log` file into the view. Encoding is detected automatically from the byte-order mark (UTF-8 or UTF-16LE), falling back to UTF-8. Opening a file replaces whatever is currently displayed.
- **File → Save As…** writes the current view out to a text file (with an overwrite prompt if the target exists). This is how you hand a captured session to someone else or keep it for later comparison.
- **File → New** (`Ctrl+N`) clears the view so you can start a clean capture.

To copy a portion rather than the whole log, select lines in the view and use **Edit → Copy** (`Ctrl+C`); `Ctrl+A` selects everything first.

## Launching and command-line options

Launch the Monitor from the build output:

- `.build\x64\Release\RedSalamanderMonitor.exe` (Release)
- `.build\x64\Debug\RedSalamanderMonitor.exe` (Debug)

It accepts a few switches (run with `-h`, `--help`, or `/?` to print the full list):

| Option | Effect |
|--------|--------|
| `-h`, `--help`, `/?` | Print usage and exit. |
| `--etw` | Show the Monitor's *own* Info/Perf/Debug events and startup status. By default the Monitor stays quiet about itself and hides events emitted by its own process; this opt-in surfaces them. |
| `--perf` | Write the Monitor's own performance metrics to the default JSONL path. |
| `--perf=PATH` | Write those metrics to a custom path. |

`--etw` controls whether the Monitor *displays* its self-originated diagnostics; it is separate from the privilege needed to *start listening* to ETW (see "Getting access" above). For the full producer/consumer pipeline, the JSONL perf sink, and the diagnostic self-test flags, see [dev/Diagnostics.md](dev/Diagnostics.md).
