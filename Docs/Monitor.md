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

- `x64\Release\RedSalamanderMonitor.exe` (Release)
- `x64\Debug\RedSalamanderMonitor.exe` (Debug)

To undo:

```powershell
.\init-etw-trace.ps1 -Remove
```

## Key features

- **Auto-scroll (tail mode)** for live logs (toggle: **Option → Auto Scroll**, shortcut `End`)
- **Filters** by message type (Text/Error/Warning/Info/Debug) with presets
- **Find** (`Ctrl+F`, then `F3` / `Shift+F3`)
- **Always on top** option
- Theming consistent with the main app (View → Theme)

