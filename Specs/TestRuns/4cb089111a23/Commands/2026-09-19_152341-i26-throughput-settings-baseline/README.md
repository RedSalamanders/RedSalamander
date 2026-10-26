# Throughput disclosure setting baseline

This is the retained failing witness for the `popupThroughputExpanded` persistence contract.
Production source was still at `89dfe40c7`; the only source change was the focused Commands selftest
that required an explicit `false` value to survive canonical save preparation as the sole non-default
file-operations setting.

The x64 Debug build completed with zero warnings and zero errors. Its receipt is
`5e551b9032173de5ba08ce4a9001937aea15a97607bf2ebf94581a2c8612a25b`.
The governed command was:

```powershell
./Tools/Run-AllTests.ps1 -Suite Commands -Configuration Debug -Platform x64 `
    -CaseFilter settings_file_operations_roundtrip -TestRoot S:/RedSalamander.Perf
```

Run `20260919T132341Z-62368-91150214a296433795dc356e987278ce` failed the expected single case
with: `Explicit popupThroughputExpanded=false must be retained as the sole non-default file-operations setting.`
The runner's sandbox audit reported zero issues.

The subsequent governed candidate run replaced the transient run directory under the fixed test root.
This archive therefore retains the complete build and runner logs plus the immutable build receipt;
it does not reconstruct or claim missing generated result files.
