# Operation Chordforge Release performance evidence

## Identity and result

- Scenario: terminal shortcut routing, global command palette, Terminal Find,
  authenticated history Suggestions, hosted zoom/reflow, and the singleton
  floating Terminal lifecycle.
- Machine hash: `4cb089111a23` (`SINON`).
- Environment: Windows 11 Pro x64, build `26200`; local console; x64
  test-enabled Release build.
- Embedded repository commit: `f9cee9d47d7652951d69ec106751ff2301fe1184`.
  The binary also contains the uncommitted Operation Chordforge worktree.
- Runner ID: `20260811T192050Z-97512-d3ee19ffa6494d449080a6f4c13556d8`.
- Result: `terminal_shortcut_perf_command_experience_release` passed `1/1`
  in 48,130 ms; runner wall time 49,013 ms; disk audit clean.

Binary identities:

| Binary | SHA-256 |
|---|---|
| `RedSalamander.exe` | `6693c062637a629ea11cd926ecebde83a97b3658b5afba9497c5050a680a9d9e` |
| `Plugins/Terminal.dll` | `c59483c17e1e1d58d348f96529e6140d3a0b6252cd70bc2238a100b2ab67f6d5` |
| `Common.dll` | `0c4041a1d4470f7b92c4b687dcc35a3baea5d437c6816b7ae540f1bc88ce4f00` |

## Reproduction

The Terminal model/router perf export and the host case are compiled only in a
test-enabled Release build:

```powershell
try {
    $env:RSBuildEnableTests = 'true'
    .\build.ps1 -ProjectName RedSalamander -Configuration Release
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}

try {
    $env:REDSALAMANDER_TERMINAL_CHORDFORGE_PERF = '1'
    .\Tools\Run-AllTests.ps1 -Suite Commands -Configuration Release -SkipBuild `
      -CaseFilter terminal_shortcut_perf_command_experience_release -TimeoutMultiplier 8
} finally {
    Remove-Item Env:REDSALAMANDER_TERMINAL_CHORDFORGE_PERF -ErrorAction SilentlyContinue
}
```

## Claim-bearing measurements

All values are microseconds and use the repository analyzer's nearest-rank
percentile convention.

| Metric/detail | Count | P50 | P95 | Max | Budget/result |
|---|---:|---:|---:|---:|---|
| `app.command_palette.open_to_visible_us` | 101 | 21,701 | 26,984 | 30,176 | <= 100,000; pass |
| `app.command_palette.filter_batch_us` | 1,101 | 153 | 357 | 1,160 | <= 16,700; pass |
| `shortcut.catalog.rebuild_us` | 101 | 1,420 | 2,152 | 3,338 | diagnostic; no paint rebuild |
| `terminal.find.query_to_visible_us` | 1,000 | 3,303 | 4,228 | 5,687 | <= 50,000; pass |
| `terminal.suggestions.load_batch_us` | 1 | 2,279 | 2,279 | 2,279 | 10,000 entries / 4 MiB; pass |
| `terminal.suggestions.open_to_visible_us` | 100 | 20,935 | 22,728 | 25,844 | <= 100,000; pass |
| `terminal.suggestions.filter_to_visible_us` | 1,000 | 7,491 | 9,687 | 15,643 | <= 16,700; pass |
| `terminal.floating.window_open_to_visible_us` | 3 | 55,611 | 120,554 | 120,554 | <= 250,000; pass |
| `terminal.floating.tab_open_to_visible_us` | 126 | 98,211 | 113,823 | 191,058 | p95 <= 150,000; pass |
| `terminal.session.root_exit_to_host_close_us` | 25 | 24,600 | 30,862 | 32,154 | <= 100,000; pass |
| `terminal.session.exit_to_host_close_us` | 25 | 14,766 | 16,948 | 19,262 | diagnostic callback slice |
| `terminal.render.frame_us` / `zoom_baseline` | 500 | 8,299 | 10,080 | 12,887 | paired baseline |
| `terminal.render.frame_us` / `zoom_candidate` | 500 | 852 | 8,615 | 10,724 | -14.5% p95; pass |

The in-run shortcut bypass comparison measured baseline/candidate p95 at
`1 us`/`1 us` across 50,000 messages, within the `15% + 2 us` and absolute
`25 us` gates. The 1,000 held-scroll messages completed in `223 us`; the 100
hosted increase/decrease/reset cycles completed in `2,420 us`. The Find
snapshot publication remained below the 4 ms UI-thread budget. The final
floating root emitted exactly one settled placement-write episode and all
three root lifetimes emitted zero retained roots, tabs, Terminal instances,
callbacks, and posted payloads.

Route and hosted-render comparisons are same-machine, same-suite paired
baseline/candidate measurements inside this run. Other new-feature metrics use
the absolute budgets in the completed plan; there is no misleading pre-feature
product baseline.

## Sample quality and analyzer commands

The 1,000+ sample filter/search/render claims satisfy the repository default
p95 minimum. The authoritative scenario deliberately defines 100 warm opens,
126 tab opens, 25 real shell exits, and three top-level window creations; their
analyzer commands therefore declare those reviewed per-metric minima.

```powershell
$run = 'Specs/TestRuns/4cb089111a23/Commands/20260811_192050_operation_chordforge_terminal_command_experience_release'
.\Tools\Show-PerfRuns.ps1 -Run $run -Metric terminal.render.frame_us -FailOnQuality
.\Tools\Show-PerfRuns.ps1 -Run $run -Metric app.command_palette.filter_batch_us -FailOnQuality
.\Tools\Show-PerfRuns.ps1 -Run $run -Metric terminal.find.query_to_visible_us -FailOnQuality
.\Tools\Show-PerfRuns.ps1 -Run $run -Metric terminal.suggestions.filter_to_visible_us -FailOnQuality
.\Tools\Show-PerfRuns.ps1 -Run $run -Metric app.command_palette.open_to_visible_us -MinimumSamplesForP95 100 -FailOnQuality
.\Tools\Show-PerfRuns.ps1 -Run $run -Metric terminal.floating.tab_open_to_visible_us -MinimumSamplesForP95 100 -FailOnQuality
.\Tools\Show-PerfRuns.ps1 -Run $run -Metric terminal.session.root_exit_to_host_close_us -MinimumSamplesForP95 25 -FailOnQuality
.\Tools\Show-PerfRuns.ps1 -Run $run -Metric terminal.floating.window_open_to_visible_us -MinimumSamplesForP95 3 -FailOnQuality
```

## Compact artifact manifest

The raw JSONL exceeded the checked-in archive limit. It remains locally at the
runner path named above and was compacted deterministically:

| Artifact | Bytes | Rows | SHA-256 |
|---|---:|---:|---|
| Raw `perf_metrics.jsonl` | 10,474,542 | 28,069 | `393b340fc5f4d238d87a08a0e1d614a5ef158b9728db3df396b291b4ea03132d` |
| Checked-in compact JSONL | 1,853,408 | 4,826 | `fe19fd3b0c954922bc30b338b52962c391d9c830899d70295c316dbfaa53a9aa` |

Compaction preserves, in original order, every claim-bearing route summary,
palette/catalogue, Find, Suggestions, hosted zoom baseline/candidate, floating
lifecycle, root-exit, placement/coalescing, and zero-retention row. It removes
unrelated application startup and high-frequency render/Kitty diagnostics.
No query text, command text, history entry, or path is present in the retained
telemetry. The JSON results and traces beside this README are copied unchanged
from the same runner artifact.
