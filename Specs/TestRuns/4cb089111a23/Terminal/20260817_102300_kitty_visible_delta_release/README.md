# Terminal bounded Kitty visible-delta cache and placement accounting

- Status: passed
- Build: test-enabled Release x64 (`RSBuildEnableTests=true`)
- Scenario: `Terminal bounded Kitty visible-delta cache and placement accounting`
- Command: `.build\x64\Release\PluginContractTests.exe --terminal-selftests`
- Samples: five independent processes; every process passed `69/0`, exited 0,
  reached the Terminal unload quiet point, and wrote no stderr bytes.
- Runtime shape: 4,096 retained placements (1,024 off-screen, 1,024 virtual,
  1,024 visible distinct-image placements, and 1,024 visible repeated-image
  placements), 1,025 visible exact image keys, 598,040 placement-map bytes, and
  one rejected `+1` placement.
- Cache reuse: the second unchanged-geometry capture submitted no conversion
  work and retained 4,100 cached/pinned BGRA bytes with zero missing, deferred,
  or rejected visible keys.
- Lookup comparison: the test-only model of the removed vector search performed
  529,915 capture and 530,940 draw comparisons per process. Production performed
  exactly 2,048 keyed draw lookups, a 99.614% comparison-count reduction. Median
  paired lookup time was 117 us for the legacy model and 24 us for production.
- Lock bounds: first missing-key capture was 414-594 us (454-us median); the
  ready-cache reuse capture was 205-230 us (219-us median).
- Existing maximum-size fixture: maximum 4,096x4,096 conversion was
  28,950-30,375 us (29,098-us median) and resident high-water remained exactly
  134,217,728 bytes. The surviving same-machine historical archive recorded
  28,834 us in one process; the candidate median is +264 us (+0.9%). The older
  archive is a one-sample directional reference, not a five-sample percentile.
- Caveat: this focused plugin fixture uses the real staged Ghostty runtime and
  real host capture/cache code but has no D2D render target, so zero upload rows
  are correctness-only. It makes no product upload-latency claim; injected and
  real-render upload recovery evidence belongs to L4-TERM-04.
- Raw evidence: `perf/perf_metrics.jsonl` contains 445 content-free rows and has
  SHA-256 `b89a35e4b44a8d019efae590ea27cce8d4b92430cd6b73f1985db6bc1a8aff61`.
