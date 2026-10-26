# Issues projection — actual application captures

The current screenshot still has the previous technical-first column layout. This component fixes
which records are retained and associated; it does not complete the selected-detail or default-column redesign.

- Baseline: 3 passed / 3 failed; four logical problems collapse to two visible rows.
- Candidate V2: 9/9, no skips; separate same-file prompt generations, legacy records and cached choices;
  stable selection, order-independent association, duplicate suppression, eviction and opaque-string safety.
- Existing long-scroll and selection regressions pass in all three repetitions.
- Full x64 Debug builds: zero warnings/errors. Static private-window harness captures; no activation.
- Captures show English UI resources and French synthetic problem/path values (currently outside the
  visible technical columns). This is not French layout or native mixed-DPI acceptance.
- Refresh median/p95: baseline 504/555 us with missing data; V2 865/1096 us processing complete data.
  V2 unchanged refresh 426/452 us, with no grid-generation change. Different visible work is explicit;
  three samples do not establish broad performance/resource non-regression. Full qualification remains open.

[Comparison gallery](gallery.html). Manifests bind each unchanged PNG to its harness source and build snapshot.
Curated archives are under `Specs/TestRuns/4cb089111a23/Commands/2026-09-21_034340-i26-projection-baseline-v1`
and `2026-09-21_035454-i26-projection-candidate-v2`; V1 is retained at `2026-09-21_035007-i26-projection-candidate-v1`.
