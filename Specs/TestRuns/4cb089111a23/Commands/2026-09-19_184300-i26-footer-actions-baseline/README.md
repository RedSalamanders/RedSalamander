# French footer — baseline

Run `footer-actions-baseline` fails 0/1, zero skips: at 480 DIP the minimum remaining label padding
is -23.463135 pixels, versus the required 36 pixels at native 144 DPI. Both live action labels are
measured; Queue/Options and action hit targets and button overlap are independently checked.
Build: zero warnings/errors. Sandbox audit: zero issues.

Receipt `04352294bc1d6e06ee17ea2a031c524b6e925f3a6d3f56f82a7eeb704cac81a6` binds workspace
`3174916b4b421655ddb8822eb3f0a1c69679924e6d3dab1bd72c03f0c3228932`. Test source SHA-256
`bb286a8b2b377c20c0afe2bcb2fc2ea598de270fd078b145b60516eb8beb2238` is identical in the
[final candidate](../2026-09-19_184300-i26-footer-actions-final/README.md).

All 20 eight-redraw samples are retained: five batches per width in the 480/640/760/480 sequence.
Median is 66,405.5 us, nearest-rank p95 67,320 us. This is same-machine x64 Debug, French, dark,
reduced motion, native 144 DPI. It is not frame pacing or Release/resource qualification.

Raw logs and artifacts remain under `C:/RedSalamander.Perf/evidence/i26-ui/footer-actions-baseline*`;
the completed sandbox run was moved to `S:/RedSalamander.Perf/evidence/footer-actions-baseline`
before the next run. Archived runner whitespace alone is normalized.
