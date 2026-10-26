# Live popup caption after a language change

The same File Operations popup now reloads its native title when the running
application applies a new language preference. The retained conflict decision and
inspection controls stay on the existing window. This is a focused localization
repair, not complete live relocalization of the Issues pane or a final I26 pass.

| Harness capture | Observed title | Source |
|---|---|---|
| [Before, returned conflict](before__returned__fr-FR.png) | `File Operations` over French content | `fileops-family-settled-20260921`, build receipt `75525f9f6a1bb0166721fcafe4774c9459f44d4c7e180988b26ffb859aa772b4` |
| [After, returned conflict](returned__fr-FR.png) | `Opérations sur les fichiers` | `i26-live-caption-inspection-20260921`, build receipt `dd1423f609f589d9c0c266590ccd80498debe23c150dbd9940a5cc890516e410` |
| [After, short focused conflict](short-body-focused-action__fr-FR.png) | `Opérations sur les fichiers` | Same candidate run |

These PNGs are actual application windows captured by the deterministic Commands
test harness; no desktop screenshot or proposal mockup is mixed in. Before and
after use the same named conflict-inspection fixture, though run-specific sandbox
paths differ. The candidate source was based on commit `f2f3e1f584f76fa7741a068651a3bac41e1f5d7d`
with the live-caption code/test/spec changes in this slice. The pinned DxUi commit
remained `78b3de389a189c7f86f611787e0489fb6d474218`.

The test changes language to en-US then fr-FR on the same HWND, dispatches the
settings-route `WM_THEMECHANGED` notification, and checks the localized caption
after each change. The three-case conflict-inspection predecessor set passed
**9/9** over three repetitions, with no skips; the exact
[runner result](run-all-tests-results.json) is retained. The full test-enabled
x64 Debug solution build passed with zero warnings/errors. Raw build/test logs,
per-case results and original harness captures are retained under
`C:/RedSalamander.Perf/evidence/i26-ui/` and run
`C:/RedSalamander.Perf/runs/i26-live-caption-inspection-20260921`.

The short-height capture still shows a scrollable conflict body. The footer
`File` is the French translation of Queue (`IDS_FILEOPS_BTN_MODE_QUEUE`), not an
English file label. This caption slice does not qualify that menu interaction,
native mixed-DPI, Release, AT, resources or the remaining I26 interaction matrix.
The separate [retained Issues language slice](../2026-09-21-issues-live-language/README.md)
now covers its cached headings/buttons and native title.
