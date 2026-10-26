# Focused native seed scratch admission

Source `676391cb3edb65c5b38ee4e3db8c0677678e9558` built x64 ASan Debug
and Release with zero warnings/errors. All 19 focused checks passed without
skips: six parser/storage checks, six ordinary native-MFT runs and seven service
regressions including the unchanged startup watchdog. These are focused passes,
not final repository qualification.

| Profile | Seed wall-time median ratio | Sampled private-memory median ratio | Verdict |
| --- | --- | --- | --- |
| ASan Debug | 0.278035 | 0.615274 | PASS |
| Release | 1.029977 | 1.014563 | PASS |

Baseline source remains `fac05d5b70b9a5933b12093c1b2337f76aee8cdc`.
All original timing/memory limits, native assertions and sanitizer settings remain.
Runs share the machine, volume, sampler and scenario; recorded MFT counts evolve
and do not represent a frozen volume image. `manifest.json` binds all 38 complete
raw payloads. Earlier failed pool comparisons remain failed in their own archives.

The next candidate adds explicit ASan boundary/lifetime poisoning within chunks
and verifies that visibility, together with the bounded final renderer-device
retirement repair. It requires fresh resource comparisons and all final builds,
Fresh Full profiles, ETW/ASan detection, public snapshot/CI and checkout cutover.
