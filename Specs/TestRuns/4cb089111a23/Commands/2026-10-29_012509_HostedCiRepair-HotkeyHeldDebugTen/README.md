# Held hotkey: ten Debug repetitions

Source `7a3a7057669c5cd9d2ae5c629133f6cdebd23875`, full-solution x64 Debug
Embedded receipt `008c86712d1d58b549cbd4e206a53b2b0cca8d9eea2ab99e76d008548a9d3bb6`.
Run `20261029T012509Z-80644-3b552289dff84eab9a82370f80162b0d` passes all ten
repetitions with zero failures, skips or retries; canonical wrapper exit zero.
Native suite duration is 3,613 ms.

All twenty real WM_HOTKEY dispatch witnesses observe F24 still held before the
activation handler, and all twenty prove exact hook/property retirement. The
original foreground, key-release, repeated registration and refusal assertions
remain unchanged. The stronger test itself is unchanged from the retained
original-helper control, whose two dispatches both observed the key released.

This establishes event ordering and threshold acceptance locally. It is not a
throughput comparison, Full acceptance or native ARM64 hosted qualification.
The separate ASan repetition and fresh hosted profiles remain required.