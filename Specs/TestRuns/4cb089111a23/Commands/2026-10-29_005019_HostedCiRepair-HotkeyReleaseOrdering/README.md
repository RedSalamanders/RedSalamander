# Original-helper hotkey ordering control

Base source `5cd2fa70464203d554dd41050bebb5eff390a6bf` plus the exact retained
test-only patch, full-solution x64 Debug Embedded receipt
`90ac8bed6ebfcb26a301bd6d97b53eccbfc5b0b3536e2e5fa4263455b1a4141e`.
The foreground helper itself is unchanged for this control.

Run `20261029T005019Z-52116-167656286d44465996a9b126bd26471d` fails only the
new key-state assertion: both real WM_HOTKEY dispatches observe F24 already released.
The original actual-foreground, original-procedure/property cleanup, key-release,
registration cleanup and refusal checks continue to pass. The case takes 32 ms;
native suite duration is 805 ms, with zero retries or skips and wrapper exit one.

This establishes premature key release. Older hosted ARM64 refusal is a separate
observation, not yet proof of its exact cause. Require the unchanged stronger witness
and fresh hosted qualification after correcting key-up ordering.