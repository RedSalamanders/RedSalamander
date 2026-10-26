# Warm quiescence queued-paint control

This archive records a native helper defect, not a production performance gain.
At source `986e623325a66135553990d5f415f140bf23fb7f`, the unchanged helper returned
quiet with a real paint timer still queued in trial five. Earlier trials dispatched
their paint. The separate window-owned timer closes its RAII owner before painting
the real pane; its positive control completes while the borrowed state is alive.

The original source `47d22a521aeb2d0c57d1f042e5d640e3735986af` Full remains failed
because its scroll scenario observed one repaint for an unchanged horizontal
request. The exact repaint cause is unproven. Its same-binary predecessor replay
and five repeats passed; those are diagnostic observations, not a repair receipt.

The correction drains work after the observation delay and checks for remaining
messages before reading counters. The scroll fixture, zero-paint assertion,
measurement span, 100 ms stable interval and 5-second timeout remain unchanged.
Baseline result bytes, trace observations, source and artifact identities are
retained beside this note. Candidate qualification is a separate gate.
