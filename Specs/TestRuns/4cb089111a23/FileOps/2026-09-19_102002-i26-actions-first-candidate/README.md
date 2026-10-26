# I26 first candidate — fixture correction required

Focused x64 Debug run `20260919T101655Z-98968-02e1cbfe1fd74d6abd72b2474ad6a1d8` on 2026-09-19.
The build passed with zero warnings/errors; its exact source/output identities are in the retained
receipt `6b6592b8f825380ed6321166a23e54343378b0ab84de707142b4158510bea4f1`.

The only production change reserved a 24-DIP ordinary task action row independently of conflict rows.
The same original fixture failed: Pause=false/false, Cancel=true/false, Speed=true/false (area/hit).
Setup and Cleanup passed. The test assumed Pause despite the eligible discovery action being Skip;
it also lacked a wait for the existing layout-motion flags. This attempt is retained for diagnosis,
not accepted as proof of correctness. Correct those preconditions and repeat both baseline and candidate.

Raw redraw batches: 132469, 120327, 127570, 132781, 133045 microseconds at 144 DPI. Do not use this
attempt to claim closeout or compare it with a different fixture revision. The independent sandbox
audit still reports 225 unrelated pre-existing run directories; none were removed.
