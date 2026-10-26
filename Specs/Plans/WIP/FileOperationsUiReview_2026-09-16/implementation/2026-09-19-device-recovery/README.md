# Recovery and hidden-window captures

The final Debug harness run passes 6/6 cases with no skips. Three forced device recreations
and three hide/show cycles per repetition recover the same readable French client. Nine hidden
intervals add no hosted frames. [Receipts and all samples](../../../../../TestRuns/4cb089111a23/Commands/2026-09-19_213557-i26-device-recovery-idle/README.md)
retain the exact source/build identities. The initial 3/3 recovery-only run is retained separately.

![French popup after device recovery and hidden intervals](../2026-09-19-single-surface/throughput-uncovered__760dip__fr-FR__144dpi.png)

The three final-repetition recovery PNGs and shown-after-idle PNG are byte-identical to this
reviewed image; capture-manifest.json records each new raw source, digest and build provenance.
The shared PNG avoids storing duplicate bytes. All captures use the deterministic application
harness with no activation or desktop takeover. Native scale is 144 DPI; this is not mixed-DPI
qualification. The graph-label overlap visible here remains open, as do memory-soak, hardware
removal, idle CPU/wake-ups, Release and full-product acceptance.
