# Issues primary grid: primary-row-baseline-v1

69 passed / 3 failed / 0 skipped; x64 Debug.
Expected failing baseline: all three repetitions expose a clipped first primary row at the minimum 520 by 320 DIP client size. The other 69 checks pass; there are no skips.

Every repeat retains separate actual application PNGs, captured without activation. The French
fixtures check common/differing subjects, complete copy values, qualified parent-folder comparison,
legacy path fallback, independent DirectWrite height at 520 DIP width, fixed footer and routed Home/End.
The new witness measures the actually visible first-row rectangle. Product layout is unchanged from 790616ca; only the snapshot and assertion were added before this baseline.
Routed input is not live keyboard/UIA acceptance. Final primary columns, native DPI/AT and Release/
resource qualification remain open. Projection baseline is the retained pre-detail V2 archive.
Raw: `S:/RedSalamander.Perf/evidence/i26-primary-row-baseline-v1-20260921`. Source/build snapshots match; compact metrics retain source line provenance.
