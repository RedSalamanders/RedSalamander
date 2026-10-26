# Status semantics witness source copies

These are external, reviewable copies of the three files under
`C:/RedSalamander.Perf/evidence/i26-ui/status-progress-witness-draft-20260923/inputs`.
The absolute input and output SHA-256 values, byte counts, and CRLF line counts are
in `source-hashes.json`; `witness.patch` is the corresponding unified diff. No
repository file was changed and no native build, gallery, or input was run.

The copies add test-only observations of the glyph actually submitted to
`DrawTextW` for each completed-group badge and the name of its actual hosted
tooltip region. The dark-theme scene 137 checks all five painted badge glyphs,
the distinct Partial/Indeterminate glyphs, and their localized hosted names at
both widths on each display. Dark 480-DIP scenes 18, 19, 31, and 136 check that
the painted Done chip coexists with an operation-only heading and no terminal
summary for task 1. The current product glyphs and headings are unchanged, so
these new assertions are expected to expose existing semantic failures until a
separate product fix is reviewed.

Progress witnesses use the retained hosted `ProgressBar` snapshot: scene 13
expects 48/128 = 37.5% determinate Delete progress; new scenes 141 and 142
expect exactly 0% determinate with closed known items or bytes, and 143 expects
indeterminate with no closed denominator. Scene 140 remains the corresponding
Copy unknown-denominator witness. The count is nine per display. The new Delete
fixtures keep the base operation start tick and closed discovery; both current
item byte counters are zero. Scene 139 is explicitly named synthetic text-fit
stress because that elapsed Preparing state is not a stable engine state.

The three copies are UTF-8 without BOM and CRLF, and touched regions were
formatted with the RedSalamander `.clang-format`. These files are proposed
inputs for the parent's source review and baseline/product qualification;
this artifact asserts no native pass.
