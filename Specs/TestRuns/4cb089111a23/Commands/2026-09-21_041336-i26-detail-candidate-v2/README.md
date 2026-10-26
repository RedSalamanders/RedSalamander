# Issues selected detail: detail-candidate-v2

9 passed / 3 failed / 0 skipped; x64 Debug.
Caption and overload corrected; child coordinate defect from V1 remains. Not accepted as final visual evidence.

Every repeat retains separate actual application PNGs, captured without activation. The French
fixture checks exact problem/paths, accepted replacement versus unknown result, source above destination,
independent DirectWrite text height at 520 DIP width, fixed footer and routed Home/End behavior.
Routed input is not live keyboard/UIA acceptance. Final primary columns, native DPI/AT and Release/
resource qualification remain open. Projection baseline is the retained pre-detail V2 archive.
Raw: `S:/RedSalamander.Perf/evidence/i26-detail-candidate-v2-20260921`. Source/build snapshots match; compact metrics retain source line provenance.

Follow-up source diagnosis: the adapter returned after host-handled WM_SIZE, so its child/footer
bounds did not reflow after resize. The caption probe's E_INVALIDARG identifies a region outside
the resized client, not missing font pixels or an established capture-service failure. The shared
host also consumed WM_NCDESTROY before adapter cleanup. V3 repairs this dispatch/ownership ordering.
