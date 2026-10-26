# ViewerSpace final worker-hook Debug Commands evidence

The filtered Commands self-test run `20261009T065738Z-42144-9524e13aa58746d2bb8841fc1076aae2` completed 31/31 cases passed, 0 failed, 0 skipped (`cmd_viewer_space_`, x64 Debug). This is focused Commands evidence, not a Full repository verdict or ARM64 application qualification.

- Build receipt: `687fe30cf9b5d5959000e9ed71af515fe41fe16237b74fc4ace7f9c45732d475`; build source snapshot: `b890ae0ddbf01b41819c00e27abd4c803a32a179860ac1d57da68da7266a80aa`; source HEAD: `5f14158528e51c4202fbce82be852cfb2319a891`.
- Validation workspace snapshot: `d0558ea74aa717ae96028bf3ad4445d9b1ce3398171f50bd768b9cbcb96d59d7`. This is a distinct record identity from the build receipt source snapshot.
- Preserved original build output reports 0 warnings and 0 errors.
- Root independently replayed the source metric stream: SHA-256 `f14dd5167081cad214b4873966fed6724d726aa54005386b10bfdc696502546e`, 34,775,999 bytes / 89,187 rows; the archive retains 3,727 selected rows with original line numbers bound in `evidence.json`.
- Root replay confirmed 67 CPU-priority apply/restore records with HRESULT 0, observed effective thread counts in the [1, 8] range, and the actual serial worker hook. These observations support worker-priority hook evidence; they are not a performance-gain claim.
- The exact raw capture-provenance source was compared before cleanup to `captures/raw-provenance.json`: SHA-256 `c7fb57a2b4fc227816e64c18ee4bff1a4689701bcc11d1923ffdfab06072b9a0`, 2,883 bytes. Its original D-drive run path is recorded in `build/evidence-manifest.txt`; that raw source path is now absent, while the checked archive copy remains available. Root accepted all 11 whole frames at 144 DPI for framing, theme, nested shapes, and the boundary contract; scope limits are recorded in `captures/review.txt`. AnalystBrowser tint remains ambiguous between hover and focus, and the images do not establish linked-selection state. The six threshold frames are in the sibling boundary-capture archive.

`build/evidence-manifest.txt` records hashes, lengths, and source paths/status for the immutable build receipt, original build log, and capture provenance. `.gitattributes` disables text normalization for this archive.
