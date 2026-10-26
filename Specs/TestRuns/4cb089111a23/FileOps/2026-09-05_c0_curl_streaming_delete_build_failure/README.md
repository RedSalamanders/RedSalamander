# C0 Curl streaming Delete — first candidate compile correction

Starting from `5132d866`, the first cursor/iterative-walker candidate build for
`c0-curl-streaming-delete-debug-20260905` failed after 1:25. It emitted one C3861
error for a misspelled dot-name helper in the DOS parser and four implicit-deleted
copy/move warnings on the new ancestry frame. No build receipt or runtime
qualification resulted. The correction restores the canonical helper spelling and
explicitly deletes the frame's copy/move members; warnings are not suppressed.

`build.txt` preserves the full MSBuild log unchanged. The following candidate must
build without these diagnostics and run the native late-listing/deep/wide matrix.
