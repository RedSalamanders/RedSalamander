# Removal-focus ASan reproduction before correction

Frozen RedSalamander `a20f6680`, x64 ASan Debug, reproduces a **heap-use-after-free**
in isolated case `cmd_pane_fileops_removal_focus_exact_outcomes_and_ownership_epochs`.
Run `20260913T072127Z-94288-cd4008c0872244d3ba428c3f2049a903` fails before successful case completion. No hosted CI ran.
The filtered suite uses the legacy aggregate; it does not publish a Full checkpoint.
Its fallback failed/skipped counts are not completed native-case results.

The ASan dump is 1,533,620,422 bytes and remains outside Git. Exact dump, executable
and matching-build PDB hashes are retained. The first debugger attempts lack Windows
symbols; their logs remain. The final attempt uses the Microsoft symbol server and
recognizes sanitizer exception `e073616e` and `ASAN Error: Use of deallocated memory`.
Its generic stack is not a resolved C++ allocation/free stack; no such stack is claimed.

Source review finds two retained sets of `FolderItem::displayName` views backed by
temporary vectors. The helper is byte-identical on master `0dd0bd6d` and `a20f6680`.
Correction `64403fe5` retains named backing vectors until their views are destroyed,
including teardown ordering in the large fixture. Product ownership is unchanged.
The corrected repeated case and final Full qualification are still required.

`retained-files.json` records 10 original small receipts and the inert executed
driver. The complete reproduction and adjacent dump remain under
`Z:/RedSalamander.Perf/evidence/I19-20260909/removal-focus-asan-a20f6680-before`.
