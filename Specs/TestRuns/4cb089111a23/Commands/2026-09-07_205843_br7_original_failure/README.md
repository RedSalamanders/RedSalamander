# BR-7 original Fresh Full failure

Recovered on 2026-09-08 from immutable governed entry evidence for run `20260907T205843Z-46028-f4bfabc0e38e458cad88a4d92654fcac`, x64 Debug, receipt `0fabf2c625e565828cf8905b54aaaea1e935a11c710454f4d0984650761ce11f`.

The complete original Commands results are retained unchanged. `cmd_pane_navigationView_full_path_popup_edit_route` failed after 3,687 ms: Escape closed the popup and edit surface, but native focus and focused FolderView were both null. The selected file, path, item count and selection count were preserved. The immediate predecessor was `cmd_pane_focus_follows_pointer_always_moves_focus_without_click`.

The original result did not record active/foreground HWNDs, so it alone cannot establish the exact activation interleaving. This is a baseline failure artifact, not evidence of a repair. The original Commands trace is no longer present in the immutable entry directory; no trace has been reconstructed.
