# Observatory Track 5 — Unpack prompt conflict policy

- Date: 2026-07-16
- Machine profile: `4cb089111a23`
- Case: `cmd_pane_unpack_prompt_uses_dxui_destination_unpacker_and_mask`
- Result: 1 passed, 0 failed, 0 skipped

This focused run proves the DxUi Unpack prompt defaults to skip-existing, exposes replace-existing as an explicit
apply-to-all policy, and retains Cancel as the dialog cancellation path.

The authoritative Track 5 closeout record is
`../2026-07-16_224010/README.md`.
