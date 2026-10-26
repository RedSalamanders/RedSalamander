# Semantic-detail curated review link repair

The byte-exact gallery `visual-review-sol.md` links to `semantic-detail-v6-visual-review-sol.json`. The first curation had saved the JSON as `visual-review-sol.json.receipt.txt`, so that one gallery link failed. The updated curator saves the gallery JSON under the original linked filename and keeps the non-benchmark packet raw JSON under its `.receipt.txt` archival name. It omits the redundant packet `raw/visual-review-sol.md`: its unmodified link could not resolve beside the receipt-suffixed raw JSON. The gallery retains the original review Markdown and JSON bytes.

Reproducible repair for the already-curated output, if starting from the first curation state:

```powershell
Rename-Item -LiteralPath 'Z:/src/RedSalamander/Specs/Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-semantic-detail/visual-review-sol.json.receipt.txt' -NewName 'semantic-detail-v6-visual-review-sol.json'
Remove-Item -LiteralPath 'Z:/src/RedSalamander/Specs/TestRuns/4cb089111a23/Validation/2026-09-23_semantic-detail/raw/visual-review-sol.md'
```

Then replace the archived curator draft with the updated external script, regenerate both `SHA256SUMS.txt` manifests from current files, and verify staged Git payload hashes. Retain `raw/spec-inventory-initial-link-failure.log` as failure provenance; a later passing spec inventory does not erase it. Do not rerun curation into an existing destination.

Read-only audit after root's repair: the gallery now has `semantic-detail-v6-visual-review-sol.json` with the expected SHA-256 `0ce3f72b168f19aa7bdff130521a55626f3b318014bffd9d055caca05c555598`; the redundant packet raw Markdown is absent. All current relative Markdown links resolve: one in the validation packet and four in the gallery. This note did not change repository files.
