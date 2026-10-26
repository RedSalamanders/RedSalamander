# Semantic Issues detail v6 visual review

Compared **all 20 changed common PNGs and 6 new PNGs** from the actual Release gallery at 96 and 144 DPI against the retained direct-Issues baseline. The other 512 common PNGs are pixel-identical according to `capture-comparison.json`. There are no dimension changes. Full filenames, hashes, dimensions, and bounds are in [the review JSON](semantic-detail-v6-visual-review-sol.json).

No newly clipped or overlapping selected-detail status text or bottom buttons was visible in the 26 reviewed frames. The candidate paints warning markers for problems/history, blue decision arrows, red failed results, and green confirmed results. The six new 129b/129c/129d captures show a selected **Terminé** row and matching green result detail for publication, Delete, and Rename on both displays. Where no decision was supplied, **Décision non enregistrée** remains explicit.

The new Delete and Rename frames use synthetic diagnostic data. **129c Delete retains a Copy destination** (`D:\Sauvegardes\Portrait original.raw`), and **129d Rename uses a cross-folder same-leaf destination** instead of a same-folder new leaf. Their green checks verify typed result/icon mapping in the harness; they are not executed filesystem operations or representative final documentation captures. A later qualified gallery should correct those two fixture paths. This limitation does not change the v6 glyph witness result.

The 144-DPI accepted/failed/partial/cached long-detail views show Source beginning at the bottom of a scroll viewport. A scrollbar is present and both bottom actions remain unobstructed; the captures do not demonstrate that every below-fold fact is visible without scrolling. This is a visible limit, not evidence of a new overlap.

Review method: all 26 original PNGs were inspected in side-by-side contact sheets generated solely for analysis; the new 144-DPI originals and new 96-DPI publication original were also opened at native resolution. These are dark French harness scenes only. UIA, keyboard/focus, live interaction, Debug, other themes/locales, and ordinary performance require separate evidence.

| Reviewed changed scene | 96 DPI | 144 DPI |
|---|---:|---:|
| `105-issues-selected` | ✓ | ✓ |
| `106-issues-scrolled` | ✓ | ✓ |
| `123-issues-legacy-missing-facts` | ✓ | ✓ |
| `124-issues-accepted-pending` | ✓ | ✓ |
| `125-issues-failed-replacement` | ✓ | ✓ |
| `126-issues-partial-move` | ✓ | ✓ |
| `127-issues-cached-scope` | ✓ | ✓ |
| `129-issues-reappeared-identity` | ✓ | ✓ |
| `130-issues-history-selected` | ✓ | ✓ |
| `132-issues-history-reintroduced` | ✓ | ✓ |

| Reviewed new scene | 96 DPI | 144 DPI |
|---|---:|---:|
| `129b-issues-confirmed-publication` | ✓ | ✓ |
| `129c-issues-confirmed-delete` | ✓ | ✓ |
| `129d-issues-confirmed-rename` | ✓ | ✓ |
