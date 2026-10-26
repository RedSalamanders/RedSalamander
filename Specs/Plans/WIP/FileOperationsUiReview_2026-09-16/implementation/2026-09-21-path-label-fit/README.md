# French path labels — 2026-09-21

Actual application captures from the non-activating `FileOps_VisualGallery`
harness, dark theme, French, 480 DIP, native display 1 at 144 DPI and display 2
at 96 DPI. These are implementation evidence, not proposal mockups.
Full scenario/build provenance and paired results are in the
[measurement archive](../../../../../TestRuns/4cb089111a23/FileOps/2026-09-21-path-label-fit/README.md).

| Scenario / DPI | Baseline | Candidate |
|---|---|---|
| Delete / 144 | [before](baseline--13-delete__dark__480dip__fr-FR__display1__144dpi.png) | [after](candidate--13-delete__dark__480dip__fr-FR__display1__144dpi.png) |
| Delete / 96 | [before](baseline--13-delete__dark__480dip__fr-FR__display2__96dpi.png) | [after](candidate--13-delete__dark__480dip__fr-FR__display2__96dpi.png) |
| Verify / 144 | [before](baseline--17-verifying__dark__480dip__fr-FR__display1__144dpi.png) | [after](candidate--17-verifying__dark__480dip__fr-FR__display1__144dpi.png) |
| Verify / 96 | [before](baseline--17-verifying__dark__480dip__fr-FR__display2__96dpi.png) | [after](candidate--17-verifying__dark__480dip__fr-FR__display2__96dpi.png) |

The original fixed column clipped Vérification and let Suppression overlap the
path. The candidate measures the localized label and gives the path the remaining
width, preserving full-path inspection and UIA identity. Four bounded cached
measurements avoid repeating DirectWrite layout work on each frame. Deletion's
processed counter now says Supprimés, preserving the distinction from transfer.

The candidate gallery passes actual label bounds for delete, verify and completed
cards at 480/760 DIP on both displays. The screenshots still expose separate open
P3 issues: verification's footer is full before verification finishes, and delete
reserves a byte-throughput graph without useful byte samples. This slice does not
claim final UI acceptance or close those issues.
