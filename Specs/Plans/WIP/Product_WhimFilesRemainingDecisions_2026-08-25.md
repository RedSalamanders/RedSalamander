# WhimFiles Remaining Product Decisions

> **NON-NORMATIVE DECISION QUEUE.** Competitive analysis is frozen at `Specs/Plans/Done/Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md`. This file lists only unresolved routing. It does not authorize implementation. Do not treat Waves 1–5 as a schedule. Do not recreate root `plans/`. Historical undo spike: `Specs/Plans/Done/AdvisorPlan_013-undo-file-operations-spike_2026-06-11.md` (not an owner).

## Status

- **State:** DECISION
- **Priority:** P3 unless a row is promoted by its named owner
- **Planned at:** `56db9af9d0a19b4b3f375e16989591f29e014a6c`
- **Ownership boundary:** unresolved G\* product choices with no other live owner; this file cannot change production behavior
- **Drift check:** `git diff 56db9af9d0a19b4b3f375e16989591f29e014a6c..HEAD -- Specs/UI/UI_FolderView.md Specs/UI/UI_NavigationView.md Specs/FileSystem/FileSystem_FileOperations.md Specs/Plans/WIP/Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md Specs/Plans/WIP/README.md`
- **Authoritative specs:** `Specs/UI/UI_FolderView.md`, `Specs/UI/UI_NavigationView.md`, `Specs/FileSystem/FileSystem_FileOperations.md`, Atlas D0

## Remaining unresolved routing

| Gap | Status | Live owner or hold |
|---|---|---|
| G1 Live multi-dimensional filtering | Unresolved product slice | **D1 until promoted** |
| G2 Undo for file operations | Authority settled 2026-08-21; implementation deferred | File Operations spec + **I11** (do not edit I11 from this file). Undo may use only exact identities / Proven Engine artifacts; no pathname-set journal. `plans/013` is not an owner. |
| G3 Tabbed browsing | Unresolved FolderView presentation slice | Atlas `ATLAS-DEC-FV-01` |
| G4 True command palette | Unresolved product slice | **D1 until promoted** |
| G5 Fuzzy Go to Folder / File | Unresolved product slice | **D1 until promoted** |
| G6 File checksums in Properties/command | Unresolved product slice | **D1 until promoted** |
| G7 First-class HEIC/WebP/AVIF + batch convert | Unresolved product slice | **D1 until promoted** |
| G8 Bookmarks / Places sidebar | Unresolved NavigationView slice | Atlas `ATLAS-DEC-NAV-01` |
| G9 Hover preview / Quick Look | Unresolved FolderView presentation slice | Atlas `ATLAS-DEC-FV-01` |
| G10 Always-calculate folder sizes | Unresolved product slice | **D1 until promoted** |
| G11 Open-in-app terminal helper | Unresolved product slice | **D1 until promoted** |
| G12 Customizable button bar | Unresolved product slice | **D1 until promoted** |
| G13 Flexible pane layout / extra panes | Unresolved product slice | **D1 until promoted** (adjacent to `ATLAS-DEC-FV-01` / `ATLAS-DEC-WIN-02`) |
| G14 Directory tree pane | Unresolved NavigationView-adjacent slice | **D1 until promoted** (adjacent to `ATLAS-DEC-NAV-01`) |
| G15 Persistent command-line bar | Unresolved product slice | **D1 until promoted** |
| G16 File-type color coding | Unresolved product slice | **D1 until promoted** |
| G17 Split / combine | Unresolved product slice | **D1 until promoted** |
| G18 Encode/decode + checksum files | Unresolved product slice | **D1 until promoted** (shares hashing with G6) |
| G19 Tags / labels / ratings / comments | Unresolved product slice | **D1 until promoted** |
| G20 User scripting | Unresolved product slice; broader pane ABI is Atlas | **D1 until promoted**; public ABI breadth is `ATLAS-DEC-HOST-01` |
| G21 Duplicate finder | Unresolved product slice | **D1 until promoted** (would consume File Operations if promoted) |
| G22 Multi-column grid + metadata columns | Unresolved FolderView presentation slice | Atlas `ATLAS-DEC-FV-01` |
| G23 Collections / saved searches | Unresolved product slice | **D1 until promoted** |
| G24 Power-user long-tail bundle | Unresolved product slice | **D1 until promoted** |

Promotion rule: D0, D1, and H1 cannot change production until one approved slice moves into a scoped `ACTIVE` owner. Do not create new ACTIVE product plans for G1–G24 from this file.

## Still-open product questions (not execution)

From the archived analysis, these choices remain unresolved and stay here until a maintainer records accept/reject:

- Tab-jump chord vs Hot Paths `Ctrl+1..0`; `Ctrl+P` palette vs go-to-file.
- Filter persistence (per-pane / per-tab / named saved filters).
- Checksum host (themed Properties vs standalone dialog).
- G2 retention: session-only vs persist across restart (implementation still deferred; I11 owns FO identity/recovery axes).
- G19 storage (SQLite vs ADS); G22 grid strategy; G20 engine; G23 as `collection:` VFS; how far up-market (G19/G20/G22).

## Verification

```powershell
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

This queue is correctly classified only when it is indexed once as `DECISION` and no agent implements a G\* row from this file or the Done archive.

## Done criteria

Every row is either promoted into a named ACTIVE owner, rejected with rationale, or still explicitly **D1 until promoted**. Competitive analysis remains in Done. This file moves to Done when no unresolved D1-owned row remains.

## STOP conditions

Stop if implementation would start from this file or the Done archive. Stop if a row would duplicate I11 File Operations work (point at I11; do not edit I11). Stop if root `plans/` is recreated. Stop if G2 is treated as an implementation owner.
