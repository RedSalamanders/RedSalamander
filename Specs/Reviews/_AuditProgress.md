> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** The original audit date/commit is bounded by this file's scope metadata (or its paired `*-Audit.md`). Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.


# RedSalamander deep-audit campaign — COMPLETE

All 10 area audits done. Each was a multi-phase workflow (review units → adversarial verify → synthesize). Reports + compact findings lists saved here.

| # | Area | Confirmed | Plausible | Report |
|---|------|-----------|-----------|--------|
| 1 | File-system subsystem (file-ops bridge, FS plugins, index/search, contracts) | 69 | 32 | FS-Subsystem-Audit.md |
| 2 | Viewer plugins (Text/Hex/Sqlite/Space/ImgRaw/VLC/PE/Web) | 27 | 10 | Viewer-Plugins-Audit.md |
| 3 | Command system (registry/dispatch, handlers, dialogs, menus) | 49 | 18 | Command-System-Audit.md |
| 4 | Settings / config / credentials persistence (+ RedConfigure) | 37 | 14 | Settings-Persistence-Audit.md |
| 5 | FolderView drag-drop + selection + enumeration (OLE data movement) | 19 | 8 | FolderView-Audit.md |
| 6 | Monitor + ColorTextView (ETW + D2D editor) | 23 | 17 | Monitor-Audit.md |
| 7 | NavigationView (address bar / breadcrumb / path edit / menus) | 18 | 12 | NavigationView-Audit.md |
| 8 | DxUi framework (text/TSF/IME, grid, controls, menu, window host, a11y, theme) | 19 | 6 | DxUi-Audit.md |
| 9 | App shell / startup / single-instance IPC / crash / launcher / host | 15 | 10 | AppShell-Audit.md |
| 10 | Preferences dialog & panes (validation / persistence / list editors) | 30 | 6 | Preferences-Audit.md |
| | **TOTAL** | **306** | **133** | (582 candidates, 143 refuted/dropped) |

~135 review units, ~600 verifier+reviewer subagents. Every kept finding has `file:line` evidence and a fix; refuted claims were dropped at the verify stage.

See **_CampaignSummary.md** for the app-wide cross-cutting themes and the prioritized fix roadmap.
