# ViewerSpace future workspace prototype

Status: non-normative interactive proposal, 2026-10-22. This is a browser prototype, not the running Windows application or evidence of implemented native behavior. All paths, names, totals, dates and scan rates are invented fixture data. There are no live providers, file operations or saved scans.

The single [ViewerSpace roadmap](../../Plans/WIP/ViewerSpace_ConsolidatedRoadmapAndQualification_2026-10-22.md) owns the agreed product scope. This folder adds a reviewable interaction model without closing its implementation or qualification work.

## Explore

- The large Discovered figure remains scan-wide. Navigating or excluding items adds a separate This folder / Shown figure.
- Click a map folder header to explore it. In the ranked list, click to inspect; double-click or press Enter on a focused folder row to explore. Breadcrumbs, Back, Forward and Up work.
- Click a row's file/folder icon to exclude it from map area, shown totals and percentages. Its real size remains visible in a dimmed row with an overlay. Restore individually or use Restore all. Exclusions survive navigation and simulated refresh within this viewer.
- Name, type and age filters narrow the current-folder list and highlight map branches without changing map sizes. Summarized detail stays explicit.
- The single Layout icon offers Panel right, Panel bottom and Map only. Ctrl+1 / Ctrl+2 select the visible arrangements. A narrow window temporarily uses the bottom arrangement.
- Area switches between normal Size, File count and known Size on disk. Unknown allocation is never replaced by logical size. Color offers folder groups, file type, own Modified date and newest known contained-file date, with a legend.
- The inspector shows one name, its parent location and compact facts. Meaningful incomplete/excluded/refresh states replace routine Done/Kind fields. Exact values are available through details and tooltips.
- The list and map context menus, Selection actions and File menu expose Copy path. Alt+Insert uses the same real item; text-field Copy keeps its usual meaning. A clipboard failure provides a selectable path. Synthetic aggregates have no copyable invented path.
- Open folder in other pane has an explicit menu action and no added shortcut. The host handoff is simulated.

## Cleanup and secondary flows

Cleanup checkboxes are separate from exclusion icons. Mark matching items snapshots included real rows in the current filter; nested targets are deduplicated. Review displays whole-folder consequences, allows removing targets, and simulates sending the chosen operation to File Operations. It cannot delete or recycle real files.

The Problems count opens unreadable-folder and summary explanations with Locate. Export previews CSV/JSON for current filtered entries, including excluded, summarized and unavailable facts; Download saves only invented prototype data. The prototype's View settings popup is a proposal artifact superseded by the accepted native placement: global Preferences → Plugins → Space, with technical fields under Advanced. It does not change application settings.

The top strip belongs only to the prototype: scan-state fixtures, theme preview and Reset demo. Scanning and refreshing are deterministic presentation states, not measurements. Escape dismisses a popup/filter, then cancels a running scan or closes an idle simulated viewer; Reopen prototype restores the page.

## Local development and validation

The local preview is served at `http://127.0.0.1:4173/` while its development process is running. It is not deployed.

```powershell
npm ci
npm run dev -- --port 4173 --strictPort
npm run build
```

React/Vite and Fluent UI React Icons are recorded in the lockfile. Product UI lives in [App.jsx](src/App.jsx), layout tokens in [styles.css](src/styles.css), and synthetic accounting/map geometry in [model.js](src/model.js). Hosting scaffold files are retained from the Product Design template; no Sites publishing was requested.

See [design QA](design-qa.md) for inspected states, screenshots, corrections and evidence boundaries. These browser captures are labelled proposal mockups; they are not native application screenshots and do not replace the governed synthetic Windows test-harness captures required for implementation closeout.
