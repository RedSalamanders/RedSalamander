# Core Search Service Foreground Dashboard (FTXUI)

## Objective

Replace the hand-written foreground console dashboard in `RedSalamanderSearchService` with an `FTXUI` full-screen UI that gives useful live feedback during long searches and lets developers inspect service history interactively.

## Scope

- Link `RedSalamanderSearchService` against the `ftxui` vcpkg package already present in `vcpkg.json`.
- Replace the fixed repaint loop with an event-driven dashboard wake-up path so the terminal is not redrawn when nothing changed.
- Provide a full-screen overview page for live request state and a history page for browsing recorded service events.
- Keep redirected stdout/stderr behavior as readable lifecycle logs instead of trying to render the dashboard without an interactive console.
- Update spec and license documentation for the new dependency and dashboard behavior.

## Verification

- `.\build.ps1 -ProjectName RedSalamanderSearchService`
- `RedSalamander.exe --compare-selftest --selftest-case=search_service_status_and_query_roundtrip --selftest-fail-fast`
