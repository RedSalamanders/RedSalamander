## Summary

- 

## Review Readiness

Complete the items that apply. Mark non-applicable items as `N/A` with a short reason.

- [ ] DxUi, window, dialog, menu, or text-input changes cover keyboard routing, focus restore, tooltip/hover behavior, UIA exposure, DPI/hit testing, and hidden-bridge contracts with focused selftests where behavior changed.
- [ ] Cross-thread UI, plugin callback, or window-host changes use `PostMessagePayload(...)` / `TakeMessagePayload<T>(...)`, initialize and drain posted-payload windows, and include a teardown/cancellation stress path when payloads can queue.
- [ ] Perf-sensitive changes define the scenario, metric, deterministic validation command, and archived `Specs/TestRuns/...` evidence, or state the blocker and follow-up owner.
- [ ] User-facing strings are localized through resources with positional placeholders; diagnostics use modern formatting and avoid `sprintf_s` / `swprintf_s`.
- [ ] New warning suppressions are local and documented with owner, scope, rationale, and expiry/review condition.
- [ ] Durable behavior, UI contracts, validation rules, and workflow requirements are reflected in the authoritative spec, not only in a WIP or PR note.

## Validation

- 

