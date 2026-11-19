## Summary

- 

## Validation

- 

## Checklist

Mark an item that does not apply as `N/A` with a short reason.

- [ ] One `Changes/` fragment describes the change, and `CHANGELOG.md` is untouched.
- [ ] Durable behavior, UI contracts, validation rules and workflow requirements are in the owning spec, not only
      in a plan or in this description.
- [ ] Every native profile the change affects builds with zero warnings, or the change is prose only.
- [ ] New warning suppressions are local and documented with owner, scope, rationale and review condition.
- [ ] Window, dialog, menu or text-input changes cover keyboard routing, focus restore, tooltip and hover
      behavior, UI Automation, DPI and hit testing, with focused self-tests where behavior changed.
- [ ] Cross-thread UI, callback or window-host changes publish on the UI owner through the repository's
      posted-payload helper, cancel by request identity, reject stale completion, and drain queued payloads at
      teardown.
- [ ] Performance-sensitive changes name the scenario, metric and deterministic command and archive paired
      evidence, or state the blocker and its owner.
- [ ] User-facing strings come from resources with positional placeholders; diagnostics use modern formatting.
