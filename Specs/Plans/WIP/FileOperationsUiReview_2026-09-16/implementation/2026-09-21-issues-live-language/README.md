# Live language changes in the Issues pane

[Open the before/after screenshot gallery](gallery.html).

An already-open Issues pane now updates its title, primary and technical column
headings, action captions, localized row/detail labels, hidden-sort explanation,
empty-state text and accessible detail name together. The same retained grid keeps
its selected issue, semantic sort and column preferences. Original diagnostic
messages and provider paths remain recorded facts.

| View | Before | Revised actual application |
|---|---|---|
| Switch to English on the existing pane | [Mixed French and English](before-primary-en-US.png) | [English labels](primary-en-US-live-language.png) |
| Switch back to French | Same-window return after the English switch | [French labels](primary-fr-FR-live-language.png) |
| Technical view | Cached title/headings/actions | [English](technical-en-US-live-language.png), [French](technical-fr-FR-live-language.png) |

The first view intentionally keeps its original French diagnostic message while
changing presentation labels to English. This distinguishes recorded data from
current UI text. The filename stays prominent, and technical columns remain an
explicit opt-in with their saved order and widths. The complete selected detail
is independently scrollable; these images do not claim that every technical table
cell fits without truncation. The native caption still uses the existing
“failed items” wording; a neutral title covering pending and resolved conflicts
remains a separate wording correction, not a claim of this language-switch fix.

These five reviewed PNGs come from the deterministic application test harness,
not desktop capture or mockups. `capture-manifest.json` binds their exact source,
hash, pixel dimensions, scenario, language and build receipt. The final repeated
fixture passes 3/3; its four language/presentation transitions per repetition
preserve selection, logical focus, sort and column bounds. The File Operations
command family passes 50/50 with no skips. Standalone localization tests pass all
76 assertions, including successful/same-language/rejected revision semantics.

The initial repair introduced redundant work on unchanged theme notifications.
The final version uses the shared localization revision to skip translation work
unless language resources have been reapplied. Five unchanged notifications have
a measured 2.011 ms median versus 2.220 ms before; actual language changes now do
the required extra row/label refresh. This is a focused timing comparison, not a
claim of complete resource, frame-latency, native mixed-DPI, Release or AT acceptance.

The [retained failure/candidate sequence and paired measurements](../../../../../TestRuns/4cb089111a23/Commands/2026-09-21-issues-live-language/README.md)
preserve the raw results and exact provenance.
The I26 parent plan remains active for its remaining implementation and gates.
