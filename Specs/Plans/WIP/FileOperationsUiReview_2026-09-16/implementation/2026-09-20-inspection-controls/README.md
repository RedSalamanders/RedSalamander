# Pending-conflict inspection controls

Bounded R3 implementation, 2026-09-20. [Native gallery](gallery.html), [capture provenance](capture-manifest.json), [measurements](paired-measurements.json).

Source and destination now have separate full-width inspection buttons beneath their metadata. The action uses the painted immutable prompt, navigates the last active pane through the existing owner transaction and selects the exact role item. It never decides the conflict or overwrites destination bytes. Complete French outcomes distinguish selection, missing/filtered/renamed items, ambiguity, unavailable identity, stale requests, restored navigation failure and rollback failure. UIA and keyboard use the same hosted action; results have accessible text and current-prompt notifications.

The popup still stays visible after inspection. Guarded pane foreground focus, hide/return to the invoking control and the broader provider/interaction matrix remain unfinished. The pinned UIA Button invocation itself establishes popup/button focus; the tests check that navigation does not move it again and separately cover non-activating harness dispatch. Do not describe this as complete R3 acceptance.

The accepted Debug matrix is **18/18**, with zero warnings/errors/skips or sandbox audit issues. It covers both role/pane combinations, real UIA invocation, missing item, stale painted prompt refusal, replacement-result removal and unchanged exact destination bytes. Real retained button typography plus padding fits both French labels at 480/640/760 DIP. Dynamic result text retains its complete measured height immediately and precedes the scope checkbox. The first selected-result capture exposed an overlap during card-height animation; expanded decisions now use their measured height without interpolating through a smaller old card. The rejected frame is retained.

The gallery contains 10 accepted current captures, two earlier comparison frames and one rejected frame. Four accepted frames and the rejected frame were directly inspected. The new French labels and result sentences are complete in those frames. **Existing narrow primary-action clipping and short-viewport footer/action overlap remain visible and unaccepted under R2.** Result captures do not qualify physical mixed-DPI transitions or native assistive technology. A full final catalog remains required.

Performance remains open: baseline common median/p95 **66496/66865 us**, differing **66547/66755 us**; accepted **66583/66852** and **66588/88361**. One unchanged-source confirmation gives **66620/79436** and **66504/81740**, 45 samples per presentation. Do not treat stable medians as a tail-latency pass or erase outliers. Profile CPU/layout versus presentation waits and retain a matched baseline before final resource acceptance. These are Debug redraw batches, not Release FPS or resource/lifetime evidence.

- [candidate](../../../../../TestRuns/4cb089111a23/Commands/2026-09-20_170357-i26-inspection-controls-candidate/README.md)
- [final](../../../../../TestRuns/4cb089111a23/Commands/2026-09-20_171005-i26-inspection-controls-final/README.md)
- [verified](../../../../../TestRuns/4cb089111a23/Commands/2026-09-20_171557-i26-inspection-controls-verified/README.md)
- [accepted](../../../../../TestRuns/4cb089111a23/Commands/2026-09-20_172529-i26-inspection-controls-accepted/README.md)
- [perf-confirmation](../../../../../TestRuns/4cb089111a23/Commands/2026-09-20_173013-i26-inspection-controls-perf-confirmation/README.md)

The earlier candidate/final runs are 15/18 and retained: the first assertion counted the library's explicit UIA invocation focus as a late navigation change; the second used an ifstream reader that cannot share a retained DELETE-capable conflict binding. The existing fixture reader now uses bounded shared native reads. A subsequent layout test edit failed compilation before execution; its log is retained with the accepted receipt. No failed run is promoted to a pass.

Source snapshot: `25db5fe194861741fedcbbd103fd84dadcec7c420add73775c4c623aba680550`. R2 publication/exact-main CI/pin adoption, R3 guarded focus/return, Release/Fresh Full, native mixed-DPI/keyboard/AT and paired resource acceptance remain open. Both plans remain ACTIVE.
