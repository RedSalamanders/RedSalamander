# Guarded inspection focus and return

[Gallery](gallery.html) · [Provenance](capture-manifest.json) · [Measurements](paired-measurements.json)

Actual source/destination inspection now selects and focuses the exact item in the captured active pane, hides the popup through its existing lifecycle, and returns through the ordinary File Operations command to the same unresolved card, scroll, scope and invoking control. The command now establishes actual foreground/keyboard focus; automatic presentation remains non-activating. A consumed one-shot ticket validates live prompt, post-selection visit, provider context, active pane and unchanged input before focus and across Windows reentry. Background navigation never steals focus or hides the popup. No operation decision or destination mutation follows from inspection.

The **18/18** x64 Debug matrix has zero warnings/errors/skips/audit issues. Both UIA roles/panes and keyboard Space/Enter traverse the real flow. A discovered Space shortcut interception is repaired: decision/inspection controls own their keys. Changed focus refuses and consumes its ticket; replay fails. Missing outcomes, stale painted prompts and original destination bytes remain checked. Reopening preserves default/cancel actions and exposes the focus ring in the keyboard capture.

The first candidate passed internal-focus assertions but its returned window was inactive. That frame is retained as **before**, and the test now requires actual foreground/focus. Two returned frames were directly reviewed. Current selected/missing frames are retained with provenance but not newly visually accepted. The final keyboard return frame has complete French labels and a visible source-inspection focus ring. Narrow action clipping and short-viewport overlap remain separate R2 work; this is not the finished UI gallery.

Renderer common/differing p95, microseconds per eight-redraw batch: candidate 66788/66824; foreground/scrollbar correction 66859/66859; keyboard final 66788/66782. Forty-five samples per presentation. The original control slice's 88361/81740 tails remain in its retained evidence; these results do not erase them or prove a cause. Redundant native scrollbar updates are skipped only when actual range/page/position is unchanged. Focus costs are new-feature measurements, not paired non-regression evidence. Native DPI/AT, resource/Release/Fresh Full and qualified layout-pin adoption remain open.

Source snapshot: `a9fa1560a4575edfbc869e059d25919257a0c56f08d6a17b3b883455097c8edd`.
