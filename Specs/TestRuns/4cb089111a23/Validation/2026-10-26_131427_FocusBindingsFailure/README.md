# Focus return keyboard-binding failure

At `a4209ee8cd7aa8f548f18706ca4a08c3bd272479`, all six current-source
builds passed. The original Commands replay passed 11 of 12 checks and failed
one two-pending-task iteration. The second unchanged iteration passed; that
repeat does not clear the failed first iteration.

All eight disclosure routes resolved and focused their exact task. After B
completed, returning from A's destination inspection had the correct foreground,
native focus, A control/generation, scope, disclosure and current unresolved
context, but both default/cancel bindings were empty. The complete trace records
this mismatch. Focus changes previously updated visibility alone; keyboard
bindings were only recomputed in a later frame. The proposed correction binds
from the last published task/generation during the focus transaction, deferring
reconciliation callbacks until the retained arrays are consistent. A new immediate
A/B focus check dispatches no frame between focus and binding assertions.

The original return assertions and timing limits remain intact. Full runtime
reruns remain omitted under the user's explicit updated scope. Earlier failed
Full receipts remain failed. `manifest.json` binds complete native verdict/trace,
the wrapper and failed scope state, the waiver, and all six build receipts.
