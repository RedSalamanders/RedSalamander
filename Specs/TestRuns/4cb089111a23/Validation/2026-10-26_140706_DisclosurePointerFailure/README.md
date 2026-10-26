# Original ASan closure and retained Debug pointer failure

At `7bb6990285ee0027eddde7bb8e68cbe2732c61b4`, all six build profiles
passed. The original Commands replay passed 12/12, original search passed 3/3,
and FileOps selection plus mandatory setup/cleanup passed 3/3, with no skips.
These native receipts cover all eight original ASan failures. The immediate
A/B binding assertions and original return assertions both passed in ASan.

The subsequent Debug order passed theme retention and breadcrumb lifecycle,
but the two-task case failed at its real-pointer route. The exact B disclosure
provider and focus existed, yet its UIA bounding rectangle was zero-sized.
A focused offscreen task deliberately retains its controls/provider. The exact
lookup had incorrectly treated that provider's existence as proof of visibility,
then SetFocus on the already-focused control did not request another scroll.

The fixture correction requires a nonzero measured disclosure rectangle before
returning from the existing bounded outer-scroll lookup. Other action lookup
semantics stay intact because comparison actions can require inner scrolling
on focus. The existing eight-page bound, all real pointer/input assertions,
original timing limits and task/generation predicates remain unchanged. The
pre-reveal visibility trace now measures bounds rather than provider existence.
No production code changes follow this packet. The failed Debug receipt remains
failed and the fixed-source order must pass before publication. Full runtime
reruns remain explicitly omitted under the user's updated scope.
