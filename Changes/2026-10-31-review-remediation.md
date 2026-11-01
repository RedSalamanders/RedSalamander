- Harden file operations and viewer lifetimes after revalidating the October merge reviews.
  - Preserve exact Local publication/abort authority, long filenames, link targets and Hidden attributes;
    use direct non-atomic MTP creation with truthful partial-publication and Move-source outcomes.
  - Fix worker host-service dispatch, cancellation wakeups, directory-cache cancellation, interrupted-Move
    notice ownership, conservative leftover cleanup and concurrent Preferences changes.
  - Preserve Windows ordinal path identity for supplementary-character filenames in cached overlap
    checks and required rename keys.
  - Dispose ViewerSpace's private cache at Close, retain visible unreadable-subtree markers, and improve
    viewer navigation, loading, localized counts, popup minimization, path-detail feedback and Issues export.
  - Pin formatter behavior, exclude build/test sidecars from installers, preserve historical evidence while
    shortening checkout paths, and strengthen scoped test verdicts and failure diagnostics.
  - Check offline dependency constraints against immutable pinned registry data while preserving existing
    managed checkouts and keeping live tool/bootstrap acceptance separate.
