# I66 same-parent inline Rename host contract

The host captures an optional provider source token when its exact bind is unavailable
and a single inline Rename qualifies. The mutation interlock covers the captured parent.
The token can attempt a free sibling. A proven pre-dispatch destination collision can
offer Keep both or Cancel. The token cannot authorize replacement.

The first x64 Debug build found missing Dummy trace-provider linkage and four ownership
warnings. After correction, `0dc6b190a` built cleanly, but non-activating run
`20261125T121428Z-35536-bf7f2b38c2484e63a74b412f0ec62500` reached Dummy's ordinary
Rename refusal. The fixture now explicitly enables the typed route on its owned instance.

Full x64 Debug `7cb3e0b06` built with zero warnings and errors. Receipt
`7b16c133ae81fb54b4b7bc24de6eca958c9b63ecc765394fe6a1b176c1f03332`.

- `FileOps_ReviewRemediationContracts` run `20261125T123338Z-45856-3d37ee4484db4a82a956d62d2a1d2478`
  passed setup, case and cleanup without activation. The case checked actual retained bytes at the
  occupied name and published bytes at the Keep both name. It checked exact case-only spelling and
  terminal receipts for success, cancellation, and a source replaced before dispatch.
- `FileOps_ProviderCapabilityMatrix` run `20261125T123647Z-27324-21a1b24fb8494ce18d7ed8325a045723`
  passed three stages without activation. It preserves the default Dummy refusal and the existing
  negative unbound-route witnesses.
- The standalone provider contract suite passed in
  `20261125T123904Z-82640-89601f729cfe4427848a67a30b55188c` against the same verified build.

The review-remediation native suite took 91.8 seconds. The in-product trace records the four
Rename controls. Live providers, Batch Rename, Change Case and full native qualification are open.
