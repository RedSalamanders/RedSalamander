# I66 Microsoft Drive Rename receipts

The x64 Debug PluginContractTests project union built at `8751070f8` with no warnings or
errors. The verifier checked the selected build manifest and the provider DLLs. Receipt
`4a490e536eaf99414cfebffdddfa670708236d65bf65b76daa45066380793e2e` ran in
`20261125T113539Z-62976-dc59cafd69c24a0890b199aadc8fae48` without desktop activation.
The standalone suite passed. Microsoft Drive reported 259 passes and zero failures.

The new public entry-point checks require zero PATCH or DELETE requests for occupied-name
refusals, including singular and batch calls with overwrite set. Both paths return a known
non-commit and preserve the two item IDs. The case-only batch uses one revision-guarded
PATCH with the same ID and exact final spelling. A lost reply followed by reconciliation
at the wrong parent remains unknown even though the item ID and leaf match.

The earlier Debug plugin run `20261125T112201Z-3452-67d01c0fa25a433587e510261946c92c`
reported 252 passes and zero failures for Microsoft Drive before adding these seven checks.
Request counts are fixture assertions. They do not establish live OneDrive or SharePoint
qualification or a throughput claim.
