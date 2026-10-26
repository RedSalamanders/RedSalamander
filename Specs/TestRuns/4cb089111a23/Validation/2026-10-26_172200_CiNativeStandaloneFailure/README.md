# Hosted native Suite PR failure

Public head `d36cb3b7f66c8bb61ae05bf788ed70b97ada213d` passed formatting,
migration guards, CodeQL and both native Debug solution builds. Its
[Suite PR run](https://github.com/RedSalamanders/RedSalamander/actions/runs/40358566877)
failed standalone validation on both architectures. Complete job logs, selected
failed-entry receipts/results and baseline records are published here. Both full
uploaded artifacts remain in the approved local evidence directory, byte-verified
against manifest.json's files/externalOnly maps. Oversized archive intents and
redundant transaction/history projections are retained there to respect Git archive
size limits. No source evidence was deleted from the original downloads.
These failures are not promoted by subsequent passing isolated checks.

x64 failed the S3 directory-transfer probe group, four tooling assertions and
ViewerPE noninteractive checks. ARM64 passed plugin contracts but failed five
tooling assertions and ViewerPE checks. The tooling failures include two omitted
search discovery names, nested schema-error diagnostics, and an ARM64 fixture
that assumed MSBuildTaskHost.exe lived beside the native MSBuild executable.
The nested-error omission matches the
[upstream PowerShell diagnostic regression](https://github.com/PowerShell/PowerShell/issues/28218).
Invalid locks were still rejected; their field diagnostics were missing.

On unchanged qualified native source `582d644be5e6147acb0b8a8aae05511b75f64d17`,
the isolated x64 Debug browser-profile, lifecycle-capture, directory-summary and
S3 directory-transfer checks passed. The 200k-file/20k-directory case passed in
about 48 seconds and also passed in the retained single-CPU diagnostic control.
Those ordinary-host results do not establish a passing hosted environment.
The native source was unchanged throughout each recorded diagnostic run.

The candidate repairs discovery and tooling contracts, uses an owned WARP WGC
capture/readback device, fits lifecycle fixtures to their actual work area,
and waits for the held provider's UI-published scan state within the original
eight-second deadline. Existing data/resource/timing predicates remain intact.
Browser-profile values, capture HRESULTs, S3 timing pairs and periodic large-tree
progress now make hosted failures observable. Current-source native validation
and public CI remain required. Full runtime reruns stay omitted under the user's
accepted scope; installed-helper, package, release and manual gates remain separate.
