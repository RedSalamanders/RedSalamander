# Consumer compiler-host propagation

The final ARM64 Debug attempt at `cc7595de` failed with 13 compiler-host
mismatch rejections. First-party projects selected x64 through the new
Directory.Build.props policy, but standalone DxUi identity evaluation selected
x86. The original failed build and resolved identity are retained here.

The correction evaluates RedSalamander/RedSalamander.vcxproj for the selected configuration
before fingerprinting DxUi, scopes the selected host to the canonical identity
probe and restores the caller environment even when the probe throws. The
consumer's DxUi project reference explicitly forwards PreferredToolArchitecture;
MSBuild does not inherit locally evaluated properties across that boundary.
The existing mismatch guard stays enabled. The DxUi pin and source are unchanged.

Both new regression checks fail on the old implementation: restore records x86
instead of x64, and the library project reference omits the selected host.
With the correction, all eight dependency tests and 29 inventory/governance tests
pass. Coverage includes real six-profile MSBuild evaluation, explicit overrides,
profile-dependent restore policy and restoration after an injected probe failure.
No compilation or hosted CI is invoked by these evaluation fixtures.

The first correction at `57ff0f8d` used an incorrect probe-project path and
stopped before compilation. The corrected private bootstrap probe uses the same
real project as BuildEvidence. The strengthened tests invoke that production
probe for all six profiles and explicit overrides; all eight tests pass. The
failed attempt and the new test receipts are also retained.

The full ARM64 build retry is still required. Native ARM64 execution remains a
separate gate. retained-files.json preserves every copied original byte.

The original resolved-property snapshot is retained with a `.props.txt` suffix. Its bytes and
recorded source path are unchanged. A bare `.props` suffix made the Full impact validator
correctly demand an active build-graph import; this file is diagnostic evidence, not a build input.
The failed preflight is retained in [the following attempt](../ad7f7735-preflight/README.md).
