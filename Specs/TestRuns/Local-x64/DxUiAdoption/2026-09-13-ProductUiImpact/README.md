# Product UI affected-test mapping regression

On base `db35652f`, the new live-project ownership test fails against the retired
DxUi project mappings: **8 pass, 1 fail**. After replacing the obsolete rules with
the live `ProductUiTests.vcxproj` owner and existing standalone tag, **9 pass,
0 fail, 0 skip**. The exact test source hash is unchanged across runs.

The first corrected-manifest attempt fails canonical LF validation before running
cases. That failed attempt is retained. Correcting line endings produces the
passing result; no coverage assertion is removed. Both exact dependency-pin and
shared-adapter changes still select product UI tests and conservative Full
coverage. Product test source changes select the standalone test surface.

Tool inventory and normative spec inventory pass, with zero blocking spec findings.
The archived scripts and manifest snapshots are inert `.txt` files. Original logs,
source identities and byte hashes are in `retained-files.json`. This is focused
local tooling evidence, not final native product qualification; no hosted CI ran.
