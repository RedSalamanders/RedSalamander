# Curated run evidence

Keep new curated records under `Measurements/TestRuns/<Profile>/<Area>/<RunId>/`. Use a unique run ID
and record the source commit, profile, scenario, outcome and artifact digests required by the owning
validation contract. Preserve every case's status and each skip reason.

Promote compact evidence only. The limit is 2 MiB per file and 5 MiB per run. Retain larger raw
captures in the governed local or CI archive, and preserve the identity and claimed rows in the
curated record. Git ignores generated run directories and preserves the exact bytes of force-added
files.

Run `./Tools/Test-TestRunArchive.ps1 -RunPath <run-folder>` before staging. At closeout use
`./Tools/Test-TestRunArchive.ps1 -Inventory` to validate the governed history. The
[run-evidence contract](../../Specs/Testing/Testing_RunEvidence.md) owns publication and reuse rules.
[Historical archives](../../Specs/TestRuns/README.md) retain their recorded paths.
