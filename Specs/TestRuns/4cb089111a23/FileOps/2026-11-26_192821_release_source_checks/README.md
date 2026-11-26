# Release source check corrections

The non-interactive x64 Release PrePush on `3eed44771d3874e1f00d370d2e58adcbc07a38a7`
used build receipt `7529c57d9b9f7e6e60e9e19cbc29b9ef738701eaa300798cf1b2a68e0420640e`.
Run `20261126T173111Z-61920-ce6f482dfc3d40509c66ebd85a0141df` recorded 1,800 passes,
three failures and 29 environment skips. Thirteen of 14 entries passed. The change and repository
verdicts were both `FAILED` with filtered coverage. The three failures were in Tooling Pester.

| Failed check | Correction |
| --- | --- |
| Published scenario links | Retired the Monitor Preferences entry from the active gallery. Its historical image and provenance remain available. |
| Credential-bound Graph upload validation | Checked the mode-aware validation call and both typed upload-session parsing paths. |
| Preferences reverse-navigation focus helper | Matched whitespace after the call's opening parenthesis and kept all category and key requirements. |

The first focused run uncovered one new finding on the integrated main revision. The live remote
Rename case creates a private hidden owner, and the Win32 audit lacked that exact fixture. The
audit now applies the existing source, title, style and geometry rule to that owner. The fixture
table exercises both the visible and wrong-path variants. A trial assertion left a visible variant
in a shared fixture and caused a later assertion to fail. The final correction reuses the existing
private-owner fixture and isolates each variant.

| Focused run | Source | Documentation Drift | Test Harness Source |
| --- | --- | --- | --- |
| `20261126T182334Z-10452-2b34f95413da44c4b0c313ad8f0f4535` | `444dd03cd` | 30 passed, 1 failed | 244 passed |
| `20261126T182619Z-45068-76beb13cdcfa46388de805888055450d` | `279da7ebc` | 30 passed, 1 failed | 244 passed |
| `20261126T182821Z-72708-a7dc191ad6164f4699a24571c466d8e2` | `4aa1629a513f74ca091f107188f76a1d423d17aa` | 31 passed | 244 passed |

All focused runs had zero skips. Each used an isolated run under the marked sandbox and archived its
`artifacts/i66-release-source-checks.json`. The final result SHA-256 is
`60ff83134196091fa9c132fa35fbb1bff134096d9512800c88483d579ef373cc`.
The direct Win32 audit reported zero unclassified findings. The focused checks ran without a
native build or desktop activation.

The earlier broad archive retains both native results:
`artifacts/selftest/last_run/run-all-tests-results.json` with SHA-256
`30b01c8957b0b0d5252617b7d63966d39bf26aebcc3730e40bed6592c82181e1`
and `artifacts/selftest/last_run/run-all-tests-results-v1.json` with case-level failure reasons
and SHA-256 `f7b64eb443c034e04616fd2040a5b60c4211aed5da07ab7fb580237c980d4e7b`.
Both are under the archived run in `C:/RedSalamander.Perf/evidence/completed-runs`.

The four desktop entries remain pending. The governed rebuild and non-interactive PrePush
are running on the final focused source. Hosted x64 and ARM64 acceptance is tracked separately.
