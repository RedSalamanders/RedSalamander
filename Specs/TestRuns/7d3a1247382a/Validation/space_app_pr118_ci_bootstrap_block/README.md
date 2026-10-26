# PR 118 hosted CI bootstrap block

- PR: [#118](https://github.com/DualTail/RedSalamander/pull/118)
- Run: [37882302098](https://github.com/DualTail/RedSalamander/actions/runs/37882302098), attempt 1; workflow `Suite PR on x64 Debug + ARM64 Debug (pull request)`.
- Head: `79bace17a450d540c9f3aed660b480cf71d9df7c`.
- Conclusion: failure before runner allocation. All four jobs (`migration-guards`, `format`, x64 Debug build, ARM64 Debug build) have zero steps, runner id 0, and no runner name. Check-run annotations say the jobs were not started because recent account payments failed or the spending limit needs to be increased.
- **No tests or builds executed.** This does not provide app/native test or compilation results. The required hosted x64 and ARM64 jobs remain blocked by the external GitHub billing/spending-limit condition.
- The run Artifacts API response reports `total_count: 0`; GitHub exposed no runner logs (`gh run view --log-failed` returned `log not found: 113664470178`). No artifact receipt or runner output exists for this attempt.
- `captured-manifest.txt` and `captured-README.txt` preserve the original local capture metadata. `manifest.txt` records byte length and SHA-256 for every copied capture file and confirms its archive copy. Local `.gitattributes` uses `* -text` to preserve JSON and text bytes; it is curator metadata and is excluded from the copy manifest.
- Captured run URL: `https://github.com/DualTail/RedSalamander/actions/runs/37882302098`.
