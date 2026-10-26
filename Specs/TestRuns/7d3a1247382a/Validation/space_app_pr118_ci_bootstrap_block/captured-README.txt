PR 118 CI bootstrap failure record

- Run: 37882302098, run attempt 1, PR 118.
- Head: `79bace17a450d540c9f3aed660b480cf71d9df7c`; branch `codex/viewer-space-explorer-analyst-implementation`.
- Workflow: `Suite PR on x64 Debug + ARM64 Debug (pull request)`; run 551; event `pull_request`.
- Run conclusion: `failure`, created 2026-10-09T04:06:02Z and completed 2026-10-09T04:06:05Z.
- All four jobs (`migration-guards`, `format`, x64 Debug build, ARM64 Debug build) failed before a runner started. Each API job record has `steps: []`, `runner_id: 0`, and an empty runner name. All four check-run annotations state: “The job was not started because recent account payments have failed or your spending limit needs to be increased. Please check the 'Billing & plans' section in your settings”. This is an account/workflow-start failure, not a source build or test failure; no app/native test cases ran.
- The artifacts endpoint returned `total_count: 0`. `gh run view --log-failed` returned `log not found: 113664470178`; there are no runner logs or artifact receipts to preserve for this attempt.
- JSON files are snapshots of GitHub API response bodies obtained with `gh api`; `logs-unavailable.txt` records the CLI's unavailable-log response. `manifest.txt` contains SHA-256 and byte lengths for the captured files and this README.
- Evidence is under the initialized `D:\RedSalamander.Perf` root. No CI rerun was initiated.