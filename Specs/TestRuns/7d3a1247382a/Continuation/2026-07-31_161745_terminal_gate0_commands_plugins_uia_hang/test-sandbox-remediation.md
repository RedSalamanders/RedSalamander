# TestSandbox remediation record

The canonical full-suite aggregate remains byte-for-byte unchanged and records
`test_sandbox_audit.is_clean=false`, `issue_count=9`. Its nine issues name
unexpected children directly under
`D:\RedSalamander-terminal-gate0\.build\TestSandbox`:

1. `host-header-ancestor-reparse-4894d3f579cd45cabbe015442a817e51`
2. `host-header-tamper-99576fd73ed24083b2249516fcfc52dc`
3. `neutrality-direct-5225752b95004ac8b49a494900246d9a`
4. `neutrality-one-candidate-71d49f3db57f4d459cfc35ed0e2282ec`
5. `neutrality-one-candidate-8a4a0747d1214278a74aabda6a77ef33`
6. `round3-transition-direct`
7. `universe-v3-independent-10c424cdf2ad45ddb78fff4c2f9bee78`
8. `universe-v3-independent-530c5c8709254d269f14e7a5f21ac104`
9. `zip-size-forge-b5e789b82a434fde9bdf338fae2aa34e`

The directories' creation and last-write UTC instants range from
`2026-07-30T22:18:12Z` through `2026-07-30T23:01:46Z`. They predate the final
focused runs and the full-suite wrapper start at `2026-07-31T15:38:29Z`.
Neither later run updated them. This exact command returned no match against
the frozen source:

```powershell
git grep -n -E 'host-header-(ancestor-reparse|tamper)|neutrality-(direct|one-candidate)|round3-transition-direct|universe-v3-independent|zip-size-forge' 3e84e87b717841e28e11e1fe7e5bc04b4e3a3e0d -- .
```

The operator verified each absolute source path stayed below the intended
`TestSandbox` root and moved exactly the nine directories, without deleting
them, to the ignored recoverable quarantine:

```text
.build/terminal-engine-spike/round4-governance/test-sandbox-residue-quarantine/
```

An immediate disk audit then returned `is_clean=true`, `issue_count=0` while
allowing the canonical full-suite run. The accepted reproduction used the same
20-file list and exact 239-test assertion as the checked-in Round-4 closeout
workflow. Its final record is:

```text
StartedUtc=2026-07-31T17:30:43.9446567Z
CompletedUtc=2026-07-31T17:41:25.8137614Z
RunId=r4clean-284e0752
Files=20
Passed=239
Failed=0
Skipped=0
Pending=0
Inconclusive=0
PesterSeconds=639.63
SandboxClean=true
SandboxIssueCount=0
```

The post-Pester disk audit ran in the same PowerShell process and allowed only
the immutable full-suite run
`20260731T153831Z-35744-9dceccf8fcac42a8b91d71f30c3a489e` and the accepted
remediation run. The root contained only the `runs` directory. The accepted
logs are:

| File | SHA-256 |
|---|---|
| `sandbox-remediation-focused.stdout.txt` | `b253a566b3bcfee8a3f2878c1e6d06e01a077ae57fb19a2fc87e6cb5c4d6873c` |
| `sandbox-remediation-focused.stderr.txt` | `d2d2470773f41a8075b387bd73d02148cbf72f0c2fce554e85321cee8aa10250` |

The stderr file contains only Git line-ending warnings. The successful rerun
created no unexpected TestSandbox-root child. This closes the historical
residue finding without rewriting or misrepresenting the original aggregate.
