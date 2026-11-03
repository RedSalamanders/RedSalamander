# R0e paired containment gate — x64 Release

Five independent focused processes ran for each of main 269b06fc, PR200 32d20c9b and PR201 8dadef4f. Each native result is retained unmodified and reports Setup, selected case and cleanup passed (3/0/0), with focused runner exit 2. The identical reviewed case-body SHA and each source snapshot, receipt, run ID, quarantine and original raw-file hash are recorded in provenance.json.

The five measured samples per side retain all eight R0e metric rows per process, byte-exact and in source line order. Every process recorded 4096 parses, 4096 decisions, one uncontained rejection and zero provider calls. Baseline p95 was 817124 us; PR200 was 822324 us and PR201 819371 us. Both passed the unchanged 1.5x baseline limit of 1225686 us. The two assessment JSON files are the original derived paired assessments. This is a bounded in-process capability/admission containment comparison, not network or general throughput or Full qualification.

An earlier quiet1 launcher preflight failed before any native case because a relative receipt path resolved to the wrong checkout. Its separate preserved failure manifest is hash-bound in provenance; it is not a native test failure. The corrected launcher and worker hashes are also recorded. Full raw streams and original wrapper records remain in the approved local evidence store.

The negotiated host/no-heavy-work quiet window was 04:33–05:00 UTC (including the extension). The peer declared its prior native work drained at 04:29:39 UTC; the first native run ID records 04:40:42 UTC, and all native processes, receipt checks and assessments were observed drained by 04:58:08 UTC. These bounds do not assert a broader desktop or no-activation guarantee.

Each of the 15 separate fresh processes used the canonical `Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_ProviderCapabilityMatrix -Configuration Release -Platform x64 -BuildNumber 1 -SkipBuild` invocation with tests enabled and Embedded compiler information. The source, receipt and individual run identities are recorded separately in provenance.json.
