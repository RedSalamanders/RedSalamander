# Specs/Reviews

> [!IMPORTANT]
> **Historical snapshot — not a live work queue.** Review evidence remains bounded by its original scope. Routing was reconciled on 2026-07-17 at repository commit `f4e0c8c3bed8`; the completed disposition ledger is `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`, while live work is indexed by `Specs/Plans/WIP/README.md`.

This folder contains date- and commit-bounded audit evidence. Its `*-Audit.md`, `*-Findings.md`, campaign summary, and
progress files are historical snapshots, not executable work queues. They remain here because later implementation
needs the original evidence, rejected alternatives, and source attribution even after the reviewed code changes.

Live work belongs under `Specs/Plans/WIP/` and is routed through `Specs/Plans/WIP/README.md`. When a review finding is
accepted for implementation, create or update one WIP owner, link the originating review evidence from that plan, and
persist implementation/test progress in the WIP owner. Do not move the review report itself into WIP: a report can
contain completed, refuted, deferred, and still-open findings at the same time, so treating the whole report as a live
plan would recreate stale or duplicate queues.

When the owner is complete, merge lasting behavior into the authoritative `Specs/<Domain>/` specification, move the
owner plan to `Specs/Plans/Done/`, and update the WIP index and any aggregate owner row. Review snapshots should change
only to correct their historical metadata or routing banner, not to act as a rolling implementation ledger.

Current routing:

- Completed July 21-to-August 4 independent review and disposition:
  `Specs/Plans/Done/CodeReview_July21ToAugust4_IndependentFindingsAndRemediation_2026-08-04.md`
- Completed unified owner for the audited August 8, residual, Emberglass, and Floodgate remainder:
  `Specs/Plans/Done/Operation_Cinderstar_UnifiedRepositoryRiskRemediation_2026-08-09.md`
- Whole-repository audit disposition: `Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md`
- Live plan index: `Specs/Plans/WIP/README.md`
- Retired Squad framework decision log (archived 2026-08-26): `Specs/Reviews/Squad-Decisions-Archive.md`
- Curl short-success staged-upload historical proof: `Specs/Plans/Done/Operation_Curl_Short2xxStagedUploadProof_2026-07-21.md`
- Archived Curl/cloud follow-up evidence: `Specs/Plans/Done/Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md`
