# July 21 to August 4 Review Residual Decisions and Evidence

> **ARCHIVED 2026-08-09 — SUPERSEDED AS AN EXECUTION QUEUE.** All seven real residual
> packages are transferred to
> `../WIP/Operation_Cinderstar_UnifiedRepositoryRiskRemediation_2026-08-09.md`; the optional
> UIA Invoke row is closed there against the authoritative passive-v1 contract. This file is
> historical evidence only.

Status: **Archived — live remainder transferred to Operation Cinderstar**
Created: **2026-08-04**
Source review: `../Done/CodeReview_July21ToAugust4_IndependentFindingsAndRemediation_2026-08-04.md`

## Purpose

The independent review is closed after its accurate, actionable safety and correctness
findings were repaired. This file preserves the bounded residual work as reviewed evidence;
Operation Cinderstar is now the sole live owner. Archived review prose is not a second queue.

### 2026-08-08 dependency correction

The later review of
`0f473e67bd547bc33a72cc9c3d7f39c99e8ca81a..68d2b3bc7b6f8f4e09163735124bd9d35c45890a`
found new executable correctness and build defects. They are owned by
`CodeReview_August8_CorrectnessDataSafetyAndSimplification_2026-08-08.md`, not by this
residual plan. In particular, `CR-RES-TERM-EVIDENCE` cannot close until A8-BUILD-01 through
A8-BUILD-04 restore a clean, checkout-invariant, locale-invariant Terminal build and validate
the actual cached DLL import closure. The WindowProc and executable-seam rows remain later
architecture/test-quality work and do not supersede the immediate A8-TERM behavior repairs.

## Residual matrix

| ID | Source | Priority | Remaining decision/work | Safe current state | Acceptance criteria |
|---|---|---:|---|---|---|
| CR-RES-TERM-KITTY | J21-TERM-03 | P1 | Move Kitty image conversion and upload preparation out of UI paint and out of `_terminalMutex`. | Existing rendering remains correct but can still stall paint under image-heavy output. | Publish immutable image work, prepare it on a bounded worker, consume only ready results in paint, prove cancellation/teardown, and archive same-machine image-storm latency evidence. |
| CR-RES-TERM-UIA-ACTIONS | J21-TERM-06 remainder | P2 | Decide whether the passive Terminal document should expose actionable prompt children through UIA Invoke. | Confirmations are themed, localized, readable, and announced through the existing text provider; no misleading action is exposed. | Either add stable child identities and executable Invoke actions with teardown tests, or record a product/accessibility decision that the terminal remains a passive text document. |
| CR-RES-TERM-EVIDENCE | J21-TERM-13, J21-TERM-18, EG-BUILD-01 proof | P2 | After A8-BUILD-01 through A8-BUILD-04 close, capture official Release x64/ARM64 Terminal closeout, cold/warm pinned-cache, output-storm, input-drain, and unload evidence at durable `Specs/TestRuns` paths. | Historical deterministic contracts and canonical Full run `20260804T234641Z-49108-eea3978802d6421580a1f73c78b22d8b` were green. At current HEAD, a clean Debug/x64 rebuild fails the uucode archive identity check; raw-byte frozen identities also depend on checkout EOL and tar preflight depends on locale. Historical green evidence must not be used as current release proof. | First consume the green A8 build/tooling fixes. Then run the authoritative scenarios with the same-machine protocol, retain bounded summaries plus raw evidence within archive policy, link every cited path from the Terminal and build specs, and prove official CI uploads `toolchain-identity.json`. |
| CR-RES-TERM-LOCALIZATION | J21-TERM-14 | P2 | Move remaining Terminal settings-schema labels/descriptions and user-facing diagnostics to resources. | Newly repaired confirmation text is localized; remaining English schema text is unchanged and visible only in existing surfaces. | Resource every user-visible string, preserve positional placeholders, update satellite resources and contract tests, and pass the localization audit. |
| CR-RES-TERM-EXECUTABLE-SEAMS | J21-TERM-15 | P2 | Replace remaining regex-only behavioral certification with executable seams. | Source contracts guard structure; native selftests cover the newly repaired critical behaviors. | Add executable negative/interleaving tests for each behavioral claim, retain source contracts only for static policy, and document the seam-to-contract mapping. |
| CR-RES-TERM-WINDOWPROC | J21-TERM-16 | P2 | Split the Terminal `WindowProc` into explicit input, accessibility, session, and rendering routes. | The current procedure remains functionally guarded and tested; no new messages should enlarge it. | Preserve message ordering/return semantics, keep payload-drain rules explicit, reduce the procedure to dispatch, and pass queued-payload teardown stress. |
| CR-RES-TERM-ABI | J21-TERM-17 | P2 | Decide whether unimplemented v1 record states/fields are reserved or removed before ABI freeze. | Every public entry validates record kind and rejects unsupported input fail closed. | Record the freeze decision, update headers/schema/spec together, add size/version/kind compatibility tests, and remove promises that have no supported behavior. |
| CR-RES-WINGET-ADOPTION | R4-CI-01 remainder | P2 | Adopt/resume an exact-version Winget PR instead of stopping on every duplicate. | Submission is serialized per version and fails closed when an exact open PR already exists. | Prove no-PR creates, matching-open-PR adopts and repairs, and conflicting/closed state stops safely; never expose the token in argv or generated source. |

## Execution rules

- Do not reopen the archived source review; update this matrix and the authoritative
  domain specification when a residual closes.
- Performance work follows `Specs/Testing/Testing_PerformanceValidation.md` from
  scenario definition through bounded archived evidence.
- Workflow recovery tests use stubs or test repositories and no production credential.
- A product decision may close a decision row only when its resulting behavior is
  durable in the authoritative specification and protected by a focused test.

## Definition of done

This plan moves to `../Done/` only when every row has landed or has an explicit
product-approved rejection recorded in the owning authoritative specification, all
linked focused proofs are green, and the canonical Full suite is green.
