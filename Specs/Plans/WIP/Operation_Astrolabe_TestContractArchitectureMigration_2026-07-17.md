# Operation Astrolabe — Test Contract Architecture Migration

| Field | Value |
|---|---|
| Status | WIP — independently routed architecture-quality work; not a release blocker |
| Created | 2026-07-17 from Observatory Track 17 |
| Owner | Test infrastructure |
| Scope | Source-contract disposition, behavioral companion coverage, residual static-suite split, selftest compilation boundaries |
| Out of scope | Reopening closed Observatory product fixes or changing product behavior merely to satisfy source regexes |

## Why this is a separate operation

Observatory's runtime/data-safety work is complete. Migrating a large test-architecture surface is not a valid
tail task for an unrelated product fix, and the maintainer explicitly prefers convergence over repairing unrelated
tests. Astrolabe therefore owns the remaining architectural decision and can proceed without holding the completed
whole-repository remediation ledger in WIP.

The initial AST inventory is useful but intentionally not treated as a verdict. It currently classifies every
`It` block in `TestHarnessSourceContracts.Tests.ps1`; its `BehavioralShadow`/`MixedSourceShape` labels are inferred
only from positive/negative regex shape. That heuristic cannot determine whether a check is a lexical policy,
architecture/registration guard, companion-runtime wiring assertion, or an actual substitute for behavior.
The generated 133-candidate queue is therefore a review queue, not 133 pre-approved rewrites.

## Durable constraints

- Product behavior is proven in compiled/runtime tests with public contracts or injected seams.
- Source tests may enforce true lexical bans, dependency direction, registration/build graph, generated manifest
  parity, and the existence/wiring of a named behavioral companion.
- A source assertion must not claim that matching an implementation token proves runtime behavior.
- Do not replace a useful static ownership rule with an expensive UI/runtime test.
- Do not retain an exact implementation-body regex when an equivalent public behavioral test already fails for
  the regression.
- Intentional selftest `.cpp` aggregation may remain until compiled registration preserves debug-only linkage,
  case listing/filtering, shared fixture state, and binary-size/build-time evidence.

## Execution plan

1. Extend the source-contract inventory with an explicit reviewed disposition:
   `LexicalPolicy`, `GraphOrOwnership`, `CompanionWiring`, `ReplaceWithBehavior`, or `Retire`.
   Keep the review data next to the inventory logic and require one unique disposition per live case.
2. For `ReplaceWithBehavior`, name the owning executable/case and add the behavioral or fault-injection proof before
   deleting the regex. For `Retire`, record why no contract is lost.
3. Re-run the inventory after each subsystem batch; never publish a raw regex-shape count as implementation debt.
4. Split the residual static contracts by stable subsystem only after the shared parsing/setup helpers exist.
   The split must reduce change fan-out, not copy helper bodies into several Pester files.
5. Convert intentional selftest implementation aggregation to compiled registration only where the link/build
   graph stays debug-only and case listing, seeded order, repeat, filtering, crash-result preservation, and shared
   fixture ownership remain equivalent.
6. Update `Specs/Testing/Testing_TestCoverage.md`, `Tests/README.md`, and the live inventory output. Archive focused
   build/run evidence; broad Full reruns are required only when runner/run-plan behavior changes.

## Completion proof

- Every live source-contract case has one reviewed disposition and no heuristic-only replacement claim remains.
- Every removed behavior-shaped regex has a named green behavioral companion or an explicit retirement rationale.
- Residual static suites use shared parsers/setup and are split by stable ownership without duplicated helpers.
- Project/run-plan parity, runner case listing, repeat/shuffle, hostile-provider corpus, and test-result preservation
  remain green.
- Any selftest compilation-boundary migration records Debug executable/build-size deltas and preserves debug-only
  linkage.
