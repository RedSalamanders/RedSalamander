# Handoff plan template

Resolve the plan directory/index from repository guidance before using this template. In RedSalamander it is indexed Specs/Plans/WIP; completed plans follow its reviewed Done-admission contract. Do not create another root queue.

A handoff must be executable without this conversation. Include only context needed for the selected outcome; link authoritative contracts rather than copying entire specifications or mutable dependency history.

## Plan contents

- Title, priority, owner, current status and explicit authorized scope.
- Planned-at commit/source identity and a drift command over the actual in-scope paths. Recheck dirty/untracked files as well as committed history; compare cited excerpts before proceeding.
- Concrete current problem, reachable trigger, impact and desired behavior.
- Owning files/symbols, normative contracts, existing implementation/test patterns and dependencies on other plans.
- Ordered implementation slices keeping code, meaningful scenario coverage and specs synchronized.
- Actual build/tooling/test commands and expected observable outcomes; distinguish static checks, executed tests and platform/performance/interactive qualification.
- Files and behaviors out of scope, rollback/recovery boundaries, and specific stop conditions.
- Done criteria, remaining gates, evidence locations and the repository's index/Done closeout steps.

## Execution index

Update the resolved existing index with one owner, dependency order and honest status. Keep implemented, reviewed, executed/pass, blocked and qualified distinct. Do not mark a plan Done until its actual acceptance criteria close; preserve failed receipts.

## Review checklist

Could a competent executor with zero conversation context act using this plan and the repository? Are references current, local paths valid, tests meaningful, stop conditions concrete, and claims backed by source or execution evidence? Do not require a test per function or performance measurement for unrelated prose. No secret values in excerpts or examples.

Git actions follow the repository/user's authorized workflow. A handoff does not authorize pushing, publishing issues, deleting worktrees or modifying another repository.
