# Follow-through: execute, reconcile, publish

Use the same repository-resolved plan directory/index as the audit. No literal root queue is assumed.

## Execute

An explicit human implementation request authorizes a transition from audit into implementation within its scope. Follow the applicable implementation workflow; read trusted repository guidance, recheck drift and existing ownership, and keep tests/specs/checklist synchronized.

If delegation is requested or otherwise authorized, give a bounded executor the plan, raw context, allowed files/actions and actual acceptance criteria. Use isolation where required by ownership or concurrent work. Inherit the configured model unless the user or applicable guidance specifies one; do not assume host-specific agent names, models or tool syntax.

Independently inspect the returned diff, verify scope and consequential assumptions, and rerun relevant checks. A successful metadata check or an executor's assertion is not proof of task behavior. Preserve failed evidence and request bounded revisions for defects. Never manufacture approval because a budget or iteration count is exhausted.

Integrate/commit/push only under the user's and repository's integration authorization. Explicitly report local versus remote completion.

## Reconcile

Verify completed outcomes against current artifacts/evidence, recheck drifted assumptions, route residuals to one owner, and retire rejected/superseded findings without deleting useful evidence. A stale or blocked plan is not automatically an implementation queue.

## Publish issues

Publish only when requested. Resolve the exact repository, visibility, issue contents and duplicate state first. Never copy secrets, private artifacts or unrelated code into an issue. If sensitive content would become public, present the concrete safe draft and obtain any required missing approval before publication. Record the actual issue URL only after successful creation.
