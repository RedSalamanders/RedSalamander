# Operation Rosetta Lantern — RedConfigure Satellite Localization Review

| Field | Value |
| --- | --- |
| Status | WIP — linguistic review owner |
| Created | 2026-07-17 |
| Origin | F4-LOC-01 from the last-four-days independent review |
| Scope | RedConfigure ordinary UI strings in Czech (`cs-CZ`), Japanese (`ja-JP`), and Slovak (`sk-SK`) |
| Priority | P2 localization quality; no runtime safety blocker |
| Primary constraint | Translation is not complete until a competent human reviewer for the target language approves it |

## Why this operation exists

The RedConfigure implementation and resource plumbing are complete, but the independent review proved that the
new ordinary UI strings were copied byte-for-byte from English into all three affected satellites. Resource parity,
placeholder checks, and successful compilation cannot establish linguistic correctness. This operation is the
single owner for that human-language work; code-remediation plans must route here instead of generating unreviewed
translations and calling the finding closed.

The initial reviewed set contains the 77 `IDS_REDCONFIGURE_*` strings identified by F4-LOC-01. Before translation,
refresh the inventory against current `RedConfigure.rc` so later resource additions are not missed.

## Required outcomes

1. Produce an exact source-versus-satellite inventory for `cs-CZ`, `ja-JP`, and `sk-SK`.
2. Classify source-identical values as either:
   - ordinary UI text requiring translation; or
   - reviewed language-neutral tokens, brands, commands, file extensions, or protocol names with a written rationale.
3. Translate every ordinary UI value and obtain target-language review. Preserve placeholder tokens, accelerator
   semantics, quoting, line breaks, and resource encoding exactly as required by `docs/dev/Localization.md`.
4. Add a non-blocking source-identical-string report with the reviewed neutral-token allowlist. The report informs
   review but must not globally reject legitimate identical brands/tokens.
5. Run resource placeholder/parity contracts, build RedConfigure and all three satellite resource sets, and smoke
   the Start, Localization, Themes, and Review & Export pages in each shipped culture. Check truncation, control
   overlap, accelerator collisions, and terminology consistency.
6. Update `docs/dev/Localization.md`, `Specs/Core/Core_RedConfigure.md`, `Specs/UI/UI_RedConfigure.md`, and
   `Specs/Testing/Testing_TestCoverage.md` with the lasting review/report contract and exact validation evidence.

## Review ledger

| Culture | Inventory refreshed | Translation complete | Native review | Resource/build proof | UI smoke |
| --- | --- | --- | --- | --- | --- |
| `cs-CZ` | [ ] | [ ] | [ ] | [ ] | [ ] |
| `ja-JP` | [ ] | [ ] | [ ] | [ ] | [ ] |
| `sk-SK` | [ ] | [ ] | [ ] | [ ] | [ ] |

## Definition of done

- every current ordinary RedConfigure UI string has an approved translation in all three satellites;
- every intentional source-identical value is present in the reviewed neutral-token allowlist with rationale;
- placeholder/resource contracts, the focused build, and the four-page culture smoke are green;
- exact reviewers and evidence paths are recorded without embedding private reviewer contact details; and
- this file moves to `Specs/Plans/Done/` after the authoritative localization, RedConfigure, UI, and testing specs
  contain the durable contract.
