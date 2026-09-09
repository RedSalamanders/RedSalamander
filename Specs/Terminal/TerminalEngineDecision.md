<!-- current-executive-summary:start -->
# Terminal Engine Decision — Current Executive Summary (2026-08-04)

## Decision

**Ghostty `lib-vt` is the selected engine for the first embedded-terminal
implementation.** RedSalamander manufactures the winner by changing two
historical evaluation requirements transparently:

1. One bounded, separately shipped `ghostty-vt.dll` is allowed only inside
   `Plugins\TerminalRuntime\`, privately loaded and owned by
   `Plugins\Terminal.dll`.
2. Cross-build raw-byte identity is replaced by frozen-input verification plus
   a generated exact plugin/runtime pair. Each build may have a different
   timestamp, CodeView identity, or code layout, but its `Terminal.dll` accepts
   only the exact `ghostty-vt.dll` size and raw SHA-256 generated before that
   plugin was compiled.

Ghostty is not the RedSalamander plugin boundary; `ITerminal` is. No Ghostty
type, header, symbol, callback, or handle crosses into the EXE or the public
ABI.

The earlier zero-runtime-leaf and byte-identical-rebuild Round-4 decision and
its versioned requirement JSON remain preserved below as historical evidence.
They correctly found no winner under that older policy. The current product
contract in `Terminal_EmbeddedPlugin.md` supersedes those two requirements for
the shipped embedded terminal instead of pretending that Ghostty already
qualified under them. Historical references below retain the WIP paths that
were true when those records were sealed; the completed Terminal plans are now
archived under `Specs/Plans/Done/` and are not current authority.

## Why this is the best option

Ghostty is the strongest engine fit and the private adapter is the smallest
honest architecture. The engine already supplies the VT model and Kitty-aware
state; RedSalamander supplies ConPTY, DirectWrite/Direct2D presentation,
Windows input, pane lifecycle, settings, and packaging. A closed dynamic
function table contains Ghostty's evolving C API, permits terminal-only
degradation when the runtime is absent, and makes upgrade/rollback an atomic
plugin-runtime operation. WezTerm still lacks a comparably small public
embeddable C-library surface, while building an in-house VT engine would make
RedSalamander the long-term owner of conformance and security.

## Decision table

| Option | Boundary | Windows/product work | Upgrade and rollback | Long-term burden | Decision |
|---|---:|---:|---:|---:|---|
| Private `ghostty-vt.dll` behind `Terminal.dll` | 5/5 | 3/5 | 5/5 | 4/5 | **Selected** |
| Static Ghostty inside `Terminal.dll` | 5/5 | 2/5 | 3/5 | 3/5 | Fallback only |
| RedSalamander Ghostty fork/wrapper | 5/5 | 2/5 | 3/5 | 1/5 | Reject unless a small upstreamable patch is unavoidable |
| WezTerm caller-owned adaptation | 4/5 | 1/5 | 2/5 | 2/5 | Retain as historical contingency, not active |
| In-house VT/Kitty engine | 5/5 | 1/5 | 4/5 | 1/5 | Reject for v1 |

## Qualification result

The exact engine pin is
`b2d44625908eb56aee299d8511d803bc11ab79fc`, built with official Zig 0.16.0,
`ReleaseFast`, SIMD disabled, stripped output, LLVM/LLD, and only the shared
`lib-vt` target.

| Architecture | Artifact identity | Machine/import result | Result |
|---|---|---|---|
| x64 | Exact size/raw SHA generated into the paired `Terminal.dll` and schema 5 identity | AMD64; system imports only; required exports present | Pass |
| ARM64 | Exact size/raw SHA generated into the paired `Terminal.dll` and schema 5 identity | ARM64; system imports only; required exports present | Pass |

Zig/LLD can legitimately emit a different raw hash or code layout from
identical clean inputs. The production trust boundary is therefore
`generated-raw-sha256-v1`: each build generates an exact size/raw-hash header
before compiling `Terminal.dll`, so the shipped plugin and private runtime form
one byte-locked pair. A normalized PE hash is retained only for diagnostics and
never authorizes loading.

The tracked build-only Windows overlay
`Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch`
(`bc9ac4cbae85d8a491e32471096e43f2042c03f4558728bfb067ddb8c36cf2ac`)
omits the unused static companion, selects LLVM/LLD, and propagates stripping
to the VT module. It applies equally to x64 and ARM64 and does not change the
engine source or public API. It supersedes the former ARM64-only overlay.

## Current implementation gate

The v1 x64 implementation is complete and covered: lean typed unsuffixed
`ITerminal`, generated exact size/machine/raw-SHA paired private loader, ConPTY, structured
styled Ghostty rendering and Kitty graphics, bounded asynchronous input,
paste/selection/mouse/IME, passive UIA Text, authenticated PowerShell cwd and
prompt state, trusted idle follow, opposite-pane reuse, Windows/WSL launch,
folder Edit, contextual/full path insertion, Preferences, centralized atomic
staging with complete notices, and nonblocking teardown. The x64 and ARM64
engine products are offline-verifiable. Native ARM64 whole-product compilation
remains a release-environment qualification lane, not an engine-selection or
v1 implementation blocker. Final closeout on 2026-08-03 passed every
Terminal-focused build, source, native, lifecycle, dependency, packaging, and
offline-reproduction gate. The 2026-08-04 Release-builder correction then
passed complete x64 Release and Debug solution builds, all focused runtime
identity/dependency/source contracts, and `PluginContractTests.exe` including
all 23 Terminal runtime selftests. The broader interactive run's non-Terminal
failures remain recorded honestly, as detailed in `Terminal_EmbeddedPlugin.md`.
Those results certify their tested bytes only. The 2026-08-09 checkout/locale/cache
remediation pins the complete raw-byte governance surface to LF, makes tar metadata
inspection locale-independent, and restores the authoritative
`TerminalEngine*.Tests.ps1` aggregate to 212 passed, zero failed, and five intentional
skips under the ordinary Windows
checkout. It also proves independently locked uucode source/canonical archives and
actual PE imports for x64 and ARM64. Product Release evidence remains open to the
separate Terminal release-proof package; historical evidence is not promoted to the
current product bytes.

<!-- current-executive-summary:end -->

<!-- superseded-2026-07-31-executive-summary:start -->
# Terminal Engine Decision — Executive Summary

## Decision requested

**Approve Route A now: run the capped synchronous, caller-owned WezTerm proof.**
Do not select a production engine and do not start product implementation yet.
The proof is evidence-only, has a hard limit of 15 cumulative owner-days, and
keeps every product, security, plugin-only, native-lane, maintenance, and score
requirement unchanged.

No current candidate qualifies. The best next move is therefore not to pick the
least-bad engine or quietly loosen a cap; it is to buy one small, bounded piece
of information that can create a conforming candidate without moving the
requirements. If the proof succeeds, WezTerm must still complete a fresh
nomination, pass G1–G12 and all native lanes, and score at least 80.00 before it
can become the winner.

## Executive view

- **Best option:** manufacture a synchronous, caller-owned WezTerm boundary.
  Its current adaptation already fits the line and file caps. The hypothesis is
  that removing the asynchronous callback/`Arc` lifecycle removes the sixth
  concern and permits a materially smaller adapter/upgrade surface, reducing the
  mechanically calculated three-year effort from 87 to about 59 owner-days.
  That would fit the unchanged five-concern and
  60-day caps, but with only one reported day of headroom; the proof must measure
  this rather than assume it.
- **Why this is the best decision:** it preserves the intended architecture and
  user-visible requirements, has a strict 15-day downside, attacks the one
  identified design concern responsible for much of WezTerm's excess lifecycle
  surface, and retains the option to stop without contaminating product code.
- **Why not Ghostty:** the incomplete Ghostty v4 adaptation is already 2,213
  changed lines across 21 files and 84 owner-days before its missing Gate-0
  adapter. Completing the missing work can only increase that footprint. There
  is no evidence-based finite cap change that makes Ghostty selectable today.
- **Why not choose a fallback immediately:** Contour requires larger line,
  concern, and owner caps and still has unresolved Kitty accounting. TerminalCore,
  libtsm, and libvterm fail mandatory Kitty G7 at source review. Relaxing those
  requirements before the bounded WezTerm experiment would spend product or
  maintenance quality without first testing the least-invasive path.

## Decision table

| Route | What changes | Evidence and principal risk | Decision |
|---|---|---|---|
| **A — synchronous caller-owned WezTerm proof** | No requirement changes. Spend at most 15 cumulative owner-days testing a no-callback, no-post-return-work boundary. | Current 2,459-line/seven-file shape fits those caps; success still requires reducing six concerns to five and proving a mechanically recomputed owner ledger at or below 60 days. The projected 59 days has almost no slack. | **Approve now. Recommended.** Best expected information value with bounded cost and no product compromise. |
| **B — increase maintenance headroom** | Raise owner cap from 60 to 90 days and concern cap from five to six; keep all gates, native lanes, runtime, and score requirements. | Could admit the existing WezTerm shape for fresh evaluation, but accepts materially more three-year maintenance and still proves no winner. | **Hold as the first explicit fallback.** Consider only if Route A fails solely on honest concern/owner accounting and the business accepts that cost. |
| **C — enlarge the C++ adaptation budget** | Raise line cap to 3,200, concern cap to six, and owner cap to 80 days. | Could re-admit Contour only after its Kitty image-versus-text accounting is repaired and re-proved. | **Not preferred.** Higher maintained patch surface than A, with an unresolved correctness issue. |
| **D — defer Kitty graphics** | Remove mandatory G7 Kitty APC/static-image support from v1. | Could reopen TerminalCore, libtsm, and libvterm, but cuts a visible product capability and proves none of their other fitness requirements. | **Product decision only.** Use only if Kitty images are intentionally removed from the release. |
| **E — relax plugin/native ownership** | Permit host-owned terminal code, a web stack, or a new in-house parser. | Conflicts with the plugin-DLL-only/native architecture and expands security and maintenance scope. | **Reject under the stated requirements.** |
| **F — raise a Ghostty cap** | No defensible finite change is known because the adapter is incomplete and already exceeds the owner cap. | Final ABI, teardown, accounting, line/file, and concern footprints remain unknown; 90 owner-days is not evidence of fit. | **Reject today.** Reconsider only after a complete measurable adapter or suitable upstream API exists. |

## Stop and escalation rule

1. Stop Route A at 15 cumulative owner-days or earlier when the same boundary
   cannot satisfy ARM64, G7/G9 pre-allocation bounds, one-toolchain/no-runtime-
   leaf policy, five concerns, 2,500 lines, 24 files, or the 60-day ledger.
2. Treat a successful proof only as permission to prepare the complete v5
   nomination. Product implementation remains blocked until the full Gate-0
   decision publishes a winner.
3. If Route A fails only because the honest final concern or owner ledger exceeds
   the existing cap, bring Route B back as a visible product/maintenance choice.
   Do not grant a candidate-specific waiver.
4. If it fails an ownership, safety, native, resource, ABI, or mandatory feature
   gate, keep the requirements and open a new candidate universe or wait for an
   upstream API; raising maintenance caps would not cure that failure.

**My recommendation:** approve Route A and its 15-day hard stop. It is not a bet
that WezTerm is already the winner; it is the lowest-cost experiment with a
credible path to a winner while preserving the product. I would not approve
Ghostty, Route B, or a feature/architecture relaxation before seeing this proof.

The executable plan is
`Specs/Plans/WIP/Terminal_EngineSelectionAndDependencyProof_Round5_WezTermSynchronousCallerOwned_2026-07-31.md`.

### Historical-integrity note

The original Round-4 decision begins at the first byte of the exact
`# Terminal Engine Gate 0 Decision` line immediately following the
`round4-protected-decision:start` marker below. From that byte through the byte
immediately before `### Decision closeout audits`, its block-local length remains
50,569 bytes and its SHA-256 remains
`c9cf3f18ebe2fb53736c9589bdcbad5d441d1e5cfcd98d8ab07ba7d2e362e3b3`.
Historical statements inside that preserved block that refer to byte zero or
the prefix are interpreted relative to this protected block. The authoritative
Round-4 closeout blob at commit `6674119fd24284877b1b25a3a71dbe2f6e6ee9a2`
is unchanged.

<!-- superseded-2026-07-31-executive-summary:end -->
<!-- round4-protected-decision:start -->
# Terminal Engine Gate 0 Decision

## Decision

**Status: no current candidate qualifies; product integration remains blocked.**

Gate 0 did not select Ghostty, Contour, WezTerm, Microsoft TerminalCore,
libvterm, or libtsm. This is a completed comparison outcome, not permission to
choose the least-bad candidate. The frozen adaptation caps, mandatory gates,
score anchors, threshold, source decisions, and fixtures were not changed after
candidate measurements began.

The `builtin/terminal` user-interaction and plugin-only product contracts remain
the required design if a later engine wins. No terminal engine lock, shipped
engine artifact, or `Plugins/Terminal/Terminal.dll` implementation is authorized
by this decision.

## Frozen comparison

The effective Round 4 universe keeps the Round 1 source pins, the Round 2
protocol/fixture corrections, and the Round 3 nominations, candidate-neutral
provider, and exact cross-repository ledger. Round 4 changes no candidate,
source pin, cap, score, threshold, control, fixture, or measurement recipe; it
closes governance chronology and immutable provider consumption after the
earlier downstream records were found to predate the approved Round 3 universe.
Every candidate-specific source, adapter, build, driver, test, generated,
documentation, legal, dependency, and packaging path counts regardless of which
repository contains it.

The conjunctive adaptation caps are:

| Measure | Cap |
|---|---:|
| Three-year owner effort | 60 days |
| Changed maintained lines | 2,500 additions plus deletions |
| Changed maintained files | 24 |
| Independently regressible logical concerns | 5 |
| New compiler/language toolchain families | 1 |
| New non-system runtime leaves | 0 |
| Binary patch | forbidden |

The exact no-renames patch/result-tree proof and two independent non-author
approvals are required before a nomination can enter native collection.
Mandatory admission failure stops that nomination before performance scoring.

### Immutable governance-consumption boundary

Round 4 also freezes how governance input is consumed. The trusted evaluator
opens one no-write/no-delete read snapshot before interpreting candidate
governance. It holds the repository root, every traversed directory, and the
exact 83-path static closure through the complete read. That closure is the
union of the 65 Round-4 frozen-input leaves, the 51 collection-neutrality
surface leaves, the four live governance records, the historical universe
readers, and their closed schemas, with duplicates removed. Every path
component is opened without following a reparse point and is checked by handle;
a missing, substituted, reparsed, or unopenable component rejects the closure.
Schema validation, canonicalization, SHA-256 identity, and policy comparison
read only paths whose exact file and ancestor identities remain protected by
those held handles. Some validators reopen a protected pathname, but the
no-write/no-delete handles and reparse checks prevent it from resolving to
different bytes during the closure; no unprotected later read can become
authority.

The evaluator first validates that the universe declares the exact sorted
65-leaf frozen-input set, then verifies only those known held paths. It never
reads a universe-supplied relative path before that equality check. A bounded
pre-schema read of the held CandidateSet and CandidateAssessment may add only
the canonical patch and concern-map paths referenced by an
`adaptation-rejected` outcome or a `collect` ledger; those dynamic leaves and
their ancestors are added to the same held snapshot before semantic
validation.

Frozen helper code cannot execute merely because it was found by path. After
the universe's canonical bytes, closed schema, proposal digest, two independent
approvals, base identity, governance-schema identities, and frozen-input hashes
have validated from the held bytes, the evaluator force-imports the exact
authenticated TerminalEvidence, candidate-assessment, and
collection-neutrality modules and verifies each loaded module path.
Collection neutrality is invoked only inside that authenticated snapshot
boundary (or a newly acquired equivalent boundary), so provider code cannot
swap the scanner or a neutrality-only surface leaf between governance
validation and scanning. When the provider reacquires that equivalent boundary,
the evaluator must compare the exact SHA-256 identities of the universe,
CandidateReview, CandidateAssessment, and CandidateSet to the provider context,
in addition to the neutrality-policy identity, before scanning. A newly valid
but different governance closure is not equivalent merely because it carries
the same policy. Any future candidate-changing universe must refreeze this
complete closure; Round-4 authentication does not authorize later executable
candidate consumption from released path handles.

Round 4 has no executable cohort and does not authorize the optional
path-reopened source-archive verification path as a provider/collection
authority. The source blockers remain bound to their previously reproduced
archive, member, line-range, and digest proofs, but no archive is claimed as
part of the 83-path snapshot. Before a v5-or-later executable cohort can enter
Bootstrap, it must add the bounded review-referenced archives and exact
SourceReview module to one held authenticated snapshot and keep those handles
live through both archive hashing and member extraction. Hashing a released
pathname and then reopening it is not sufficient.

The Round-4 provider performs its authenticated neutrality proof before any
descriptor is loaded. That is sufficient only because the closed Round-4
candidate set contains no executable candidate, so descriptor-driven
collection is unreachable and unauthorized. Before a v5-or-later executable
cohort may run, the provider must repeat the proof immediately after validating
the exact candidate descriptor and before any descriptor-driven bootstrap,
host/build, subprocess, or evidence-root effect, bound to the same four
governance identities. Descriptor-only dependency or artifact literals must be
included in that second scan.

This boundary proves run-time byte consistency and reproducibility, not
cryptographic authorship. The already loaded evaluator and the reviewed clean
Git commit are the trust roots; as the evidence contract states, no signature
is implied. A same-user attacker who rewrites the evaluator and repository
before launch is outside this boundary. Bringing that threat into scope would
require an external anchor such as a verified signed commit/manifest or a
code-signed verifier, not another repository-internal hash.

## Candidate outcomes

| Candidate | Exact source pin | Measured outcome | Decision |
|---|---|---|---|
| Contour-RS v3 | `0397f9cfbfcc94d46103d3cff70ebf59326dd695` | 3,041 changed lines in 15 paths, six concerns, and 76 owner-days before any collection descriptor; a separate non-schema audit also observed that the adapter charges retained Kitty image bytes as terminal-text CPU memory rather than image CPU memory | Rejected: changed-lines, logical-concerns, and owner-days caps; the accounting observation is non-dispositive and is not represented as an additional approved failure code |
| Ghostty-RS v4 | `70c498ac3273661aebf6cce9904c0d42b2e5d299` | Its conservative partial lower-bound patch already changes 2,213 lines in 21 files across five concerns and adds one Zig toolchain family; the approved upper ledger is 84 owner-days | Rejected before completion: owner-days cap |
| WezTerm-RS v3 | `46a166d6dc8188ede1a32a96375e308081ebafc5` | 2,459 changed lines, seven files, six concerns, one new Rust toolchain family, 87 owner-days | Rejected: logical-concerns and owner-days caps |
| libtsm | `8a40f0b8eca0ec4584a28f3e6371dca8dca4d7ba` | Exact parser state has DCS/OSC but no APC state | Source-rejected at G7: required Kitty APC subsystem is absent |
| libvterm | `bazaar-r844:leonerd@leonerd.org.uk-20240724211608-as82xl1xmv3st5bc;mirror-dfc4c5e5b3dd99247dc95031a8f40087f181dea5` | APC is only forwarded to an optional generic fallback and otherwise unhandled | Source-rejected at G7: required Kitty APC subsystem is absent |
| Microsoft TerminalCore | `743f23c2cf6f430678f8ff3c6ca0d07c7b233646` | The parser treats SOS/PM/APC strings as ignored | Source-rejected at G7: required Kitty APC subsystem is absent |

Microsoft/full-control, WebView2/xterm.js, and a new RedSalamander parser remain
unscored constraint controls. They respectively violate candidate-owned UI
ownership, the no-second-web-stack boundary, and the no-new-parser maintenance
boundary.

No candidate reached native scored collection, so there is no score cohort,
tie-break, winner artifact, ABI fingerprint, or terminal engine lock.

## Why Ghostty was not accepted

Ghostty was challenged as an incumbent hypothesis rather than assumed to be the
answer. Its pinned engine has the strongest existing Kitty subsystem of the
small candidates and its Windows x64 proof build succeeded. The measured v4
patch is deliberately treated only as a conservative lower bound: it exposes a
Ghostty VT C API but contains no frozen Gate-0 adapter query, session/action, or
teardown export, so it is not an executable Gate-0 adaptation. Even before
adding and counting that missing private adapter, the 21-path patch plus the
approved upper ledger derives 84 owner-days; using every applicable policy
floor still derives 81. Both exceed the 60-day cap. Missing adapter work can
only increase the ledger and cannot rescue the nomination.

The live upstream head checked during this decision was
`4d605bf0d819df901a0332bbb320dc849fdd82e4`, two first-parent commits after the
frozen pin. Its relevant change is preparation for a future binary snapshot API,
not the complete consumer-owned snapshot/resource boundary required here. The
same adaptation still applied at that head with the same cap-breaking footprint.
Future completed upstream snapshot and accounting APIs may justify a fresh
universe and nomination; prospective work cannot pass this one.

## Product and dependency consequences

- Do not start the terminal product phases, add an engine runtime to
  `RuntimeDependencies.props`, or create `External/terminal-engine.lock.json`.
- Do not implement the terminal through `IViewer`, the EXE, a reparented console
  window, a web terminal, or a new in-house VT parser.
- Preserve the planned `builtin/terminal` plugin-DLL boundary and the
  user-interaction/settings contracts so they can be applied unchanged after a
  valid engine is selected.
- The dependency lifecycle tools and schemas remain a reusable pre-winner
  scaffold for a future round. They are tested with synthetic locks but have no
  current production lock to observe or upgrade, and their discovery metadata
  is not yet a production update oracle. Inventing a placeholder lock or
  pointing the tools at an unselected engine is forbidden.

The authoritative deferred product contract remains
`Specs/Plans/WIP/Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`,
especially **User-facing contract**, **Architecture decision**, and
**Terminal plugin configuration in Settings Store v16**. Current commands,
shortcuts, Folder `Edit`, external-shell behavior, Preferences, and plugin
interfaces remain unchanged by Gate 0. Repurposing
`cmd/pane/openCommandShell`, adding `Alt+7`, migrating `Ctrl+Enter`, adding
`Ctrl+Shift+Enter`, retiring the pseudo prompt, opening opposite-pane tabs,
selecting Windows PowerShell/cmd/pwsh versus exact WSL profiles, adding
`followPathWhenIdle`, exposing every Terminal Preferences field, and
introducing `IID_ITerminal`/generic host plumbing are all deferred until an
engine wins. None may be partially implemented through `IViewer`, the EXE, or
a placeholder engine.

For a later selected engine, the Terminal plugin owner owns the engine/source,
adapter, and dependency lock; the security reviewer owns advisory disposition;
and the release owner owns notice/package publication. The required cadence is:

- structural `Locked` validation on every terminal-affecting change and release;
- exact cached-input verification after every restore and on every official
  native build lane;
- read-only `Online` observation monthly, immediately before a release, and
  within one business day of a relevant security advisory; and
- a full exact-pin upgrade comparison before any tracked engine/dependency
  change, even when the upstream change appears patch-only.

The status vocabulary is deliberately non-operative: `current`,
`update-observed`, `security-review`, `required-manual`, and `unreachable` report
facts. The pre-winner scaffold compares only each closed discovery record's
kind, credential-free URI, optional JSON property, frozen `currentValue`, and
advisory bit; `security-review` means a marked advisory observation changed,
not that an advisory database was semantically interpreted. In this scaffold
`-RequireOnline` rejects `unreachable` only. None selects a target, downloads
executable input, or edits a lock.

Before a winner-bearing Gate-0 round may approve a production lock, the closed
discovery schema, tool, and local-fake tests must add the stable tracked
channel/ref or release feed, prerelease policy, separate advisory source,
expected repository owner and canonical final URI, disappeared-pin result, and
ambiguity policy. Redirects, fork/owner drift, prereleases, advisories,
unreachable endpoints, missing pins, multiple-result ambiguity, and unreviewed
dependencies must each be proven. Production `-RequireOnline` rejects
unreachable or ambiguous observations. Until then, Online status is advisory
scaffold output and is not sufficient for a production engine lock.

The current v1 lock schema is also a pre-winner scaffold because its mandatory
`upgrade.basedOnLockSha256` cannot represent the first selected engine. Before
a winner-bearing round, replace it with the master's v2 `provenance` `oneOf`:
`first-selection` binds the exact pre-lock universe/review/assessment/
candidate-set/finalization and lock-independent candidate-materialization
identity plus the selected candidate ID/pin; `upgrade` alone binds a real
previous-lock digest, target source digest, and changed categories. The
materialization record binds reviewed expected outputs for all five lanes but
does not mention the final lock. A sentinel/placeholder prior-lock digest or
decision/lock/Verify-manifest hash cycle is invalid. The sole initial writer is
the future non-overwriting `New-TerminalEngineLock.ps1 -Mode FirstSelection`.
Before it runs, the closed `TerminalEngineMaterialization` v1
schema/corpus/writer, dedicated shared lock-payload schema/corpus, closed
candidate build-receipt/verification-receipt/lane schemas, and lock-independent
candidate verifier must produce exactly five native lane records plus one JCS
record containing the pre-lock selection identities, the shared closed
lock-payload projection, and the exact payload-schema digest, with no lock
path/digest. Each lane binds separate subordinate build and verification
receipt digests and never its own digest; the future lock binds the same
payload-schema digest. Mixed/nonnative/missing lanes, altered projections,
wrong schema digest, schema network resolution, receipt/lane self-reference,
and any final-lock reference fail. Canonical five-lane
Build/Verify runs occur only after the lock exists and their lock-bound
manifest-set digest is recorded in this decision/the successful Gate-0 handoff,
not inside the lock. `Prepare-TerminalEngineUpgrade.ps1` remains exclusive to
subsequent upgrades.

The first winner must execute the complete non-overwriting sequence in
`Specs/Plans/Done/Terminal_EngineSelectionAndDependencyProof_2026-07-22.md`
under **First winner: exact runnable selection-to-lock sequence**: five
lock-independent candidate Build/Verify lanes and exact receipt transfer;
materialization plus two non-author approvals; a candidate lock below `.build`
plus two non-author lock approvals; one reviewed tracked adoption; then five
lock-bound Build/Verify pairs and the verification-set digest handoff. Failure
of any pre-lock lane creates no materialization/lock/adoption and never selects
the runner-up; failure after adoption bars closeout and requires reverting the
whole adoption before a fresh approved attempt. This paragraph is normative and
cannot be replaced by the shorter upgrade workflow below.

Before the first shipped-lock upgrade, the master requires the closed
`TerminalEngineUpgradeReview` v1 schema/corpus/module. Its exact JCS record
binds both lock digests and the target, exactly 14 canonical category snapshot
reviews, and an ordinally sorted unique repository-relative path/action list
with conditional current/candidate file digests. Collect must accept
explicit candidate-lock/review files plus both expected digests and bind their
paths/digests in all five lanes; Finalize must receive the same exact
transferred lock bytes and review JSON/companion, reopen/recompute both
identities, validate the companion, and copy/bind both records. Lane references
alone cannot substitute for those bytes. The separate binding avoids a lock/review hash
cycle. The production workflow cannot run until path traversal, extra/missing
category/action, ordering, digest drift, mixed-lane, and staged-file mutation
tests pass and two non-author reviewers approve the candidate-lock,
review-record, and final comparison digests.

`Prepare-TerminalEngineUpgrade.ps1` deliberately does not guess or generate a
candidate lock. The Terminal plugin/dependency owner first runs a
non-archiving discovery build below `.build` and authors one complete candidate
lock there against `TerminalEngineLock.schema.json`. Every affected source,
tool, dependency, runtime, license, output, ABI, capability, qualification,
lane, and upgrade field—including the five expected artifact identities—is
replaced with exact discovered data; every retained field is reverified. The
lock binds the current-lock digest, exact target pin/source digest, and complete
derived change set. Two non-author reviewers approve its exact SHA-256 after
its locked inputs have been restored and verified offline. The byte-identical
lock, review JSON/companion, and content-addressed cache are transferred to each
exclusive disposable native runner. The same lock/review records return with
the five manifests to the review machine. `Collect` consumes them; `Finalize`
reopens the supplied records, verifies their expected digests across all five
lanes, and copies those same bytes beside the comparison. It rejects absent or
substituted same-path bytes and never constructs or repairs either record.

The routine dependency check is intentionally easy and read-only; upgrade
discovery is intentionally not one-click. Discovery is a separately reviewed
manual boundary whose inputs are the current-lock digest, one canonical
credential-free no-redirect target URI, exact target pin, source size/digest
established by the discovery transfer, and candidate-specific upstream
build/tool identities. Source acquisition rejects redirects and verifies the
bytes. Non-archiving exploratory builds run only on exclusive disposable
matching-native machines. Their output is the complete candidate lock plus the
closed `TerminalEngineUpgradeReview` record binding the current/candidate lock
digests and target source digest. It contains all 14 changed/retained canonical
category digests and, for every changed category, each normalized
repository-relative path, `add|modify|remove` action, role, and conditional
current/candidate file digest. It contains no absolute path, command,
credential, source payload, or terminal/user content. Two non-author reviewers
approve both digests, and every subsequent shipped-lock upgrade binds the
record digest into evidence before `Collect`. An unauthenticated source,
ambient build dependency, placeholder field, unbound record, or unreviewable
output/ABI identity stops the upgrade before official collection.

The required five-stage lifecycle is:

1. Verify the checked-in lock structurally and, when requested, verify every
   cached input offline.
2. Observe canonical upstream/advisory sources read-only; never select or apply
   “latest”.
3. Prepare one exact target pin and pre-reviewed candidate lock under `.build`,
   collect the five native winner lanes only on exclusive disposable
   matching-native runners, and compare dependency/SBOM/license/ABI/capability/
   security/performance/package deltas.
4. Apply the approved lock, adapter, runtime manifest, notices, tests, and docs
   as one reviewed repository change.
5. Rebuild and verify the full matrix; rollback by reverting the complete change
   and reconstructing from the previous lock.

The normal operator commands are:

```powershell
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock <winner-lock>
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile <winner-lock>
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile <winner-lock> -Offline
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock <winner-lock> -RequireCache
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Online -EngineLock <winner-lock> -RequireOnline
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile <exact-candidate-lock-below-.build>
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile <exact-candidate-lock-below-.build> -Offline
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock <exact-candidate-lock-below-.build> -RequireCache
.\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Collect -TargetPin <exact-target-pin> -TargetSourceSha256 <64-lowercase-hex-target-source-digest> -ExpectedCurrentLockSha256 <64-lowercase-hex-current-lock-digest> -CandidateLockFile <exact-candidate-lock-below-.build> -ExpectedCandidateLockSha256 <approved-candidate-lock-digest> -UpgradeReviewFile <exact-reviewed-upgrade-record-below-.build> -ExpectedUpgradeReviewSha256 <approved-upgrade-review-digest> -Platform <x64-or-ARM64> -Configuration <lane> -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -UpgradeRoot .\.build\TerminalEngineUpgrade -PassThruManifest
.\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Finalize -TargetPin <same-exact-target-pin> -TargetSourceSha256 <same-64-lowercase-hex-target-source-digest> -ExpectedCurrentLockSha256 <same-64-lowercase-hex-current-lock-digest> -CandidateLockFile <exact-transferred-candidate-lock-below-.build> -ExpectedCandidateLockSha256 <same-approved-candidate-lock-digest> -UpgradeReviewFile <exact-transferred-upgrade-review-below-.build> -ExpectedUpgradeReviewSha256 <same-approved-upgrade-review-digest> -EvidenceManifest <five-exact-manifests> -UpgradeRoot .\.build\TerminalEngineUpgrade -RequireAllEvidence -PassThruManifest
.\Tools\Build-TerminalEngine.ps1 -LockFile <winner-or-candidate-lock> -Platform <x64-or-ARM64> -Configuration <lane> -Clean -Offline -ExclusiveDisposableRunner
.\Tools\Verify-TerminalEngine.ps1 -LockFile <winner-or-candidate-lock> -Platform <x64-or-ARM64> -Configuration <lane> -RunSmoke -Offline -ExclusiveDisposableRunner
```

`Collect` requires exact x64 Debug, x64 Release, x64 ASan Debug, ARM64 Debug,
and ARM64 Release lanes under `.build`, exact pre-reviewed candidate-lock and
review files plus expected digests below `.build`, and the explicit
exclusive-disposable-runner assertion because offline isolation is host-wide.
`Finalize` accepts only those five explicit digest-bound manifests plus the
transferred lock/review bytes, validates the review companion, and makes no
tracked edit. The approved repository change
then moves the candidate lock, adapter/runtime declaration, notices, tests,
decision, and documentation together. The previous lock remains the rollback
identity, and rollback is proven by reverting the whole change and reconstructing
its artifacts offline. Finalize reports changed comparison categories; the
separately approved review note is the authority for the exact
repository-relative path/action patch.

Stop rather than upgrade when the current lock digest is unexpected, a source
or tool byte cannot be verified, an endpoint redirects or needs credentials, a
dependency/license/notice/ABI/capability delta is unresolved, a native lane is
missing, a security or performance regression is unexplained, the runtime
closure is incomplete, or the change would require an irreversible user-state
migration. Emergency security work may shorten review latency, but it may not
weaken exact-pin, two-reviewer, native-matrix, license, ABI, or rollback gates.

These commands intentionally cannot be run against a production engine lock
until a later Gate 0 decision creates one.

### Round 4 automation boundary

`.github/workflows/terminal-engine-gate0.yml` is the frozen
executable-cohort orchestration template, SHA-256
`e6f15b6fef9c0d367e12dc7f43cb769106dc982934658f501120728599f32d1c`.
It is part of the approved collection-neutrality surface and therefore cannot
be edited after the v4 freeze. It assumes a nonempty executable cohort and is
not a Round-4 closeout workflow. Dispatching or re-enabling it while Round 4 is
current is a closeout-policy violation; its expected static-only failure is not
a collection result.

The manual-only
`.github/workflows/terminal-engine-round4-closeout.yml` and
`Tools/Test-TerminalEngineRound4Closeout.ps1` own the current automation path.
They authenticate the exact v4 closure, require zero collect dispositions,
prove evaluator exit 21 with empty success output, prove the direct provider
rejects the complete derived static cohort, require the evidence path to remain
absent, require no archived TerminalEngine root or product/winner state, run
the exact 239-test focused cohort and the full suite, repeat the no-winner
repository proof after those tests, and create no artifact.
Before checkout, the workflow queries the GitHub Actions control plane and
refuses to run unless `terminal-engine-gate0.yml` is `disabled_manually` and
has no active run or historical manual dispatch. Immediately before publishing
its closeout summary, it requires the workflow to still be
`disabled_manually` and rejects any executable-workflow run created at or after
the closeout run began.

The executable workflow is not yet registered on origin's default branch, so
there is no remote workflow state to mutate at this local closeout. Publication
has an explicit merge gate: after the frozen workflow is first registered on
the default branch and before any manual dispatch, the release owner runs
`gh workflow disable terminal-engine-gate0.yml` and verifies the API state is
`disabled_manually`. It remains disabled until a separately approved
v5-or-later universe refreezes the complete collection authority and admits an
executable cohort. The Round-4 closeout workflow enforces the disabled/no-run
condition at both boundaries of its bounded run and rejects overlap. Repository
permissions and release-owner policy must keep it disabled before and after
that run; no completed workflow can enforce future control-plane state.

## Reopening Gate 0

A later attempt must open a new, independently approved universe before any new
performance result. Valid reasons are a materially improved exact upstream pin,
a newly credible existing terminal engine, or an explicit product decision made
before measurements to change a requirement/cap. It must:

1. keep all unchanged source decisions, controls, fixtures, and measurement
   recipes byte-bound and explain every intentional change;
2. regenerate every nomination identity, patch/result tree, concern map, owner
   estimate, approval, assessment, candidate set, and provider proof;
3. freeze one exact candidate-neutral host-header archive per platform input
   set and assign it to the complete executable cohort; remove only its exact
   contained prior expansion and freshly safe-expand the verified archive in
   every consuming phase, reject any reparse point from the filesystem root
   through `include/wil/resource.h`, bracket the candidate driver with
   complete-tree verification, reverify immediately before and after each
   neutral Harness/Probe build, disable ambient vcpkg, and prove neither
   between-phase cache tampering nor a repository/developer
   `vcpkg_installed` tree can influence the build;
4. accept only canonical `-Candidate All` collection records; reject narrowed
   or static-only selections before repository inspection, bootstrap, clean,
   subprocess launch, or evidence-root mutation;
5. collect every admitted candidate and every control on native x64 Release,
   x64 ASan Debug, and ARM64 Release; and
6. select only a mandatory-pass candidate scoring at least 80.00 under the
   frozen cohort and tie-break rules.

Reducing a measured cap, merging concerns, moving candidate code into host
tools, accepting future upstream work, or selecting Ghostty by preference is
not a valid reopening.

### Requirement-change decision matrix

This matrix is the explicit escape hatch requested for a no-winner result. It
is planning guidance only: no row changes Round 4, authorizes collection,
creates a production lock, or selects an engine. Whichever row is chosen must
become a candidate-neutral v5-or-later universe change, receive its own two
non-author approvals, and freeze before any new measurement or executable
collection.

| Route | Requirement change | Candidate it could admit | What it does **not** prove | Decision |
|---|---|---|---|---|
| A — manufacture a conforming candidate | None. Replace WezTerm's asynchronous adapter with an actual synchronous, caller-owned boundary while preserving every cap and G1–G12. | A new WezTerm-RS v5 nomination, if the real patch stays at or below 2,500 lines, 24 files, five concerns, one product-new toolchain family, zero runtime leaves, 60 owner-days, and contains no binary patch. | The current v3 patch is not admissible; the 59-day calculation below is a near-cap hypothesis, not a zero-slack result. The mechanically recomputed value alone governs admission, and exactly 60 owner-days remains valid. The new patch must still pass all native lanes and score at least 80.00. | **Recommended.** It preserves product, security, plugin-DLL, and maintenance requirements. |
| B — buy maintenance headroom | Raise the candidate-neutral owner cap from 60 to 90 days and the logical-concern cap from five to six; leave every other cap, mandatory gate, native lane, score anchor, and 80.00 threshold unchanged. | The existing WezTerm-RS v3 adaptation shape becomes eligible for fresh nomination/approval and collection. | Eligibility is not a win. G1–G12, x64 Release, x64 ASan Debug, ARM64 Release, packaging, and score may still reject it. This accepts about 50% more three-year owner effort and one more independently regressible concern. | Honest fastest policy override if Route A fails. It must be an explicit product/maintenance decision, not a candidate-specific waiver. |
| C — buy a larger C++ patch | Raise the candidate-neutral line cap to 3,200, concern cap to six, and owner cap to 80 days; leave all safety, runtime-leaf, native-lane, and score gates unchanged. | Contour-RS v3 could be re-nominated only after repairing and freshly proving its Kitty image-versus-text accounting. | The current accounting observation is unresolved and the candidate has not passed G1–G12 or scoring. | Higher patch-maintenance surface than A; prefer only if avoiding the Rust boundary is worth that cost. |
| D — defer Kitty graphics | Remove mandatory G7 Kitty APC/static-image support from v1 and refreeze the fixtures/capability ledger honestly. | libtsm, libvterm, and TerminalCore could receive new source review instead of immediate G7 rejection. | It does not establish reflow, shaping, snapshot ownership, resource accounting, input, performance, or packaging fitness. It is a visible product feature cut, not an engineering shortcut. | Valid only if product explicitly removes Kitty images from the first release. |
| E — relax candidate-owned/native control | Permit host-owned terminal implementation, a web stack, or a new in-house parser. | Microsoft/full-control, WebView2/xterm.js, or a RedSalamander parser. | It abandons the plugin-DLL-only architecture, native/no-second-web-stack boundary, or maintenance/security constraint. | **Do not choose** under the stated product requirement. |
| F — raise a Ghostty cap now | No evidence-based finite change is available: the v4 patch is incomplete and already exceeds the owner cap before its missing Gate-0 adapter. | None yet. | Raising owner-days to 90 cannot prove the final line, file, concern, ABI, teardown, or accounting footprint. | Finish a complete measurable adapter in a newly frozen nomination or wait for an upstream API; do not manufacture a Ghostty win by preference. |

Route A is authorized as the next evidence-only Gate-0 round by
`Specs/Plans/WIP/Terminal_EngineSelectionAndDependencyProof_Round5_WezTermSynchronousCallerOwned_2026-07-31.md`.
That authorization changes neither the Round-4 result nor any product gate.
Route B remains an unselected fallback requiring an explicit later product and
maintenance decision.

### Recommended “manufactured winner” experiment

The most promising bounded experiment is a synchronous caller-owned WezTerm
redesign in a newly approved v5-or-later universe, not an estimate-only
resubmission. Freeze the exact upstream pin, synchronous ownership hypothesis,
unchanged Round-4 caps/gates/fixtures, candidate-neutral accounting rules, and
a 15-owner-day proof-spike limit before the spike starts. The spike is useful
code and evidence, not an off-ledger prototype: every retained candidate-
specific source, build, test, documentation, generated, and descriptor path
must appear in the final patch/result-tree and owner ledger.

The proof spike must expose a caller-owned create/destroy/feed/resize/snapshot/
input-encoding/Kitty query boundary with no adapter background thread, async
callback, `Arc`-retained session, or post-destroy work; prove exact borrowed-data
lifetime and pre-allocation resource accounting; and build the same source on
native x64 and ARM64. It stops immediately and without extending the time box
when cumulative proof-spike effort reaches 15 owner-days, even if the hypothesis
is still incomplete. It also stops as soon as the completed v5 design retains,
or the proof makes unavoidable, a sixth independently regressible concern; a
second product-new compiler/language toolchain family is required; a non-system
runtime leaf or binary patch is required; the complete maintained patch cannot
fit the unchanged line/file caps; the mechanically recomputed upper ledger
exceeds 60 days; ARM64 cannot use the same boundary; or G7/G9 data cannot be
bounded before allocation. The measured v3 baseline already has six concerns;
that transient baseline is charged to the proof but succeeds only if the
async/`Arc` lifecycle concern is demonstrably removed from the completed v5
design. A successful spike merely admits preparation of the complete v5
nomination; it then needs the exact patch/result-tree/concern-map approvals, all
G1–G12 lanes, packaging evidence, and score of at least 80.00.

Its current 2,459-line/seven-file patch already fits
the line/file caps, but its asynchronous `Arc`/atomic callback lifecycle makes
`teardown-concurrency` a real sixth concern and its boundary carries eight
adapter plus five per-upgrade days. A genuinely synchronous, caller-owned ABI
that removes that lifecycle behavior could remain at seven source paths plus
one descriptor, five concerns, adapter floor 5, patch floor 10, per-upgrade
floor 3, security 2, and notice 1. The frozen formula would then derive
`ceil((5 + 10 + 9*3 + 3*2 + 1) * 1.20) = 59` owner-days. A new v5-or-later
universe may test that design only if the code actually removes the concurrency
concern and two reviewers accept the lower upper-bound estimates; relabeling
the current patch does not qualify. Round 4 is chronology-only and cannot host
that candidate-changing experiment. The 59-day result has one reported
owner-day of headroom. Because the formula applies coefficients before the 1.20
multiplier and ceiling, it can absorb one additional day in an unweighted term
such as `initialAdapterDays`, `initialPatchDays`, or `releaseNoticeDays`:
`ceil(50 * 1.20) = 60`. It cannot absorb a one-day increase in
`perUpgradeDays`, which contributes nine pre-multiplier days, or
`annualSecurityResponseDays`, which contributes three. The independent
15-owner-day proof-spike limit remains unconditional. Stop on a mechanically
recomputed result greater than 60 or any other hard-cap violation; do not reject
an exact result of 60. Ghostty is not comparably close because its current
evidence is an incomplete, already-over-cap lower bound.

## Evidence and approvals

Decision closeout UTC: `2026-07-31T17:42:43Z`.

The exact Round 4 governance closure is:

| Record | Exact SHA-256 and closeout facts |
|---|---|
| `Specs/Terminal/TerminalEngineCandidateUniverse.v4.json` | `7621ab8562cb4b1d530d24355c67d90c905781a5d4f6bbc3364699ac059cb740`; proposal `22ba55a19bfa631093af06729b54f4eba46c5e23198f8f421accad634d106129`; frozen `2026-07-31T01:47:10Z`; approved against that proposal by `codex-dalton` at `2026-07-31T01:51:32Z` and `codex-contour-independent` at `2026-07-31T01:53:33Z`. The latter instant opens the Round 4 downstream-review window. |
| `Specs/Terminal/TerminalEngineCandidateReview.v1.json` | `ba2942c7abe8f94b6f50be22ddb78570eaade5f60a152a8b0cca48dcea91542d`; author `codex-dalton`; three exact source-rejected candidates; frozen `2026-07-31T02:29:00Z`. |
| `Specs/Terminal/TerminalEngineCandidateAssessment.v1.json` | `adf5ae89c3c90fd6018652f9ebd88aa809048a849ba29f49b04a4fde6d20343e`; author `codex-dalton`; zero admitted candidates and therefore no candidate-assessment approvals; frozen `2026-07-31T02:29:00Z`. |
| `Specs/Terminal/TerminalEngineCandidateSet.v1.json` | `fb240d3ef838daea0740417fafd2d7d44b9536295f16c607e45fe57e64d11f4d`; author `codex-dalton`; six seed outcomes, zero `collect` dispositions, and three unscored controls; frozen at the same `2026-07-31T02:29:00Z` common instant. Both x64 and ARM64 static bootstrap input sets have identity `294b689d4577dc36b0e91d9b12f7c7209c7ba8a6d087c9bb0353c264fdbebe41`. |

The fresh source approvals are strictly inside that window. `codex-root`
approved all three blocker subjects at `2026-07-31T02:27:20Z`, and
`codex-contour-independent` independently approved all three at
`2026-07-31T02:28:37Z`:

| Source-rejected candidate | Approved blocker SHA-256 |
|---|---|
| libtsm | `9031fe57aaa2b06a318ef3f9ce8d00cc63c310699713ce7541874c83102c7661` |
| libvterm | `83f5e186cf0d7ab913c87a7ef270ca179ec766079a3d4d47aab7de7cb5456f02` |
| Microsoft TerminalCore | `4767963af7a2f9a76dfbecc809286c71915e5ae8cd5073d060b6f967e03fe920` |

The fresh adaptation-rejection approvals are also strictly inside that window.
`codex-root` approved each exact evidence subject at
`2026-07-31T02:27:20Z`, and `codex-contour-independent` independently approved
each at `2026-07-31T02:28:37Z`:

| Nomination | Patch SHA-256 | Concern-map SHA-256 | Approved rejection evidence SHA-256 | Mechanical result |
|---|---|---|---|---|
| Contour-RS v3 | `26a41b1288c2a2b25e6933954fac543a5b237b359fbc175733a170b75eb8f3a5` | `e1111e62edf6e09fb34d00084d1bdcab8e3985a23ce9099e5c65cb99953f1f86` | `3c74ef9a9d738bbaf4305f6cfaba4d80ed4ccbff454b339090d7a5fa6dd00022` | 3,041 changed lines, 15 files, six concerns, 76 owner-days: line, concern, and owner caps fail. |
| Ghostty-RS v4 | `e85d4a504dcf61bfd16b9af37b7e8b53ea6300e558264a86ee56060ae5457c3a` | `cd9489bd2c6790b28c82781a31ee1210587479bee766325ba8bff6ceabf24500` | `e90f8fff84edc066ab849e9c45ac861c61ec195a8894e5b87978b5c09684b432` | 2,213 changed lines, 21 files, five concerns, 84 owner-days: owner cap fails. The incomplete patch already has an 81-day policy-floor lower bound before its missing adapter. |
| WezTerm-RS v3 | `205a95a633da0e94548da485644cdb220fcdf9809e3f5cf35b1443dc23dac72e` | `ae9d132585d7003a32003ce65c2c63a010de736c2cd49054761412a3ac3e8340` | `00535c77295d9ed7768d0fc65514fa3897493000395dcefeda0075b2be71e6e1` | 2,459 changed lines, seven files, six concerns, 87 owner-days: concern and owner caps fail. |

These schema-bound approvals approve the exact universe proposal, blocker
subjects, and adaptation-rejection subjects defined by their schemas. They do
not purport to approve later file hashes. The final Markdown decision is
separately cold-audited below against a non-circular prefix digest.

The immutable-consumption and no-effect evidence is:

- the Round 4 collection-neutrality policy is
  `Specs/Terminal/TerminalEngineCollectionNeutrality.v1.json`, SHA-256
  `7f9eee4b83576357a9407c39681c5d3f6a0d25a0d13bd30b3fb6885286863d77`;
- the frozen executable-cohort workflow has SHA-256
  `e6f15b6fef9c0d367e12dc7f43cb769106dc982934658f501120728599f32d1c`;
- `Tools/Test-TerminalEngineRound4Closeout.ps1` reauthenticated the complete
  v4 closure, derived all six outcomes, and exercised both public entry points.
  The evaluator returned exit `21` with zero success-stream bytes and the
  static-only diagnostic; the direct provider also rejected the complete
  derived static cohort. The evidence root was absent before and after, no
  archived `TerminalEngine` root existed, and no engine lock, product Terminal
  directory/ABI, winner build root, or runtime-manifest change appeared. The
  command was
  `.\Tools\Test-TerminalEngineRound4Closeout.ps1 -RepoRoot <workspace> -EvidenceRoot <fresh-absent-.build-path>`.
  Its disposable ignored stdout at
  `.build/terminal-engine-spike/round4-governance/tracked-closeout2.stdout.txt`
  has SHA-256
  `8c64f13d3018f89ce04c433173315190c0c83764701ccebc30c4dc8ca636291f`.
  That local log is not durable clone evidence; the checked-in tool, its
  focused test, and the exact governance records are the reproduction path;
- the required post-full-suite repetition started `2026-07-31T16:59:44Z` and
  completed `2026-07-31T17:05:00Z` with
  `Status=valid-no-winner-product-blocked`, six candidate records, zero
  `collect` candidates, evaluator exit `21`, zero success-stream bytes, direct
  provider rejection, absent evidence roots before/after, no archived
  `TerminalEngine` root, no winner lock, no product Terminal directory/ABI, no
  winner build root, and a frozen runtime dependency manifest. Its archived
  stdout at
  `Specs/TestRuns/7d3a1247382a/Continuation/2026-07-31_161745_terminal_gate0_commands_plugins_uia_hang/zero-cohort-post-full.stdout.txt`
  has SHA-256
  `cb9c0eb8a415d4f040e6299b71b5603b76f3003b556029219d9048a4a240cbe4`;
  its empty stderr has SHA-256
  `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`;
- `.github/workflows/terminal-engine-round4-closeout.yml` is manual-only,
  requires the frozen executable workflow to be disabled with no active run or
  historical manual dispatch before checkout, rechecks disabled state and
  rejects any overlapping executable run immediately before its summary,
  invokes the no-winner proof both before and after the exact focused/full
  tests, and uploads no artifact. The frozen workflow is not registered on the
  current origin default branch, so the publication merge gate—not this local
  run—must first disable and verify it after registration.

Verification on the closeout workspace:

| Verification | Exact result |
|---|---|
| Gate-0 focused Pester set | Started `2026-07-31T03:33:15Z`, completed `2026-07-31T03:44:31Z`: 20 files, 239 passed, 0 failed, 0 skipped, 0 pending, 0 inconclusive, 675.325 seconds. Disposable ignored `.build/terminal-engine-spike/round4-governance/focused-final2.stdout.txt` has SHA-256 `2fd9ccef6c46534c608c5fbef139710af2f08598640c88ca45518c25a3b00847`; sibling `focused-final2.stderr.txt` contained only Git line-ending warnings and has SHA-256 `d2d2470773f41a8075b387bd73d02148cbf72f0c2fce554e85321cee8aa10250`. A second exact post-runtime-edit run immediately before the frozen-review commit again exited `0` with all 239 tests passing in 699.5 seconds. After clearing the historical sandbox residue disclosed below, accepted isolated run `r4clean-284e0752` started `2026-07-31T17:30:43Z`, completed `2026-07-31T17:41:25Z`, passed the same 20 files and all 239 tests in 639.63 Pester seconds, then returned `SandboxClean=true` and `SandboxIssueCount=0` in the same process. Its archived stdout SHA-256 is `b253a566b3bcfee8a3f2878c1e6d06e01a077ae57fb19a2fc87e6cb5c4d6873c`; stderr contains only Git line-ending warnings and has SHA-256 `d2d2470773f41a8075b387bd73d02148cbf72f0c2fce554e85321cee8aa10250`. |
| Dependency-lifecycle Pester contract | Started `2026-07-31T15:32:37Z`, completed `2026-07-31T15:32:44Z`: 12 passed, 0 failed, 0 skipped, 0 pending, 0 inconclusive; 7.005 wall seconds and 5.69 Pester seconds. It used only synthetic locks, local fake remotes, and disposable build roots. |
| `Tools/Run-AllTests.ps1 -Suite Full` | **Completed with classified pre-existing non-Terminal baseline exceptions; the canonical wrapper is not claimed green.** Outer wrapper started `2026-07-31T15:38:29Z`, completed `2026-07-31T16:56:01Z`, and exited `1`; its build succeeded with zero warnings/errors. Aggregate run `20260731T153831Z-35744-9dceccf8fcac42a8b91d71f30c3a489e` contains 17 suites, 776 total entries, 722 passed, 3 failed suite/case entries, and 51 skipped. Fourteen suites passed; the three classified entries are detailed below. The aggregate also honestly retains a separate nine-item dirty-TestSandbox audit, which is disclosed and closed below rather than counted as a suite/case entry. |

The full-suite classification is bounded and reproducible:

- `Commands` passed 371 ordinary cases with zero ordinary failures before the
  already-known
  `cmd_preferences_dialog_plugins_live_search_dx_interaction` UIA/D3D boundary
  stopped returning. The runner's native-process `WaitForExit()` is unbounded,
  so after more than 15 minutes without trace/result progress the exact
  `RedSalamander.exe` child was stopped only after live process/thread/trace
  capture. Its aggregate `selftest_result_coverage` failure is the expected
  consequence of every later Commands case being absent. The same outer-case
  non-return occurred in a pre-freeze run while HEAD was still base
  `1a5d838d0c18298d9e560e243e67b821a40fcf6a`, and the same clean-process
  UIA/D3D stack is archived from 2026-07-21. Owner classification:
  Preferences/DxUi self-test infrastructure;
- `ViewerPETests` failed only
  `ViewerText shows one localized warning after truncating an over-cap hex
  clipboard selection` and its enclosing fresh-harness assertion. The exact
  same failure exists in the archived pre-freeze base continuation. Gate 0
  changed no Viewer product/test source. Owner classification: ViewerText;
- `ToolsPesterTests` completed 584 passes and three unrelated baseline failures:
  a documentation-drift test still reads an already-moved File Operations plan
  from WIP; a historical-review banner assertion expects an ASCII hyphen while
  the unchanged file uses an em dash; and the unchanged source-inventory
  contract expects 725 Commands registrations while deriving 714. The two
  failing Pester files and all named inputs have an empty name-status diff from
  the base commit to frozen review
  `3e84e87b717841e28e11e1fe7e5bc04b4e3a3e0d`. Owner classification:
  Testing/documentation baseline maintenance;
- the wrapper completed every remaining lane: Compare Directories passed 227
  and skipped 31; all 18 File Operations families completed with 112 passed
  and 20 environment-gated skips; DxUi, FileSystemCurl, ViewerSqlite, Monitor,
  Localization, RedConfigure, PluginContract, SettingsSchema, CrashHandling,
  monitor ETW latency, PerformanceTests2, and VcpkgMergeSynthetic all passed.
  The Tools Pester TerminalEngine Round-4 closeout case itself passed in 267.97
  seconds. No failed entry is owned by, introduced by, or exercises a product
  Terminal implementation.

The full aggregate additionally records
`test_sandbox_audit.is_clean=false`, `issue_count=9`, for nine
TerminalEngine-named fixture directories directly below the unified
`TestSandbox` root. This result is not omitted or reclassified as one of the
three failed suite/case entries. Every directory's creation/last-write time was
between `2026-07-30T22:18:12Z` and `2026-07-30T23:01:46Z`, before both final
focused runs and the full-suite start, and neither later run changed it. An
exact prefix search of frozen commit
`3e84e87b717841e28e11e1fe7e5bc04b4e3a3e0d` found zero source matches. The
residue was therefore earlier Round-3/ad-hoc test material, not a leak
reproduced by the frozen tests.

After verifying all nine absolute paths stayed below the intended sandbox, the
operator moved exactly those disposable directories recoverably to an ignored
quarantine outside `TestSandbox`. An immediate audit returned clean. The
accepted exact focused rerun described above then passed `239/239` and its
same-process post-Pester audit again returned clean with zero issues; the root
contained only `runs`, and only the canonical full-suite run plus the
explicitly allowed accepted rerun existed there. The original aggregate stays
unchanged. The exact issue list, timestamps, command, disposition, accepted
logs, and hashes are archived in
`Specs/TestRuns/7d3a1247382a/Continuation/2026-07-31_161745_terminal_gate0_commands_plugins_uia_hang/test-sandbox-remediation.md`.

The exact aggregate, current/prior traces, prior-base Viewer log, owner
classifications, zero-impact diff, and wrapper streams are archived at
`Specs/TestRuns/7d3a1247382a/Continuation/2026-07-31_161745_terminal_gate0_commands_plugins_uia_hang/`.
The aggregate SHA-256 is
`e9a9e7bf6b5e2882753d5c5ff566121f3b1c914c089cf1c129e7c9032f7998f4`;
final wrapper stdout/stderr SHA-256 values are
`31be7113f12e4fd528f14abbd7bbedec914c4e5fba8e9305c9f52439bcbed29b`
and
`d2d2470773f41a8075b387bd73d02148cbf72f0c2fce554e85321cee8aa10250`.

The closeout tree contains no `External/terminal-engine.lock.json`,
`Plugins/Terminal/Terminal.dll`, product Terminal implementation, terminal
entry in `RuntimeDependencies.props`, scored collection manifest, finalization
manifest, or winner artifact. Those absences are required results, not missing
evidence. No `Specs/TestRuns/.../TerminalEngine/...` run is archived because
the executable cohort is empty.

### Decision closeout audits

For a reproducible non-circular review identity, each auditor hashes the exact
UTF-8 bytes from the start of this file through the final LF immediately before
this heading. This heading and everything after it are excluded. The reviewed
prefix SHA-256 is
`c9cf3f18ebe2fb53736c9589bdcbad5d441d1e5cfcd98d8ab07ba7d2e362e3b3`.

- `codex-round4-cold-audit-one` at `2026-07-31T17:48:06Z` independently
  recomputed
  `c9cf3f18ebe2fb53736c9589bdcbad5d441d1e5cfcd98d8ab07ba7d2e362e3b3`
  and returned **CLEAN — no P1/P2 findings**.
- `codex-round4-cold-audit-two` at `2026-07-31T17:47:47Z` independently
  recomputed
  `c9cf3f18ebe2fb53736c9589bdcbad5d441d1e5cfcd98d8ab07ba7d2e362e3b3`
  and returned **CLEAN — no P1/P2 findings**.

## Post-closeout Round-5 executable runbook addendum

This addendum is outside the immutable Round-4 reviewed prefix above. The
Round-4 decision identity remains the exact file blob and SHA-256 at closeout
commit `6674119fd24284877b1b25a3a71dbe2f6e6ee9a2`; consumers authenticate those
historical bytes from that commit rather than treating this mutable current
operator summary as the predecessor subject. For Round 5, this addendum and the
named Round-5 plan supersede the shorter historical first-winner sequencing
paragraph above. That protected Round-4 text does not authorize a Round-5
same-round post-effect retry.

Round 5 uses
`Tools/TerminalEngine/TerminalEngineRound5Coordinator.psm1` and its action
wrapper as the sole global admission/commit boundary. At most one pipeline
action may be active across every local and remote runner. The separate
serialized no-effect control plane owns live supersession operations and
post-quiet governance. Every `Pipeline` entry point requires its token path/raw
digest, expected epoch/input identity, runner-local claim, and unique create-new
output root. Every `ControlLive|ControlQuiet` entry point instead requires its
matching control token path/raw digest and expected epoch/input identity;
`ControlQuiet` also requires zero Pipeline actions and the exact quiet receipt.
Control operations never require or fabricate a runner claim or candidate output
root. Tokens bind `actionKind` and the canonical logical key; the append-only
journal permits at most one original token/claim and only the Round-5 closed
allowed-next transition. Every class rejects bypass before effects.

The reviewed clean P boundary admits the bounded proof and all post-proof
nomination/governance writers; after quiet, the reviewed clean A boundary admits
collection and first selection. Before any candidate effect, the coordinator
creates and reopens the exact digest-free input set, transfers byte-identical
bytes to all five native runners, completes all five preflight-only checks, and
derives rather than accepts the receipt disposition. Collection, Finalize,
input-set creation, and lock-independent lanes bind A. Lock-bound lanes start
from clean B adoption. P-candidate and A-candidate commits each require their
commit-bound focused-suite Pass and distinct quiet before approval. Review,
Assessment, and Set publish only as the complete bytes reserved by the
authoritative one-way GovernanceFreeze journal.

Only `CandidateLane` admission is pre-effect and retryable. If attempt 1 is
`RejectedPreEffect`, attempt 2 with the byte-identical input-set path, raw
digest, and RS-JCS identity plus fresh validation of all five preflights is
mandatory unless approved supersession closes admission first. Declining that
retry without supersession leaves the round WIP. A second rejection is terminal;
a third attempt and every later-stage retry are forbidden. After the action
stops and coordinator quiet is proven, the tracked branch-neutral ledger has
exactly `AdmittedFirstAttempt`, `AdmittedAfterRetry`, `SecondRejection`, or
`StoppedAfterFirstRejection`; it embeds the canonical input set, preflights,
tokens, receipts, and quiet receipt.

An admitted receipt opens the candidate pipeline. The invariant successful
order is five lock-independent lanes, materialization and approvals, candidate
lock and approvals, reviewed B adoption, five lock-bound Build/Verify pairs,
verification set, tracked dependency status/cache/notices reports, quiet,
attempt ledger, immutable winner-handoff proposal, and two separately immutable
ordered proposal reviews. The proposal binds the ledger and exactly the six
pre-closeout values: finalization path/digest, candidate ID, engine-lock digest,
and verification-set path/digest. It cannot name a future decision, Done
path/blob, closeout commit, or success marker. Reviews bind the proposal digest,
never vice versa; two distinct non-author Approve records are required.

If the second pre-effect attempt rejects, or any later non-superseded
substantive stage fails, the round closes no-winner through its versioned
immutable qualification-failure schema/corpus/record. The record binds the
winner finalization, selected candidate, branch-neutral ledger or its sole legal
validation substitute, failing stage/reason, and the exhaustive stage-valid
materialization/approval/lock/adoption/verification/dependency/proposal/review
state. It never selects the runner-up or patches in place.

Adoption cleanup is exactly `NotStarted`, `CancelledBeforeCommit`,
`CommittedThenReverted`, or `AdoptedThenReverted` as allowed by the branch. A
revert stages only the adoption path set, restores that projection byte-for-byte
to A, and proves zero engine-adoption state. Full-tree equality to A is required
only when no governance commit intervened; otherwise the tree may differ solely
by the closed digest-bound governance allowlist. A writer-only absent-output
pause reacquires a linked coordinator token for identical logical inputs, UTC,
and verdict and rechecks the epoch. It never reruns candidate work, preflight,
or reviewer judgment; if the epoch changed, the stale writer cannot publish.
Cleanup begins only after pre-cleanup quiet. When required, C is one resumable
manifest-bound `TransactionalCleanup` Git-boundary Pipeline action, and its
post-cleanup quiet is exactly C's post-transaction quiet. Only then may
Qualification/Supersession/D control governance run; `NotStarted` uses the exact
cleanup/post-quiet N/A values.

A possible post-v5 viability delta first creates an ignored immutable proposal
and ordered reviews. The proposal, not the final result, receives the two
reviews. `ObservationPending` starts no next pipeline action; the current action
may reach its permitted commit/quiet gate. The second Approve is pre-rendered,
then an authoritative no-replace fence journal advances to `Superseding`.
After the pre-fence action completes/cancels and quiet is proven, freeze the
cumulative observation ledger and final tracked Supersession result. The result
contains every action admitted before the fence, including one completed or
aborted afterward; post-fence admission is illegal. Every post-v5 closeout binds
the final cumulative observation ledger as
`Empty|RejectedOnly|ApprovedSupersession`.
After the fixed result, `SupersededCarryForward` leaves Pipeline/result
replacement permanently closed while every later pre-E observation appends
exactly the next create-new cumulative ledger generation. Reject preserves the
fixed result/branch; Approve carries the delta to the next ordinal universe.
Neither replaces the result. E binds the final explicit generation; overwrite,
gaps, forks, or implicit-highest discovery are forbidden.

Round 5 uses P plus six fixed logical Git boundaries: A execution freeze; B
adoption; C state-specific revert; D post-effect evidence; E audited
decision cutoff; and composite F closeout/handoff publication. Every transaction
uses the manifest/private-index/commit-object/ref-CAS recovery contract. After final D,
the commit-bound FinalFocused and FinalCanonicalFull Pipeline receipts plus their
quiet identities must close before the immutable Round-5 decision prefix
receives two generation-qualified
non-author audits. Approved bytes remain ignored until the coordinator's single
`ClosingPublish` CAS proves no pending observation, then the exclusive lease
publishes the fixed decision/current-pointer cutoff as E while leaving the WIP
plan, Done path, success marker, and master routing unchanged. Observation
admission then remains continuously sealed through deterministic recovery of
FArchive (E -> FA), FDone (FA -> FD), FPublish (FD -> FP), and FSeal
(FP -> FS). FArchive adds only the tracked RecoveryArchive. FDone alone
fills/moves the plan to Done and adds the success marker only for Winner, while
master routing remains unchanged. FPublish records the now-known FD closeout
commit/Done blob plus FA/path/digest and updates only master/README/next-plan
routing without changing decision/Done/archive bytes. After FP's fixed quiet,
FSeal adds only the canonical tracked CloseoutSeal embedding the E/FA/FD/FP
quiet set and binding those commits. E, FA, FD, and FP are non-publishable
intermediate refs. An
observation either wins before the E CAS while the decision path is absent, or
waits until verified FS for the separately approved post-publication
reselection/revocation process. A failed or superseded closeout publishes no
success marker. E only freezes the audited decision; FDone only creates a
dormant, non-publishable marker; FPublish creates routing, and only the valid
FSeal creates the durable successful-handoff authority after both approving
proposal reviews and the full E -> FA -> FD -> FP -> FS sequence.

After verified FS, later viability evidence is post-publication lifecycle work,
not another marker-bearing Gate-0 round. A winner uses the separately approved
emergency dependency revocation followed by the locked upgrade/replacement and
rollback workflow. For a no-winner, FP renders the audited exact successor plan
with pending predecessor FP/FS/seal-digest sentinels; the mandatory
branch-neutral RoundStart is FS's single-parent child and replaces only those
  sentinels in successor/master before any carry audit, observation, P, or legal
  initial D. Verified RoundStart carries the evidence into the next ordinal
  universe; its later RecoveryArchive only authenticates that carry
  retrospectively for clone, Plan-2, and Phase-9 verification. Neither path
  rewrites immutable Done bytes.

### Executable dependency command form

The historical block above preserves the immutable reviewed Round-4 bytes and
uses angle-bracket metavariables. For actual operator use after a winner lock and
the named tools exist, copy this syntactically valid form and replace only the
quoted values with reviewed exact identities:

```powershell
$winnerLock = '<winner-lock>'
$candidateLock = '<exact-candidate-lock-below-.build>'
$transferredCandidateLock = '<exact-transferred-candidate-lock-below-.build>'
$upgradeReview = '<exact-reviewed-upgrade-record-below-.build>'
$transferredUpgradeReview = '<exact-transferred-upgrade-review-below-.build>'
$targetPin = '<exact-target-pin>'
$targetSourceSha256 = '<64-lowercase-hex-target-source-digest>'
$currentLockSha256 = '<64-lowercase-hex-current-lock-digest>'
$candidateLockSha256 = '<approved-candidate-lock-digest>'
$upgradeReviewSha256 = '<approved-upgrade-review-digest>'
$platform = '<x64-or-ARM64>'
$configuration = '<lane>'
$evidenceManifests = @(
    '<x64-Debug-manifest>',
    '<x64-Release-manifest>',
    '<x64-ASan-Debug-manifest>',
    '<ARM64-Debug-manifest>',
    '<ARM64-Release-manifest>'
)

.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock $winnerLock
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile $winnerLock
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile $winnerLock -Offline
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock $winnerLock -RequireCache
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Online -EngineLock $winnerLock -RequireOnline
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile $candidateLock
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile $candidateLock -Offline
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock $candidateLock -RequireCache
.\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Collect -TargetPin $targetPin -TargetSourceSha256 $targetSourceSha256 -ExpectedCurrentLockSha256 $currentLockSha256 -CandidateLockFile $candidateLock -ExpectedCandidateLockSha256 $candidateLockSha256 -UpgradeReviewFile $upgradeReview -ExpectedUpgradeReviewSha256 $upgradeReviewSha256 -Platform $platform -Configuration $configuration -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -UpgradeRoot .\.build\TerminalEngineUpgrade -PassThruManifest
.\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Finalize -TargetPin $targetPin -TargetSourceSha256 $targetSourceSha256 -ExpectedCurrentLockSha256 $currentLockSha256 -CandidateLockFile $transferredCandidateLock -ExpectedCandidateLockSha256 $candidateLockSha256 -UpgradeReviewFile $transferredUpgradeReview -ExpectedUpgradeReviewSha256 $upgradeReviewSha256 -EvidenceManifest $evidenceManifests -UpgradeRoot .\.build\TerminalEngineUpgrade -RequireAllEvidence -PassThruManifest
.\Tools\Build-TerminalEngine.ps1 -LockFile $candidateLock -Platform $platform -Configuration $configuration -Clean -Offline -ExclusiveDisposableRunner
.\Tools\Verify-TerminalEngine.ps1 -LockFile $candidateLock -Platform $platform -Configuration $configuration -RunSmoke -Offline -ExclusiveDisposableRunner
```

### Round 5 evidence-capacity closeout

Round 5 may not accumulate evidence until its RecoveryArchive becomes
unpublishable. Its checked-in mandatory-payload table enforces an exact
prospective RS-JCS count/byte calculation before every history-growing
publication: no more than 192 objects/786,432 bytes of ordinary history, with
64 objects/262,144 bytes reserved for the largest legal terminal path, inside
the immutable 256-object/1-MiB mandatory archive envelope. The round admits at
most eight observations, one P boundary, and one A boundary. A prospective
capacity loss allocates no action/path and takes the canonical
`evidence-budget-exceeded` administrative closeout.

A failed P/A focused suite is immutable and never authorizes a same-round repair
candidate or a mutated frozen universe. Normally it selects the clone-durable
`ProtocolClosed` no-winner branch. If an already-published
`ObservationPending` resolves Approve, or a Superseding fence already exists,
Superseded wins and embeds that failed suite as invalidated; Reject completes
its ledger first and then permits ProtocolClosed. Candidate/protocol repair
moves to the successor ordinal's newly reviewed universe. A failed ordinary final suite gets
one reserved administrative SuccessorD and both fresh final-suite keys. Failure
on that administrative D is explicit WIP/STOP, not an invented engine result.
Any byte repair in that one boundary must be on the pre-result exact
candidate-neutral allowlist, recorded as a before/after path/mode/digest
ProtocolRepair, and approved by two non-authors; gates, expectations,
schemas/corpora, fixtures, evidence, candidate/product/runtime inputs, and
observed results cannot change.
When TransitionRejected, Superseded, ordinary No-winner, or
Qualification-failed is already immutable, the same ProtocolClosure preserves
that terminal branch instead of replacing history. The final decision and
RecoveryArchive embed the closure, every applicable P/A candidate/receipt/
quiet/approval object, and the exact invalidated or retained dispositions.
