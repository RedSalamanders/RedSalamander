# Terminal Engine Gate-0 Adapter Contract

This is an evidence-only, candidate-neutral ABI contract. Product integration
may statically link the selected engine inside a private plugin; the Gate-0 DLL
exists only so the same frozen runner can compare candidates without importing
engine-native types.

## Frozen inputs and outputs

`Generated/TerminalEngineFixtures.v1.h` is authoritative for fixture order,
input bytes, schedules, declared action order, action arguments, and the two
checks. Every declared fixture, schedule, and action is mandatory. Returning
`RS_TERMINAL_GATE0_INVALID_ARGUMENT` or another unsupported result for a
declared action fails the row; there is no optional-action capability path.

The runner executes `feed-current-schedule` itself by calling `session_write`
with exactly the frozen schedule chunks. It calls `session_action` once for
each later action in declaration order.

After all actions, `session_state_jcs` returns the candidate's actual result as
one RFC 8785 canonical JSON object. Its exact key set and value types are the
`state` check's `expectedValue` for that fixture. The adapter must derive each
value from candidate state and action observations; it must not include the
generated fixture header, read the expected value, or return a stored expected
hash. The harness hashes returned bytes and compares them with the frozen
expected SHA-256. Missing, extra, duplicate, reordered, or non-canonical JSON
fails.

The adapter does not return `resource-trace`. The harness owns that check from
the reserve/release callbacks, action observations, checked arithmetic,
session destruction, and process-ledger quiet point.

`fixture_id` and `schedule_id` in the session config select the result
serializer and are borrowed until `session_destroy` returns. They do not
authorize reading the oracle.

## Action rules

Every successful non-feed action sets `ACTION_OBSERVED` and
`ACTION_CHECKED_ARITHMETIC`. Limit probes also set
`ACTION_BOUNDARY_ACCEPTED`, `ACTION_PLUS_ONE_REJECTED`,
`ACTION_REJECTION_BEFORE_ALLOCATION`, and `ACTION_NO_PARTIAL_EFFECT`.
Ordering probes set `ACTION_ORDER_PRESERVED`. `teardown-sequence` additionally
sets `ACTION_QUIET_POINT_REACHED`.

| Action | Required candidate-owned observation |
|---|---|
| `encode-paste-disabled` | Encode `payloadHex` with bracketed paste disabled and retain the exact bytes for the fixture state. |
| `encode-paste-enabled` | Encode the same payload with bracketed paste enabled; verify start, payload, and end ordering. |
| `encode-focus-gained-lost` | Encode gained then lost focus in the declared order using the candidate input encoder. |
| `encode-mouse-matrix` | Exercise modes 1000, 1002, 1003 and formats 1005, 1006, 1015, 1016 through the candidate mouse encoder; retain the exact mode coverage and order result. |
| `probe-osc8-limits` | Exercise the 1,024-byte parameter, 8,192-byte URI, and 16,384-byte record boundaries plus one for each applicable bound. The accepted hyperlink remains exact and a rejected record has no partial hyperlink. |
| `probe-kitty-pixel-limit` | Exercise the 16,000,000-pixel boundary through the real Kitty preflight with ledgers sized to admit it, then reject 16,000,001 pixels before decode/allocation. |
| `probe-kitty-replacement-ledger` | Publish an old image, reserve the new compressed/decoded/conversion/publication peak while retaining old storage, inject failure after the new reservation, and prove the old publication and balances remain atomic. |
| `probe-generic-escape-limit` | Admit a 65,536-byte generic escape and discard 65,537 bytes before an effect or unbounded growth. |
| `probe-osc52-limits` | Admit 1,048,576 encoded bytes / 786,432 decoded bytes, reject plus one before decode/allocation, and emit no partial clipboard effect. |
| `probe-pending-effect-ledger` | Admit exactly 2,097,152 pending bytes and 256 descriptors, reject either plus one atomically, then release both balances. |
| `select-logical-range` | Select the declared logical start/end range using the candidate selection API and retain a copied selection snapshot. |
| `resize-sequence` | Resize through 80x24, 20x24, 100x24 and prove logical line, grapheme, wide-cell, and selection stability. |
| `probe-terminal-text-ledger` | Admit the 67,108,864-byte terminal-text boundary, handle plus one before an overflowing allocation, and release the balance. |
| `snapshot-copy-fail-retry` | Inject copy failure after the declared cell count, prove publication did not consume dirty state, then retry a full snapshot with no lost notification. |
| `snapshot-end-mutate-probe` | End the borrow, mutate by resize and write, prove the borrow is invalid, and prove the prior copied snapshot is immutable. |
| `snapshot-static-kitty` | Copy the declared decoded image bytes and placement into an owned immutable snapshot, delete the live image and placement, and prove the prior snapshot and exact decoded RGBA bytes remain stable. |
| `teardown-sequence` | Stop producers, cancel, stop posting, detach callbacks, release, and reach a callback-free quiet point in the declared order. |
| `probe-title-cwd-limits` | Admit the 4,096-byte title and 131,072-byte cwd, reject each plus one without changing the accepted value or emitting a partial callback. |
| `admit-external-events` | Preserve protocol replies ahead of subsequently admitted input, resize, and side-channel events; prove no reordering. |

## Full-measurement actions

Full collection additionally requires four evidence-only action IDs. They are
not fixture actions and therefore do not appear in the 43 conformance rows.
Each returns `ACTION_OBSERVED | ACTION_CHECKED_ARITHMETIC`.

| Action | Timed operation |
|---|---|
| `benchmark-full-snapshot` | Copy the complete immutable consumer snapshot, including cells, styles, colors, cursor, palette, graphemes, hyperlinks, image bytes, placements, and independent generations. The harness uses it once on an empty terminal inside `create-destroy-us`, on the seeded 120x40 cell plus static-image state for `full-snapshot-us`, and once per storm session before retained bytes are sampled. Arguments: `{"copy":"full"}`. |
| `benchmark-dirty-snapshot` | Copy only the dirty publication after the harness collected a full snapshot and then mutated exactly two rows and one cursor state. Arguments: `{"copy":"dirty"}`. |
| `benchmark-resize-reflow` | Reflow the harness-seeded 10,000 deterministic wrapped lines from 120x40 to 80x50 and back to 120x40. Arguments: `{"dimensions":"80x50,120x40"}`. |
| `benchmark-kitty-1024-rgba` | Admit, decode, snapshot, and delete one deterministic 1024x1024 RGBA image while retaining exact reservation telemetry. Arguments: `{"format":"rgba","height":"1024","width":"1024"}`. |

The other timed operations use only `session_create`, `session_write`,
`session_action`, and `session_destroy`. The 8 MiB parse payload is generated
solely from `kPerformancePayload`, must match its frozen SHA-256 before
measurement, and is fed with the repeating
`1,2,3,5,8,13,21,34,55,89`-byte corpus schedule. QPC conversion uses
integer nearest, ties to even. The storm feeds 1 MiB to each of eight sessions
in round-robin 4,096-byte chunks, publishes one complete snapshot per session,
samples CPU-only retained ownership before teardown, and then destroys every
session. The eleven metric IDs and sample schedules are:

- `create-destroy-us` (10 warmups, 50 samples);
- `parse-8mib-us` (5, 30);
- `full-snapshot-us` (10, 50);
- `dirty-snapshot-us` (10, 50);
- `resize-reflow-us` (5, 30);
- `kitty-1024-rgba-us` (3, 20);
- `eight-session-storm-us` (3, 20);
- `eight-session-retained-bytes` (1 warmup, 5 samples);
- `added-release-package-bytes` (one value);
- `crossed-abi-elements` (one value);
- `non-system-runtime-dependencies` (one value).

The package value comes only from
`terminal-engine-contribution-zip-v1` in the frozen measurement protocol. The
provider creates the normalized adapter and complete notice entries twice with
the fixed order, timestamp, attributes, and `ZipArchive` compression, requires
byte-identical archives, and subtracts the same recipe's empty-ZIP byte length.
Raw DLL plus notice sizes are not a valid substitute. The normalized adapter
entry is evidence-only and does not change the winner's eventual
`source-static` or `private-dynamic` product topology.

## State-only fixture observations

Fixtures with no non-feed action still require direct candidate-state
inspection. This includes complete cell/style/color/cursor/palette copies,
malformed UTF-8 replacement, synchronized-output publication, Kitty
image/placement bytes and independent generations, z ordering/crop/delete,
and immutable snapshot ownership. The exact result object remains the frozen
`state` check; the adapter must not substitute a generic pass boolean.

## Resource and lifecycle rules

All candidate and adapter terminal-text, image CPU, image GPU, scratch,
conversion, and pending-effect ownership crosses the host callbacks with the
correct kind. A replacement reserves new peak ownership before releasing old
ownership. Rejected sizes are checked before callbacks or candidate allocation.
Callback arithmetic is exact and callbacks may be concurrent.

`session_destroy` stops producers and callbacks before returning, releases all
session reservations, and leaves no adapter thread that can execute DLL code.
The harness closes the session ledger, verifies process balances, and unloads
the DLL only after every session reaches that quiet point.

## ABI layout identity and public robustness seam

`TerminalEngineGate0Probe` is compiled by the same lane MSVC toolchain as the
candidate adapter. Its `layout` mode serializes one RFC 8785 canonical record
containing the architecture, C calling convention, pointer size, both public
ABI constants, every value of every public enum, every public function-pointer
type's size/alignment, and every public structure's size/alignment plus each
field's offset/size/alignment. The record identity is exactly
`layout-v1-<sha256-of-layout-JCS>`.

The adapter translation unit includes `TerminalEngineGate0Layout.h` and exports
`rs_terminal_gate0_query_layout_jcs`. That versioned bootstrap export derives
the complete JCS from the enum values, function-pointer types, `sizeof`,
`alignof`, and `offsetof` expressions compiled in the adapter translation unit.
The probe independently derives the same JCS in its own translation unit,
requires exact byte equality, and verifies the common SHA-256 before it calls
`rs_terminal_gate0_query_adapter`. An injected or echoed layout label, a
source-header hash alone, or a handwritten build label is not layout proof.

`identity.build_identity` is a separate
`build-v1-<sha256-of-build-JCS>` receipt identity. The collector recomputes
that build JCS from the frozen candidate, source archive, descriptor, patch,
result tree, lane, host ABI headers, resolved path bindings, actual toolchain
receipts, declared artifacts, and final artifact hashes. Candidate-authored
code cannot supply or replace this proof.

Every future executable round freezes exactly one candidate-neutral
`host-wil-headers` dependency archive per platform input set. Its consumers
are the exact executable cohort, never a candidate-specific subset. The
provider verifies the archive digest, removes only the exact contained prior
expansion, performs a fresh bounded safe extraction in every consuming phase,
and requires the exact member `include/wil/resource.h`. An existing expanded
tree is never reused or self-attested. The resulting `include` tree identity
is the canonical sorted file-record JCS digest and is part of the build
receipt.

The harness and probe receive only that explicit include root and its expected
tree digest. Candidate-driver isolation is bracketed by provider revalidation,
and both neutral wrappers revalidate the complete tree immediately before and
after MSBuild. A persistent mutation by any child fails the lane before
provider-owned probes or evidence publication. Both projects set
`RSUseVcpkgManifest=false`, `VcpkgEnabled=false`, and
`VcpkgEnableManifest=false`, require `HostHeaderIncludeRoot`, and put it first
in their include search. Ambient repository `.build/vcpkg_installed` headers,
developer package caches, and inherited manifest installs cannot satisfy a
missing or changed explicit tree. Round 3 admits no executable candidate, so
it does not fabricate this input; a later round must freeze it before
executable Bootstrap or Collect can proceed.

The descriptor's closed `build.kind` selects exactly one provider path:

- `provider-generic-cmake-msvc-v1`: provider-owned CMake configure/build
  orchestration with the descriptor's sorted targets;
- `provider-generic-zig-msvc-v1`: provider-owned Zig/MSVC orchestration with
  the descriptor's sorted steps;
- `provider-generic-cargo-msvc-v1`: provider-owned Cargo/MSVC orchestration
  with the descriptor's package and sorted feature set;
- `candidate-wrapper-v1`: the only kind allowed to name a candidate-owned
  driver, mechanically deriving `maintainedBuildWrapper=true`.

All provider-generic kinds forbid a driver and mechanically derive
`maintainedBuildWrapper=false`. The assessment, patch proof, descriptor
reader, and provider must agree with that derived Boolean before build work.

A candidate wrapper is one patched `.ps1` file and is invoked with exactly
four argument tokens:
`-InputPath <absolute-path> -ResultPath <absolute-path>`. The provider writes
the input with create-new semantics below the clean lane build root. It is
nonempty, at most 32,768 bytes, BOM-less UTF-8, exact RFC 8785 JCS, and must
validate against `TerminalEngineCollectionDriverInput.schema.json`. Its
closed top-level shape is `{formatId,version,candidate,lane,paths,pathBindings,
artifacts,gate0}`, where:

- `candidate` binds candidate/pin/source/patch/result-tree/descriptor
  identities and `buildKind="candidate-wrapper-v1"`;
- `paths` contains only the provider-authorized source, build, notice,
  tracked ABI header/layout, verified host-header include root, MSBuild,
  build-receipt, and source-dependency paths;
- `pathBindings` is the exact sorted resolution of every declarative
  descriptor binding, including its value kind, absolute path, and identity;
- `artifacts` contains only the descriptor's adapter DLL leaf and complete
  private-runtime leaf closure;
- `gate0` binds the tracked ABI/layout hashes, complete host-header tree hash,
  and decimal capability mask.

The provider validates that exact input itself before isolation. The wrapper
must not infer another path, mutate an input, publish evidence, or add an
artifact. It may only configure and build within the authorized roots.

`ResultPath` is exactly `candidate-driver-result.json` below the same lane
build root. The wrapper publishes it atomically from a create-new temporary
file. It is nonempty, at most 4,096 bytes, BOM-less RFC 8785 JCS, and validates
against `TerminalEngineCollectionDriverResult.schema.json`. Success is
exactly `{formatId="red-salamander-terminal-engine-wrapper-result",version=1,
status="pass",failure=null}`. The only failures are the schema's closed
configure/build phase-and-code pairs. No wrapper-provided identity, gate,
score, measurement, native, toolchain, ABI, legal, artifact, or C7 claim is
accepted.

For a validated configure/build failure, the provider writes a bounded
create-new reproducer, binds the input/result and captured stream digests, and
emits a G1 build failure. G2 through G12 are not run. A missing, oversized,
noncanonical, late, repeated, or schema-invalid result fails closed.

Artifact discovery comes exclusively from the descriptor. The adapter and
private-runtime leaves are unique under ordinal-ignore-case comparison and
must each resolve to exactly one regular build output. CMake static-library
records additionally bind their frozen source input, project leaf, and library
leaf. Unknown, missing, duplicate, case-colliding, undeclared, or extra
artifacts fail before ABI or performance evidence.

The provider independently verifies the selected build-kind receipts, MSVC
lane identity, CRT/LTCG/reproducibility policy where applicable, source
dependency trace, PE architecture/import closure, declared private-runtime
closure, and final artifact hashes. It reruns layout and fuzz probes under
runner-enforced network isolation and folds those independent results into the
Gate-0 evidence. Generic build orchestration and wrapper orchestration converge
on the same provider-owned post-build verification; a build-process success
claim is never artifact proof.

Before fuzzing, the probe also requires exact candidate ID, upstream pin,
source SHA-256, adapter-header SHA-256, build identity, architecture, and the
complete required capability bitset. A valid layout export from the wrong
adapter therefore cannot produce Gate-0 evidence.

The frozen C ABI's `session_write` entry point is also the public parser
robustness seam. Probe `fuzz` mode stays local and bounded: it feeds fixed
empty, malformed UTF-8/VT, contiguous, and split cases plus a deterministic
seeded arbitrary-byte corpus. It executes every baseline twice and compares
exact result, state, and host resource-callback trace digests. It then combines
the same bounded corpus and repeats every callback allocation-denial ordinal
observed in the successful baseline, up to the hard runner cap. Every run must
balance the callback ledger, destroy a created session, reach the callback
quiet point, and unload the DLL before the next run.

The fuzz record contains only counts, identities, and SHA-256 digests. Input
bytes, terminal state, candidate paths, and callback payload content are never
emitted. A crash, timeout, nondeterministic trace, unvisited denial ordinal,
unbounded state, resource imbalance, destroy failure, or module remaining
loaded fails the record.

## C7 score-proof extension

The separate candidate-neutral shell-semantics score proof is defined by
`TerminalEngineC7ProofContract.md` and
`Specs/Terminal/TerminalEngineC7ProofContract.v1.json`. It is not part of the
29-fixture/43-row mandatory corpus and cannot weaken or replace a row. A
candidate seeking C7 rating 5 must implement the exact evidence-only
`probe-c7-engine-effects` action. The provider independently executes that
action and the existing `admit-external-events` ordering action through the
built adapter DLL in every required lane; driver-supplied proof facts are not
accepted.
