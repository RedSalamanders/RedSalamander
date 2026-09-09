# Terminal Engine C7 Score-Proof Contract

This contract is a candidate-neutral score-proof correction frozen before
performance collection. It does not add, remove, replace, or make optional any
fixture in `TerminalEngineFixtureCorpus.v1.json`. Every collection or
performance record created before this correction is stale and must be
discarded or quarantined; all three required lanes must restart from Bootstrap.

`Specs/Terminal/TerminalEngineC7ProofContract.v1.json` and its closed schema are
authoritative. `TerminalEngineC7Proof.psm1` constructs and validates the neutral
authentication state machine, exact engine state, bounded provider proof, and
manifest descriptor. The candidate patch implements only the public-engine
observation action. Authentication, nonce, epoch, replay, framing, and
authorization policy never enter a candidate patch.

## Candidate action

The provider creates a Gate-0 session with
`fixture_id=c7-engine-effects-v1`, feeds the exact 192 bytes returned by
`Get-RSTerminalEngineC7FixedStreamBytes`, invokes
`probe-c7-engine-effects` with canonical arguments
`{"contractId":"c7-engine-effects-v1"}`, and copies `session_state_jcs`.
It repeats that work in fresh sessions under `all-at-once`, `bytewise`, and
`terminator-split`. All three state strings must be byte-identical canonical
JSON equal to `Get-RSTerminalEngineC7ExpectedStateJcs` and at most 8,192 bytes.

The fixed stream uses only synthetic values. It drives OSC 0 title, OSC 7 cwd,
OSC 133 A/B/C/D/A/B with decoded command `echo rs-c7` and exit code 23, and one
OSC 8 hyperlink cell. The action must derive its result only from public typed
shell-integration callbacks, `livePromptSpan`, and owning copies of cwd, title,
and hyperlink state. It writes 1049h through the real engine, samples
`livePromptSpan` while the alternate screen owns the page, then writes 1049l.
It must not scan terminal cells or PTY text for semantic events.

The candidate state contains only event kinds, ordinals, exit code, prompt-span
coordinates/status, and SHA-256 identities of the three synthetic copied
values. It never contains terminal text, a path, title, URI, nonce, or payload.

## Provider-owned execution

The collection provider, not the candidate driver, owns the score proof:

1. After the normal 43-row mandatory harness passes, the provider launches a
   provider-built neutral C7 probe in the same network-isolated lane.
2. The probe loads the exact adapter DLL itself and validates the same
   candidate ID, pin, source archive, build identity, public-layout identity,
   header identity, and capability mask as the provider fuzz probe.
3. It rejects a pre-created output, uses fresh sessions, applies the exact
   three schedules, invokes the real action, copies bounded state, destroys each
   session, proves callback/resource quiet, and unloads the DLL.
4. In a separate fresh session it reruns the frozen
   `write-response-order/terminator-split` input and
   `admit-external-events` action from the generated fixture contract. It
   returns the actual canonical state plus the actual observation flags
   `action-observed`, `checked-arithmetic`, and `order-preserved`.
5. Probe output is closed and bounded. It may contain the synthetic
   content-free state strings only; it may not contain input bytes, terminal
   contents, paths, titles, URIs, nonce, payload, or environment values.
6. The provider passes those actual probe results to
   `Add-RSTerminalEngineC7ProviderEvidence`. Candidate-driver evidence objects
   named `score.c7-proof` are rejected.

The exact attachment call is:

```powershell
$proof = Add-RSTerminalEngineC7ProviderEvidence `
    -CandidateResult $driverResult.Candidate `
    -RepoRoot $Context.RepoRoot `
    -SourceArchiveSha256 ([string]$Candidate['sourceArchiveSha256']) `
    -Platform $Context.Platform `
    -Configuration $Context.Configuration `
    -AdapterIdentitySha256 $gate0Proof.BuildProof.BinaryProof.AdapterSha256 `
    -StateJcsBySchedule $c7Probe.StateJcsBySchedule `
    -OrderingStateJcs $c7Probe.OrderingStateJcs `
    -OrderingActionFlags $c7Probe.OrderingActionFlags
```

The provider may attach this object to every complete candidate, but it must
execute it for every candidate whose frozen `score.c7` anchor is `c7-r5`.
Failure cannot be converted to a claimed descriptor or accepted driver hash.
The provider proof record is RFC 8785 canonical, at most 32,768 bytes, and the
manifest receives only its SHA-256, byte count, and `verified` status.

## Authentication composition

The neutral harness parses actual canonical frame bytes with closed fields:
protocol version, 128-bit nonce, epoch, unsigned 64-bit sequence, event kind,
operation ID, and bounded payload. It validates origin metadata before parsing
authority, consumes sequence only for valid authenticated frames, rejects
duplicate operations, and invalidates the epoch on authentication, framing,
bounds, replay, duplicate, gap, order, or wrap failure.

`IdleAtPrimaryPrompt` authorizes only when a valid activity frame's engine
ordinal equals a typed `input-start` observation and the same observation says
live prompt span available, primary screen, not running, not alternate, and not
nested. A valid authenticated frame with a non-authorizing engine observation
consumes sequence but leaves the epoch active. Ordinary OSC 7/133, remote, and
nested terminal output never authorize and never repair an invalid epoch.

The frozen matrix contains 31 positive/negative composition cases plus three
boundary cases: an actually valid 2,097,152-byte frame passes, the same shape at
2,097,153 bytes invalidates, and a valid frame at the active epoch's
`UINT64_MAX` next sequence invalidates rather than wrapping.

## Evaluator rule

Static assessment facts remain bound by `score.c7`. For `c7-r5`, finalization
also requires exactly one provider-owned `score.c7-proof` with
`resultCategory=verified` in ARM64 Release, x64 ASan Debug, and x64 Release.
The three proof hashes must be distinct because lane identity is inside the
proof. Missing, duplicate, replayed, malformed, zero, oversized, or non-verified
descriptors reject finalization. Lower C7 anchors do not receive r5 credit from
this proof.
