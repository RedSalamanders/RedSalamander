# Terminal-engine declarative collection descriptor contract

The collection provider is candidate-neutral and owns the only supported build
orchestration path. A collectable candidate adds exactly one mandatory path:
`red-salamander/collection-descriptor.v1.json` in its exact verified result
tree. The descriptor is declarative JCS data. Candidate executable drivers,
configure/build scripts, command fragments, environment edits, hooks, compiler
paths, linker paths, toolchain files, launchers, and post-build commands are
not part of this interface and are rejected by its closed schema.

The provider and independent patch proof use
`TerminalEngineCollectionDescriptor.psm1`. They accept at most 65,536 exact
RFC 8785 JCS bytes satisfying
`Specs/Terminal/TerminalEngineCollectionDescriptor.schema.json`. The
descriptor path must be one exact changed path in the no-renames result-tree
diff. Every relative path is slash-normalized, bounded, contained, free of
dot/dot-dot, ADS/device aliases, trailing dot/space, reparse traversal, and
Windows case-fold collisions.

## Provider-owned build

The provider-generic build union contains
`provider-generic-cmake-msvc-v1`, `provider-generic-zig-msvc-v1`, and
`provider-generic-cargo-msvc-v1`. For each kind the provider resolves the
toolchain, owns a fixed configure/build/test argument order, uses only the
verified result tree and a fresh lane directory, applies the lane target and
configuration, and accepts only bounded sorted target, step, package, and
feature names. No generic kind executes a candidate-authored build wrapper.

`pathBindings` are typed path capabilities. They may bind an `RS_GATE0_*`
key only to an exact frozen input file, an expanded declared header package,
or one closed provider path. They cannot carry literal values,
commands, environment variables, executable/tool paths, or arbitrary
filesystem paths. The provider maps them through the fixed mechanism for the
selected generic kind. Adapter, static-library, and private-runtime leaves are
data for provider discovery and proof; they are not candidate commands.

The closed union also retains `candidate-wrapper-v1`. That kind alone requires
an exact patched `driverRelativePath` and the separately frozen wrapper
input/result ABI. The provider mechanically derives
`maintainedBuildWrapper=false` for all three provider-generic kinds and
`maintainedBuildWrapper=true` for `candidate-wrapper-v1`, then requires the
assessment build ledger to match. A renamed or provider-hidden wrapper cannot
improve C4.

## Header and extraction inputs

`headerPackages` references exact frozen input IDs. The generic
`nuget-header-package` recipe applies provider-fixed pre-extraction limits for
compressed bytes, entry count, per-entry and total expanded bytes, expansion
ratio, path length and depth, encryption, compression methods, entry types,
duplicate/case-fold names, device aliases, ADS, links, and trailing dot/space.
The candidate cannot weaken any limit.

## Inline legal closure

`legal.dependencies` is the complete inline notice manifest metadata. Each
dependency has a stable ID, SPDX expression, and at least one license input.
An input is exact bytes from the candidate result tree, an exact member of a
frozen archive, or an exact whole frozen input. The provider derives pins,
versions, URLs, source SHA-256 values, and retrieval timestamps from frozen
candidate/bootstrap records; the descriptor cannot replace that provenance.

The provider copies verified legal bytes to a fresh materialization root,
synthesizes a transient canonical manifest for the unchanged notice formatter,
runs materialization twice into distinct fresh roots, and requires byte-equal
manifest, notice, receipt, and dependency identities. The descriptor legal
dependency IDs must equal the independently derived source/build/header/private
import closure. Self-declaring a smaller legal set is not evidence.

The generated notice output must be absent, untracked, contained, and
non-reparse before materialization. It is provider-generated and is not a
candidate patch path.

The descriptor never carries fixture inputs, expected terminal state, gate
ratings, score anchors, performance expectations, arbitrary commands, or tool
paths. Those remain provider-owned frozen contracts.
