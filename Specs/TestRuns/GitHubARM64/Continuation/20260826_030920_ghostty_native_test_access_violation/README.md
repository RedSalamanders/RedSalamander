# Ghostty native ARM64 C-ABI test access violation

## Reproduction

- Workflow run: `32923375705`
- Source commit: `3f1bab6b514ba3b1b68e267fbfcb82783930a771`
- Runner image: `windows-11-vs2026-arm64` version `20260817.130.1`
- Candidate commit: `9c3ec931d64561a8407dde7ac984ce156ae91539`
- Candidate tree: `359872ab74fe1ecc96428133c0546be3cb969fdd`
- Corrective lock SHA-256: `d1bad7bd545164b94f4cf6d10c3143fa5e25db6d0bd62ffa6aa36fecfbcc9078`
- Corrective overlay SHA-256: `06847f6594427cc584cef0f0ec66a54080f236b2414fe40e61b381accaf38264`
- Fixed Zig test seed: `0x97d35bc7`

The workflow executed the unchanged upstream target graph, serialized with
`-j1`, against `aarch64-windows-msvc` Debug with SIMD disabled. The ARM64 ABI
schema validated 153 types and the Zig suite reported 2,904 passed, 50 skipped,
and 0 failed. The sibling C-ABI test executable returned process code 5 before
test output:

```text
.zig-cache/o/7cbd97bf612ffbdefcd2c6b860f63d82/test.exe
```

The workflow authenticated the complete bounded failure shape, waited two
seconds, hashed the exact cache-contained executable, and retried that same
image once with the same seed. The retry exited `-1073741819`, Windows status
`0xC0000005`, before emitting a test summary or named test failure.

## Retained reproduction

Workflow run `32925672178`, from source commit
`53c1efe080c2228eda149ffec7b8fd9129ca7e35`, reproduced the same access
violation and uploaded `ghostty-native-arm64-crash-32925672178`. The downloaded
artifact was internally verified and contains the initial output, empty retry
output, hashes, build metadata, and the exact 55,152,640-byte `test.exe` with
SHA-256
`58885dc091476573ddf69b1843f14c955857436f946121291ddc5f8fc26bbba9`.
WER did not materialize a dump or matching event record.

## Root cause

The private Windows shared-library overlay sets `strip=true` on `mod.vt_c` in
`src/build/GhosttyLibVt.zig`. Upstream `build.zig` later reuses that same module
when linking the Debug C-ABI test, so a shipping-only size policy leaks into the
native test executable. The retained image has no PDB or usable unwind data and
crashes before its first test summary.

The candidate overlay cannot clone the test module: changing candidate
`build.zig` perturbs the production ARM64 graph and violates the frozen runtime
identity. The qualification workflow therefore creates a temporary test-only
clone after the production candidate build, sets `strip=false`, retains frame
pointers, enables synchronous unwind tables, and restores `build.zig` during
cleanup. A local ARM64 cross-build produced a 66,208,768-byte C-ABI test image
with its PDB under this configuration. The reviewed overlay and both production
identities remain byte-for-byte unchanged.

P6.5 remains unchecked until the repaired workflow passes all three upstream
targets on native ARM64 without using the bounded code-5 retry. No crash waiver
or release qualification is claimed from this archive.
