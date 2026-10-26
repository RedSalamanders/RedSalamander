# ViewerSpace nonactivating desktop observation repair

Scope: test attribution and fixture readiness on branch `ej/viewerspace-desktop-preservation`,
based on `5141305d9c1d284c031d325d2e448fcd4a13a1d9`. Production ViewerSpace bytes are unchanged.
This is focused x64 Debug validation, not a new PR/Fresh Full or interactive qualification.

The old UIA test compared the entire desktop's foreground HWND and physical cursor
coordinates across a roughly two-second scenario. Human input or another application's
foreground change could fail it without ViewerSpace having activated anything. The
original failed PR did not record which value changed, so its physical event cannot
be attributed retrospectively. Source inspection and the retained DLL import audit
found no physical cursor/input writer on the reviewed ViewerSpace/DxUi UIA path.

The shared observer now requires the existing activation guard, samples read-only
desktop state and retains process-filtered foreground WinEvents, including transient
owned activation. Missing guard, hook, capture or bounded callback-drain evidence still
fails. Cursor/foreign-foreground/last-input changes are diagnostic; they never excuse
owned activation. Its hidden message-only name event drains the ordered callback
stream within 250 ms. It never emits a fake foreground event, moves/restores the
pointer, injects input or restores another application's focus. Source guards cover
those bans; simulated observations do not qualify physical-input behavior.

The broader first candidate exposed a second fixture race: the held provider read
was entered while the UI still showed Queued and four pending updates. It advanced
to Scanning/zero updates during capture; the unchanged stability check correctly
failed. The capture now waits for Scanning and an empty update queue while the
provider remains gated. New native negative controls reject both unsettled states.

| Evidence | Result and interpretation |
|---|---|
| Original merged PR, `original/` | FAILED: 949 passed, 1 failed. The two original isolated repeats passed; neither replaces the failed record. |
| `fail-first-run.txt` | Expected exit 1. Read-only injected cursor changes reproduce the old exact foreground/cursor assertion. No physical input was generated. |
| `candidate-focused-results.json` | Four focused cases passed before the refresh-readiness correction. These identify that earlier binary and remain separate. |
| `pre-settle-candidate-noninteractive*` | FAILED first candidate group. Desktop checks passed; refresh capture changed Queued -> Scanning with the provider still gated. Its invalid provenance is retained. |
| `settled-candidate-noninteractive*` | PASSED: 29/29 fresh-process noninteractive cases, 0 failed, 0 skipped; 78,846 ms. Includes ambient and owned-activation controls, original UIA case, lifecycle captures and unrelated viewer coverage. |
| `settled-candidate-build.txt` | Targeted ViewerPETests x64 Debug build: 0 warnings, 0 errors. |
| `settled-source-guards*` | 204 passed, 0 failed, 0 skipped, including all three new desktop source-contract checks. |
| `scope-inventory.json` | 118 native files, workflow digest matches. Explain remains exit 2 with five interactive obligations pending. |
| `spec-inventory*`, `tool-inventory-status.json`, `archive-*` | Contract/inventory checks; see their explicit outcomes. These do not launch or qualify product runtime. |

Final build receipt: `7f2bb0d411e03846560bfab1ce96076fa6e82ebc7e1722a8b321f3e7ecca42b4`.
The receipt attests the targeted test executable, not a complete solution/runtime closure.
Its executable SHA-256 is `E3603E752DABE9729B0704C97D75A0143BE49977514A84D5AEAFFC03EB1715F1`.
The unchanged ViewerSpace DLL SHA-256 is `04C185D078886429299DA0605EAB7DE0412AC4C6C35A9A26C47CE7D619CA41DD`.
The final run records matching pre/post test-source hashes and executable hashes.
Docs, ledger, archive and plan closeout edits followed the build; they are not claimed
to have the receipt's complete workspace snapshot identity. Earlier receipts and
their failures retain their original identities. The transient fail-first source
tree was not serialized; its observed build identity/assertion trace is diagnostic,
while the current ambient regression fixture is retained in the source tree.

Final run: `20261016T125146Z-56600-d18f09e2a9d84e359d47ee3f35d5d4fd`,
`ViewerPETests.exe --group=noninteractive`, marked `Z:\RedSalamander.Perf` sandbox.
The normal native child deadline remains 120 seconds. The contained outer group
bound is 600 seconds; no native deadline or assertion was relaxed.
The exact driver is retained as text; it uses the public profile lock, sanitized
environment and owned-process containment helpers. Lifecycle fixture manifests and
binary hashes accompany selected capture provenance. Raw PNGs remain in the named
sandbox run; they have not been promoted as visually reviewed documentation images.

The final real UIA case observed five cursor changes and five input changes, with
zero owned transitions. The ambient case observed ten cursor/foreground/input
changes, also with zero owned transitions. Final drain samples were 85/116/64 us
for original UIA/ambient/lifecycle cases. Those are individual diagnostic timings,
not percentiles, a speedup claim or production performance qualification.

All validation stayed nonactivating. Real focus/input scenarios, Release, ASan,
ARM64 and whole-repository Full qualification are outside this bounded result.
The unchanged production DLL and passing native group do not prove which actor
changed the desktop in the earlier PR failure.

Raw archives retain exact bytes via the owning `.gitattributes` pattern; the file
manifest records SHA-256 and size for each curated payload except itself and the
archive validation outputs generated afterward. No fixture directory was copied
recursively.
