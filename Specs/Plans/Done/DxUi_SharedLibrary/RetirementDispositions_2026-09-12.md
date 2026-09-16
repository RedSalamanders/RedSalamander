# Legacy DxUi retirement dispositions

Supporting evidence for [I19](../DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md).
This records source ownership and test routing; it does not close product runtime, resource,
packaging or native-platform gates.

The retirement removes 27 files under `Common/DxUi` and 28 under `Tests/DxUiTests`, including
the old projects and their eight baseline images. Their prior bytes remain in Git at
`ba72f601` and its ancestors. All production projects already consume the external archive.
The original extraction/case ledger and the September 9 gap audit remain historical evidence.
The canonical library retains its own implementation, control tests and reviewed baseline images.

## Product PowerShell policy dispositions

These are additional consumer tooling checks that still read the legacy tree; they are separate
from the 65 source-policy exclusions already dispositioned in DxUi. Shared implementation spelling
checks are removed from the consumer. Existing runtime tests remain in their actual owner.
No test assertion is weakened to turn a current failing runtime result green.

| Prior policy or group | Current owner and disposition |
| --- | --- |
| DxUiTests unknown/empty CLI switches | The retired consumer executable no longer has a CLI. The standalone parser in DxUi `Tests/Controls/DxUiTests.cpp` still rejects unknown switches and empty suite/perf arguments. These consumer filename/spelling guards are retired. |
| DxUiTests artifact default paths | The retired executable no longer writes product artifacts. DxUi owns its output-root policy; product TestSupport sandbox/environment checks remain. |
| Modal shell quit/last-error implementation | Product modal owner, restoration and delegation checks remain. Library implementation spelling checks are retired; the product credential-prompt modal-quit guard and Commands coverage remain. |
| Accessibility source/build debug toggles | The old projects are removed. The exact public dependency/profile and module provenance tests enforce the current integration; standalone library tests own provider behavior. Product test-enable policy still covers active projects. |
| Shared message-pump and snapshot polling | Repoint to the existing `TestSharedTestSupportPumpsMessagesAndBoundsSnapshotPolling` in ProductUiTests. Viewer and TestSupport assertions remain unchanged. |
| Foreground warning and focus-call allowlist | Retire only the deleted DxUiTests entries. The application, Commands, viewers, TestSupport and ProductUiTests checks remain. |
| Win32 text/edit messages | DxUi runtime `TestNativeTextInputBackendEditMessagesRoundTripWin32Protocol` owns the actual protocol; duplicated consumer source-case spelling checks are retired. |
| Interactive probes/no-activation | DxUi owns its Menu/NativeTextInput capability probes and `TestNoninteractiveWindowActivationBlockerRejectsFocusStealing`. Product activation guards, viewer/Monitor policies and reviewed focus callers remain checked. |
| Observatory Track 9 | Product NavigationView owned-popup activation checks remain. Shared runtime cases remain in DxUi: `TestMenuBarActivationCanReplaceRootSafely`, `TestNativeMenuBarNestedPopupCanDestroyHostSafely`, `TestTreeExpanderReResolvesStableItemAfterSelectionReorder`, `TestTreeSelectionDelegateCanReplaceRootSafely`, `TestTextFieldReplaceSelectionSynchronizesBeforeTerminalNotification`, `TestMaskedTextFieldGeometryMapsUtf16SourceToDisplayElements` and `TestAccessibilityTreeItemProviderKeepsStableIdentityAcrossReorder`. Legacy implementation/diagnostic spelling guards are retired. |
| Observatory Track 10 | Repoint `TestCloudPaginationGuardBoundsProgressAndCancellation` to ProductUiTests; all cloud-provider assertions remain. |
| Observatory Track 12 | Product Terminal/configuration/provider guards remain. DxUi Foundation tests own cached font availability and per-factory/all-factory invalidation; legacy header spelling checks are retired. |
| Observatory Track 14 | Product final-save and stale-suggestion checks remain. DxUi `TestLargeMenuPaintsOnlyVisibleRowsWithCachedOffsets` owns bounded menu work; legacy counter/implementation spelling checks are retired under G5. |

The consumer adds ownership guards rejecting the old directories, solution entries and owned
project inputs that enumerate legacy or external library implementation files. Existing exact-pin,
dirty/missing-source, archive identity, module-sidecar and package-provenance rejection tests remain.

## Tools and documentation

Product source audits remove the six allowlist entries for deleted library/test files. Typography
auditing includes all product `Common` helpers. Source-line reports exclude restored dependencies.
The two legacy project impact-closure rules are removed; the active dependency imports and lock
remain governed. Tool inventory descriptions and the MSBuild example identify current ownership.

Active build, UI, testing and developer guidance point to public headers and the standalone library.
Historical reviews/Done plans retain their original context; obsolete paths there identify historical
source, not a supported checkout location. The I19 top checklist continues to report actual validation.

## Additional native source guard found during Full qualification

At `f7b2bce5`, Compare completes 241 cases with 30 capability skips and one failure:
`search_low_hardening_smoke` still reads `Common/DxUi/DxUi.Controls.cpp`. This is a
mixed product search-hardening case, outside the extracted control-suite ledger.
All product source checks remain. Its checkbox implementation-spelling assertion
is replaced by a public-library interaction in the same case: Left on an unchecked
mixed checkbox and Right on a checked mixed checkbox must clear the glyph and
increase the native host's invalidation counter without changing the Boolean value.
No HWND or foreground activation is required. The replacement is implemented and
`search_low_hardening_smoke` passes in the retained `3a43186a` Release Full
[per-case results](../../../TestRuns/Local-x64/DxUiAdoption/2026-09-13-Release3a43186a/compare_results.json).
That Full run still fails five unrelated ViewerText diagnostic cases; the I19
opening checklist owns final qualification after their correction.

## Upstream changes since the original audit

The comparison from audited RedSalamander `3f0df2f2` to rebased master `0dd0bd6d`
contains four changed legacy files. `DxUi.FrameRuntime.cpp` and `DxUi.h` only add/reorder
Windows include setup; canonical DxUi's shared build properties already define
`WIN32_LEAN_AND_MEAN` and `NOMINMAX`. The canonical six-profile build mapping and local
build receipts cover the selected integration.

The other two files add font-availability refresh/invalidation and its regression. DxUi
`13788e9` implements that API in the public Typography header. Foundation
`TestFontAvailabilityInvalidation` verifies two isolated factories, seeded stale negative
answers, selective invalidation, all-factory invalidation and invalid inputs. It replaces
the newly added legacy `TestDxUiTypographyFontAvailabilityRescansAfterInvalidation`.
This post-audit test is additional to the original 941-case extraction ledger; it has not
been silently counted as one of those inherited cases. No further shared behavior changes
were found in this upstream comparison.
