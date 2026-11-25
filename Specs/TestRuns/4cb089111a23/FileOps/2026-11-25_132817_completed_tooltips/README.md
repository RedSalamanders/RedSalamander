# I66 Completed-group tooltip check

Hosted x64 Release [run 45588264577](https://github.com/RedSalamanders/RedSalamander/actions/runs/45588264577)
found one painted badge and six tooltip controls while problem cards remained visible. The assertion
counted every popup tooltip. The corrected witness counts actual Completed-group badge controls
and checks their localized tooltip text, accessible name and help text. The first-frame disclosure,
problem cards and retirement of the clean task accessibility controls remain covered.

Full x64 Release `1a7e7cbc5` built with zero warnings and errors. Receipt
`054384ca616ad085e0ef17e1a73520044b7f817e4d584a4f6c7f5dc493a64b56`.

| Exact case | Run | Result |
| --- | --- | --- |
| `cmd_pane_fileops_completed_group_and_navigation` | `20261125T122741Z-75188-fa83689ed40c4dda9f8c82438116ca4c` | 1 passed, 0 failed, 0 skipped |
| `FileOps_ClearCompletedScope` | `20261125T122444Z-25040-c6dc42811b554f9fa959e80ae3462fe0` | 3 passed, 0 failed, 0 skipped |
| `FileOps_ProviderCapabilityMatrix` | `20261125T122540Z-25040-95faf27977f14b998685544b7556c9ec` | 3 passed, 0 failed, 0 skipped |

The approved Commands check measured 1.095 seconds for its native case. The first collapsed
frame reported one painted badge and one badge tooltip. The two File Operations cases
ran without activation and include setup and cleanup. Hosted native qualification remains open.
