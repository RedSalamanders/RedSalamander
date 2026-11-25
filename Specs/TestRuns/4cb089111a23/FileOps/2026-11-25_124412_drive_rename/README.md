# I66 Google Drive Rename collision refusal

The fixture submits singular and batch Rename onto an occupied folder and a singular Rename to
the raw name of two decorated siblings. It also performs case-only and ordinary free-name
Rename of a folder and checks its child.

Both x64 Debug builds used the same PluginContractTests project union with zero warnings and
errors. The verifier checked the provider DLLs and native executable in the selected manifest.
Both runs used the owned loopback endpoint without desktop activation.

| Source and run | Collision PATCH / DELETE | All GET / PATCH / DELETE | Google Drive checks |
| --- | ---: | ---: | --- |
| `83c53b21b` · `20261125T113710Z-32912-629f9377104c4b6b976f011766582628` | 4 / 2 | 14 / 6 / 2 | 42 passed, 4 failed |
| `6bdc5816f` · `20261125T114412Z-72396-7197ac9e95dd42ba8df174d58a7fdc9c` | 0 / 0 | 10 / 2 / 0 | 46 passed, 0 failed |

The old route entered recursive folder Move and moved or removed fixture entries. It also
accepted the duplicate raw name. The corrected route refused all three occupied destinations
and preserved the source folder, batch source, other sibling and leaf. It published each free
name with one PATCH and kept the subtree. The standalone plugin suite passed on the correction.

The receipt IDs are
`226f378084880df5cc63425323d20c1f10d13411dd0d2028ce6e7f5abb518032` and
`be6ad3335ed0149d3edb14fac185ea3218c189b635bf8d541b3bdd76df704475`.
The fixture times were 34,646 and 24,749 microseconds. These single loopback observations
mainly compare different mutation shapes. Live Drive qualification remains open.
