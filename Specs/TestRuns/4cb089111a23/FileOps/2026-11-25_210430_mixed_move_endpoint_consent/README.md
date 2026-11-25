# I66 mixed Move confirmation and endpoint recheck

| Profile and source | Verified receipt | Governed run | Result |
| --- | --- | --- | --- |
| x64 Debug `54d9fc69a` | `a658d8c7ea2bac6f7b3359a282efd6bed91e336b2b56d453c06b945922decb41` | `20261125T193637Z-86976-585b1df3b0ef422698098a3c4c4fca98` | Mixed Move and Copy-only Move, 3 passed |
| x64 Release `387e477db` | `42cc7ea4fe3bcd4c11b429c75185b4ebb80005e0d85e80e7344b865c2306eda1` | `20261125T200336Z-69136-a1e66005bf7d4aa996fb7e9476b50ea4` | Mixed Move, changed root and Copy-only Move, 3 passed |
| x64 Release `387e477db` | Same Release receipt | `20261125T200449Z-69136-9c8a2703b4134c7b98346386c89f59b9` | Deferred exit and Create Folder, 3 passed |

Both test-enabled builds produced zero warnings and errors. Each focused selection used no desktop activation and reported zero failures and skips. Repository coverage was `NOT_EVALUATED`.

The owned two-root fixture verified one Continue/Cancel question per request. Cancel retained both inputs and left the destination empty. Stable acceptance published both exact outputs, moved the native item and retained the Copy-only original. Acquisition of stronger writer proof preserved the accepted Copy-only plan.

The held preparation control changed one captured root while keeping its files intact. The worker rejected the whole selection with `0x800703EE` before publication, and the task preserved that exact status. Both originals stayed intact. The successful Release run measured the two-source route check at 29 µs for the stable fixture and 163 µs for the refused one. These are individual fixture observations.

Preserved controls include the original missing native mutation receipt in `20261125T192030Z-59472-2aea1d26039d4393ae20de491c131cf7` and the first endpoint check in `20261125T194820Z-82772-b7a37c58b4814e59954306338983d2f5`. The latter stopped preparation with the correct reason, then lost that reason in its aggregate result.
