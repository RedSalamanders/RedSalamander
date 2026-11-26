# Live Samba Copy and Move selection shapes

Run `20261126T083830Z-52184-e06300c5781d490285f6906a04ee0b14` selected `FileOps_SelectionShapeSmb` without activation. It used test-enabled x64 Release `41211242ce6016c1fba02c339a96055ebe0f05b4` with the byte-verified build receipt `4c2358e6f1001670af6e135b3340a032211d53832bafe63199414e8729132ce8`. Setup, the case and cleanup passed with zero failures and skips.

Each route used independent inputs and wrote to an owned child of the registered Samba fixture. A folder selection and its equivalent individual-file selection each carried eight 4 MiB files and one empty file.

| Route | Folder elapsed | Files elapsed | Peak active transfers for each |
| --- | ---: | ---: | ---: |
| Native Copy | 5,252,129 µs | 6,821,465 µs | 4 |
| Host bridge Copy | 4,958,722 µs | 6,294,953 µs | 4 |
| Managed Move | 4,981,489 µs | 6,187,158 µs | 4 |

All six scenarios verified nine output files, 33,554,432 bytes, exact content and the required source outcome. Each individual-file selection reported determinate progress once discovery closed, with no growth of the frozen total. The runner removed its owned Samba child, and the registered resource manifest still matched its original SHA-256.
