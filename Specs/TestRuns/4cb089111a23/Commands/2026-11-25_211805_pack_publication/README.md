# I66 transactional Pack publication

Stored ZIP and 7z encode to a private sibling held by the shared Local file transaction. The test injects a failed commit for each format. It checks that the previous archive bytes and original sources remain, the stage is removed by its handle, and a later overwrite produces a readable archive.

Both full test-enabled x64 builds at `d54ad80eb1c6173e68dd8957b5715bdcbe4b46b6` had zero warnings and errors. The focused `cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation` selection ran without desktop activation.

| Profile | Verified receipt | Governed run | Passed | Failed | Skipped |
| --- | --- | --- | ---: | ---: | ---: |
| x64 Debug | `827315dc6bd10024b3d574e49a1e32271e6eeaab4a82a87738daf35e4bf99b31` | `20261125T200337Z-44120-3d0833aa0007421fb2b799d2e3dc1a0b` | 1 | 0 | 0 |
| x64 Release | `119e90f82ca45bfad5ce2ecb9de9464ea4d32f7e7d628d2b6e965700b482e3ea` | `20261125T201722Z-57800-5b081c1e21fe4ae497559ccdfe82dea1` | 1 | 0 | 0 |

For the first stored ZIP with 16 entries and 23 source bytes, `archive.pack_us` measured 6,242 µs in baseline x64 Debug run `20261125T185016Z-37604-8000aef41f024e65870c2d6df59b7de3`. The candidate measured 8,791 µs in Debug and 4,499 µs in Release. The new `archive.pack.publish_us` recorded 2,000 µs in Debug and 1,707 µs in Release for the same ZIP. These individual fixture observations include the flush and publish boundary. Worker execution and full repository coverage remain open.
