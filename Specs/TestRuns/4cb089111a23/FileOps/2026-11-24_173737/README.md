# I66 Samba corroborated F2 Rename

The x64 Debug non-activating run `20261124T173007Z-100696-ce72a4511ad04f9e9076f0aa92525ad2`
passed setup, `FileOps_SmbCorroboratedRenameProvider` and cleanup with zero skips. The selected
application build emitted zero warnings and zero errors. Its build receipt was
`87b31c051d8ac0579db2157090f5145911fae3673051440b352088bf29e9f273`.

The Samba fixture has no exact file ID. The case verified completed folder Rename, case-only
spelling and independent occupied-target preservation with Keep both. Cancel while paused
preserved the source and destination state. Editing a paused source made the route refuse it
before dispatch. The provider loop verified 100 separate same-parent mutations and their
source and destination state. Cleanup removed all owned fixtures. The representative
single-folder capture and mutation each held five directory pins and released them before
the next stage. Directory pin count does not measure whole-process handle peak.

| Measurement | Duration (µs) | Items |
|---|---:|---:|
| Folder Rename through the host | 285218 | 1 |
| Cancel before dispatch | 47185 | 1 |
| Changed source refusal | 144138 | 1 |
| Provider Rename with capture and verification | 24606919 | 100 |

The initial focused run stopped at a callback-scope assertion during Preparing. The repaired
harness reported both assertion controls without a dialog in
`20261124T164133Z-97372-233c6dc7388c4dcf912e8f73efb7ba85`. A later SMB run exposed
inconsistent timestamps on metadata-only and read-data opens. The replacement captures
both phases through read-data handles and preserves the pre-dispatch refusal. Functional
reruns exposed transient SMB fixture cleanup refusals. The final run passed the full case
with its no-follow cleanup.

The exact Windows SMB run is a separate server and authority profile. This observation
establishes the measured Samba cost and has no like-for-like performance baseline. A
second Samba server or configuration, x64 Release and ARM64 remain pending.

The original metric file has 759512 bytes and 2042 rows. Its SHA-256 is
`f772a241d4990eecc3b6d92e492a0e5e412b6e6991d7841dc5b6971ce47648c4`.
The retained rows preserve the original order at lines 259, 260, 297, 298, 418,
467, 511, 1193, 1348, 1431, 1570 and 1971.
