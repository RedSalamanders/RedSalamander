# I66 Unpack retention after a collision skip

The owned ZIP roundtrip exercised a real Unpack command with **Skip existing
files** and an approved **Delete archive after unpacking** option. The existing
destination kept its bytes. Unpack reported a partial result, kept the
selected ZIP and admitted no cleanup Delete. The existing complete
roundtrip and replacement assertions passed in the same case.

- Full x64 Debug source `756d6258a` built with zero warnings and errors.
  Receipt `d709b0d3a31595246da2c5918fcdfd80583df451dd4a1d3b388764b59827d8d1`.
- Governed exact run `20261125T154439Z-41676-269af778083b49cf8d6377eaee7534b5`
  passed `cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation`
  without activation in 860 ms, with zero failures and skips.
- Governed tooling run `20261125T153650Z-16008-2bb93906a81a4295816493a56d1daa64`
  passed 67 runner-plan checks and 276 source and tooling checks.
  Localization run `20261125T153908Z-60892-63c308c2edfe46e99583041c519dd8a3`
  passed all 12 checks.

The focused run leaves repository coverage `NOT_EVALUATED`. Archive worker
migration and captured-source revalidation remain in the I66 plan.
