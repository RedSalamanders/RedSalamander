# I66 SFTP read-only fixture verification

The dedicated SFTP profile's saved sign-in worked in the existing approved
test-only directory. The exact cross-provider comparison read the remote
listing and found the unique Local test file on its own side.

- Test-enabled x64 Debug source `af8eda155`, full build receipt
  `9893c3cd8eff92e184f5f0a63fa29f3ddbed8ab904dac8be91876f5908c5f661`.
- Non-activating `remote_file_sftp` run
  `20261125T143148Z-25884-078b682010554393a6607da084b0432f` returned one pass,
  zero failures and zero skips. The run used the user-authorized Windows Hello
  exception for a single isolated test profile copy.
- The installed settings digest matched before and after the run. The seeded
  test settings file was removed before archival.

The read-only comparison leaves SFTP Rename and its mutation receipts open.
