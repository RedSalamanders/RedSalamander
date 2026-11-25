# I66 Create Folder with a delayed provider listing

An owned Dummy pane displayed the occupied default name before the test set
its directory-read delay. A direct read of the same parent took
2,001.464 ms. The real F7 prompt offered the numbered sibling in 32.660 ms.
The admitted worker kept the occupied folder, created the numbered sibling
and refreshed the pane.

- Full x64 Debug candidate `c3c6f62ac` built with zero warnings and errors.
  Receipt `de796237687cafe23715748083f3e8e9f1d7fdb9b8576c08671187e8d17e1942`.
- Exact non-activating run `20261125T141259Z-52476-4b56ae7c376e4b0183789dcafa1ec6b8`
  passed setup, `FileOps_DeferredExitRefusesNewWork` and configuration cleanup.
  The case took 15.688 seconds, including the delayed reads and the existing
  deferred exit assertions.
- First run `20261125T140210Z-21096-72ffd14a4bf1435e83c2c14677e0882c` failed
  calibration before F7. Dummy caps its configured delay at one second
  per child. The corrected fixture explicitly creates two children.

These are measurements from one controlled fixture. The focused run leaves
repository coverage `NOT_EVALUATED`.
