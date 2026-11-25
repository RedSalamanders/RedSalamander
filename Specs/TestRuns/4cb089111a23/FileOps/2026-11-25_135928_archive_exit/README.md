# I66 Pack and Unpack close admission

Pack and Unpack check the shared accepted-close state before opening options and again
after returning from options or permanent-removal consent. The refusal names the archive
command. The prompt lifetime check handles a deferred owner reset before archive work.

Full x64 Debug `ebb4912bc` built with zero warnings and errors. Receipt
`3e739d9c7469101677c0e34a741d3e3aada637701091cfd3f8d8028919c9f0c1`.

The non-activating `FileOps_DeferredExitRefusesNewWork` run
`20261125T125848Z-73424-b648007869814755a389e75125acd4fd` passed setup, the case
and cleanup. It exercised the pane Pack and Unpack commands during an accepted close
and checked each localized refusal. The native suite took 4.1 seconds. The test covers
admission. Archive worker migration, cancellation during archive I/O and cleanup
qualification remain open.
