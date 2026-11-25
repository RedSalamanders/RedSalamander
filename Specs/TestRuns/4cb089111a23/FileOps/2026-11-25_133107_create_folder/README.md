# I66 Create Folder prompt and worker

The F7 prompt reads the already displayed listing for its initial selected text. It does
no provider listing on the UI path. The admitted worker resolves collisions using provider
keys, revalidates the captured location, and creates the first eligible numbered suggestion.
A custom name stays exact.

The full x64 Debug build at `76facc43e` emitted zero warnings and errors. Receipt
`1c774883dc8c3fdae575d127470cd295c12331341a5b581e9cb70699ca76d1be`.

- Approved `cmd_pane_createDirectory_prompt_suggests_unique_selected_default` passed in
  `20261125T122836Z-2720-81624de068b7450a829e0eedd248351a`. The exact case took
  233 ms and verified the selected numbered name, creation and refresh, and no creation
  on Cancel after reopening. The bounded desktop campaign ended after 110.021 s.
- Non-activating `FileOps_DeferredExitRefusesNewWork` passed setup, the case, and cleanup
  in `20261125T123020Z-63672-609ab4199fe642b8ab3575df6bff05d3`. It included
  cancel during held Create Folder preparation, preservation of the occupied folder,
  creation of its numbered sibling, and exit refusal during a pending name prompt.

The focused runs cover an owned Local fixture and a held worker. They leave actual
slow-provider timings and the full native qualification open.
