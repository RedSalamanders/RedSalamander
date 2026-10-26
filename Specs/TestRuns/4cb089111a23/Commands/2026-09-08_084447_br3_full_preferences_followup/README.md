# Preferences Keyboard Full-gate failure: isolated follow-up

Command: Tools/Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter cmd_preferences_dialog_keyboard_reordered_resized_copy_follows_visible_columns_after_search_roundtrip -SelfTestRepeat 3 -TestRoot Z:\RedSalamander.Perf
Runner: 20260908T084447Z-10708-23e3459b74f44ff598e7c88b0037c3c4.
Unchanged P1/P2/BR-3 worktree on c36d3497a0d5214168c062bf3f67a16e6baa909a;
test-enabled x64 Debug, verified full-solution receipt
d9c4193d03465a6dae3cbf594d748a63b797012eb44e0bcee511e3d39517ba1e.

All three repetitions passed. This does not reproduce the Preferences Keyboard
no-match search failure from Fresh Full run 20260908T073251Z-67752-50d9f956c380407d9fb1ef91d0b66c49.
No source change was made between these runs. This is isolated follow-up only;
no broad shuffle/order triage was run, no flaky classification is claimed, and
the original Full repository verdict remains FAILED. The unresolved wait/focus
sequence is routed through the LaterFeatures audit follow-up table.