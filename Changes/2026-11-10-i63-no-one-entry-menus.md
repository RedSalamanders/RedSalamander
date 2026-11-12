- A menu button in the File Operations popup and in the pane decision mirror is now shown only when its
  menu holds at least two commands. `Skip all` is a direct button before Cancel on a collision or retry
  prompt, and a finished task with a single command shows it as its own button. A finished task with
  no command no longer shows a `More...` button that opened an empty menu. The button and its menu
  are built from one list. Recycle escalation's `Delete all remaining permanently` and `Delete all
  remaining by name` were each alone under `More...`; each is now a direct button before Cancel, never
  the default and never the Escape action. The policy self-test, the popup cases and the visual
  gallery are updated and a test hook counts flyouts opened with fewer than two commands.
