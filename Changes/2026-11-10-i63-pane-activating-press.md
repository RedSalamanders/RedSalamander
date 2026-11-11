- A plain click in the pane that is not active now makes that pane active and moves the cursor to the
  clicked item without changing that pane's selection; a click on the empty area changes nothing. Ctrl,
  Shift, right-click and double-click keep their meaning, a middle click only activates the pane, and a
  drag from an item outside the kept selection drags that item alone. The second plain click selects as
  before. Removed the main window's parent-notify handler, which compared a value Windows leaves
  undefined for mouse buttons and never matched. Three Commands cases
  (`cmd_pane_folderView_activation_press_*`) cover it; they are written and compile, and their execution
  waits for an agreed desktop time because the Commands suite takes keyboard and mouse focus.
