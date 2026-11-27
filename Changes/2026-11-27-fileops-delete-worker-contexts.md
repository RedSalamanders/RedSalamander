- Fix heap corruption during parallel Local Delete by giving the scheduler's reserve worker its own
  item paths and progress context. The nested-delete test verifies one successful completion per
  selected root and supports a focused run without taking desktop input.
