- A copy or move task can now carry, from its start, what it does when a file already exists: Ask (the
  default), Skip, Replace older or Keep both. No window sets it yet; the start surface of plan I18 will.
  - `OperationAdmission::presetCollisionPolicy` reaches the task through
    `Task::SeedPresetCollisionPolicy`, which writes it once on the conflict arbiter before the first
    collision.
  - Skip and Keep both decide every collision of two regular files that could take an Apply-to-all
    answer, a read-only occupant included; Replace older is the existing `Replace all older` grant set
    from the start. Links, folders and type mismatches still ask.
  - New self-test `RunFileOperationPresetCollisionPolicySelfTest` in
    `FileOps_ReviewRemediationContracts`. Not built or run when written.
