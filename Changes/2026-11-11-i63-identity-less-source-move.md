- A Move from a location that cannot identify files (a share without file IDs) now copies each
  item and removes the source when it is unchanged, instead of refusing the Move.
  - Before the copy opens a file, the Local provider reads its size, dates, attributes and file
    number. After the copy is published, the provider deletes the source by name only while those
    facts are unchanged, comparing them on the handle that deletes it. Nothing more is asked.
  - Otherwise the source is kept and the item reports `Copied; source kept` with the reason: the
    file changed, it is in use, it carries alternate data streams, extended attributes, encryption
    or a cloud placeholder (this route does not copy them), or a replaced destination file may be
    the source itself under another name. This is a Warning, not an error.
  - A source folder is removed only when the Move emptied it. A source whose open fails is still
    refused. Permanent Delete by name and its confirmation are unchanged.
  - New optional provider method `ReadUnidentifiedMetadataFeatures` on
    `IFileSystemExpectedOccupantPublication` (Local provider), four strings in five resource files,
    and new self-test `RunFileOperationIdentityLessSourceMoveSelfTest` (ten cases) in
    `FileOps_ReviewRemediationContracts`.
  - Limit: a change that keeps a file's size, dates and attributes is not noticed.
