# Supplemental capture routing failure

The 45-scene x64 Debug App gallery builds with zero warnings/errors but fails
its FTP named-row assertion. The trace reports 0x80070035; an unprefixed
double-slash FTP provider path is interpreted as a Windows UNC path by the
normal navigation API. The fixture must use canonical provider-qualified paths
and assert that its injected provider remains bound. No failed frame is published.

Whole-frame review also found mojibake in the compiled Terminal Find instruction
separators. The resource source is UTF-8 but lacked its code-page declaration.
The correction adds UTF-8 compilation plus a native loaded-string assertion.
Raw frames, traces, manifests and results are preserved in the private verified
ZIP linked by preservation.json. Corrected native validation passes all 45 App scenes, including the provider
identity and loaded Unicode-string assertions. Successful frames and receipts
are retained in the supplemental scenario validation archive.

The next attempt resolves FTP and both Drive listings and corrects Terminal text,
but S3 fails after navigation reapplies catalog defaults. Visual inspection also
rejects the provider frames because the prior workspace occludes the new child.
The repair pins/restores catalog configuration and asserts the top visible child.
The build's relative-include warning is removed by querying the compiled hint
through Terminal's own test interface. This attempt remains failed evidence.


The following zero-warning build resolves the four network-provider scenes,
but MTP upload seeding reaches recovery-journal replay and fails. The captured
metrics show authority rejection and failed stale-journal quarantine. The fixture
had not scoped LOCALAPPDATA; no manual legacy profile search or cleanup follows.
The correction pre-seeds read-only RAM data, witnesses an unchanged journal-probe
counter, and binds all profile/temp paths to marked scratch before any fixture.


The read-only RAM scene renders, but an initial zero-probe assertion conflicts
with the existing MTP absence-cache contract. The corrected predicate permits
at most one initial absence check per root/device identity in the isolated empty
profile, then requires no further probes through native browsing/capture. A shadowed scratch-path local
is renamed to remove C4456. No product recovery policy is changed.

