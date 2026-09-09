# Confirmation preferences and accurate Copy/Move outcomes

Status: **ACTIVE, I18 — specification prepared; implementation not started**

Priority: **P1**. Effort: **L**, delivered in the ordered slices below. Implementation risk: **High**, because suppression of a prompt must not weaken object identity or mutation rules.

Planned at: `3dcf7ab167211451882f2c50e202aec68a5a4318`, 2026-09-08.

Owner: this plan owns the Confirmation page, confirmation-policy integration, and metadata-outcome changes described here. The authoring request is for a WIP specification only; no application behavior has changed.

This is a **non-normative implementation proposal**. Defaults below are the recommended product design. Current contracts remain in force until the corresponding implementation, tests, and authoritative-spec updates land together. The supplied Open Salamander screenshot is a reference for the two-section organization, not a requirement to reproduce its controls, exemptions, or unsupported features.

## 1. Purpose and decisions

Add **Preferences > Confirmation**, immediately after File Operations, with exactly two main sections: **Confirm on** and **Show message**. Make ordinary Copy and Move show the existing options/confirmation surface by default. Apply a consistent policy to keyboard, menu, clipboard, drag/drop, Find, Compare, and archive commands, without repeatedly asking the same question.

Separate three concerns which a checkbox must not conflate:

1. **Operation confirmation**: the user reviews the operation, destination, and applicable options before starting it.
2. **Mutation decisions**: an exact collision, protected item, changed object, or dangerous consequence needs a decision at the point where it becomes known. Confirmation of an operation does not authorize every later overwrite or loss.
3. **Result reporting**: report whether content was completed, what metadata was preserved or omitted, whether verification succeeded, and whether the source remains. Hiding a message never changes those facts.

Accepted design for this proposal:

- Copy, Move, paste, and drag/drop start confirmations default **On**.
- The explicit **Copy with Options** / **Move with Options** commands always show their options even if routine confirmation is disabled. Do not reassign their command IDs or default Shift+F5 / Shift+F6 bindings.
- File overwrite defaults **On** in every capable destination, including extraction from archives and any future writable archive provider. Reject the example's blanket “ignored in archives” exception.
- Permanent deletion, Recycle-to-permanent escalation, known security/data loss, uncertain identity, and unresolved mutation outcomes cannot be silently accepted by a message-suppression preference.
- A complete Copy with an allowed omission of ordinary metadata must not be labeled **Partial** solely for that omission. It must retain an accurate explanation. This does not make a batch atomic or turn incomplete bytes, skipped files, or failed verification into success.
- Deferred capabilities receive a documented future default, not an inert live checkbox or a persisted setting with no consumer.

## 2. Current state and evidence

Line numbers below are anchors at the planned commit; verify symbols after drift.

| Evidence | Current fact | Consequence for this plan |
|---|---|---|
| `Specs/FileSystem/FileSystem_FileOperations.md:264`, `Specs/UI/UI_FileOperationsPopup.md:137` | Routine accepted-default Copy/Move intentionally skips generic confirmation. | This plan changes an explicit product policy; current behavior is not an intermittent dialog failure. |
| `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp:874` | `transferConfirmationRequired = requireConfirmation || knownCopyOnlyMoveCount > 0u;` | Replace scattered policy selection with one immutable host policy decision; retain mandatory Copy-only Move disclosure. |
| `RedSalamander/RedSalamander.cpp:11213`, `RedSalamander/FolderWindow.FileOperations.cpp:4232`, `RedSalamander/ShortcutDefaults.cpp:180` | Routine pane commands pass the default `withOptions=false`; explicit siblings pass `true`. | Stable IDs remain; routine commands consult Confirmation defaults. |
| `RedSalamander/FolderWindow.FileOperationsInternal.h:376` | `OperationOptions` captures Links, Verify, Queue/Parallel, and bandwidth. | Add the policy snapshot and metadata policy without worker reads of live Preferences. |
| `Common/SettingsStore.h:353`, `RedSalamander/Preferences.FileOperations.cpp:80` | Existing File Operations settings/page own transfer defaults, not confirmation choices. | Create a separate page and optional settings section; do not move unrelated defaults into it. |
| `RedSalamander/Preferences.Dialog.cpp:1992,2318,2436` | Dirty detection and settings-save merge have explicit section projections. | Rendering a new page alone is insufficient; both projections must participate. |
| `Specs/FileSystem/FileSystem_FileOperations.md:971` | Permanent Delete prepares and pins selected roots, then asks on the task card before mutation. | Preserve this timing and exact-object continuity; do not move provider enumeration to an ingress modal. |
| `Specs/FileSystem/FileSystem_FileOperations.md:655`, `Plugins/FileSystem/FileSystem.FileOps.cpp:7524` | Shell-owned Recycle reports selected roots and hands the tree to the Shell without host descendant enumeration. | Protect unknown descendant classes by disclosed root-scope consent before handoff; do not promise per-child inspection or add recursive preflight. |
| `Specs/FileSystem/FileSystem_FileOperations.md:838` | Typed conflicts, cached decisions, and exact replacement authority share an engine policy. | Suppression must be integrated into that policy, not implemented as unconditional `ALLOW_OVERWRITE`. |
| `Common/PlugInterfaces/FileSystem.h:1241,1274` | Exact metadata APIs expose feature masks and per-class transfer results. | Reuse them; add only the missing result provenance/requirements needed for policy. |
| `Plugins/FileSystem/FileSystem.FileOps.cpp:4014` | Direct-final Local Copy emits `FileOps.Local.DirectFinalMetadataLost` and can return `S_OK` after complete content with metadata loss. | A trace is not a user-visible receipt. Qualify this route before adding stronger consent or strict preservation promises. |
| `Specs/FileSystem/FileSystem_FileOperations.md:1095` | Ordinary attributes/timestamps are best effort; compression/ACL inheritance is intentional; security-significant Move loss has a separate consent gate. | Do not treat all attributes as equivalent or all inheritance as loss. |
| `RedSalamander/FolderWindow.FileSystem.Commands.cpp:1369,1374` | Compressed and Encrypted checkboxes are explicitly disabled in Change Attributes. | NTFS compression/encryption command confirmation is deferred until actual mutation support exists. |
| `RedSalamander/FindFilesWindow.cpp:4477` | `OnClose()` marks close requested and cancels the session before deferred teardown. | Evaluate user-close policy before changing that state; cancellation/teardown remains asynchronous. |
| `Specs/Core/Core_SettingsStore.md:502`, `RedSalamander/FileActionLauncher.cpp` | User Menu launches configured external commands through shared macro/selection handling. | A generic command review is meaningful; there is no dedicated “names to compare” contract to reuse. |
| `Specs/Plugins/Plugins_VirtualFileSystem.md:2717`, `RedSalamander/FolderWindow.FileSystem.Commands.cpp:12403,12608` | Mounted 7z is read-only; separate Pack/Unpack commands exist, with options and extraction conflict policy. | Distinguish browsing, extraction, archive-file replacement, and future archive editing/update. |

Read-only searches of current command/settings surfaces did not establish a general Always on Top command, archive edit/write-back lease workflow, User Menu two-file comparison dialog, or Compare one/two-hour exception-notification contract. Those rows are **deferred/not applicable**, not claims that the underlying capabilities have been implemented.

### Drift check

Run from the repository root before implementation:

```powershell
git diff --stat 3dcf7ab167211451882f2c50e202aec68a5a4318..HEAD -- Common RedSalamander Plugins Tests Tools Specs
git status --short
```

Reconcile changed symbols and current domain contracts before editing. Do not overwrite independent changes. Do not reopen completed execution records. I3 retains stress evidence, the domain lifecycle contracts qualified by completed I12 remain in force, and H3 retains its six separately deferred features.

## 3. Review of the supplied list

“On” means show the identified confirmation/message when applicable. “Always” means a required decision or data-preservation contract, displayed as explanatory text rather than an editable checkbox. “Deferred” means no live row or setting in the first implementation.

### 3.1 Confirm on

| Supplied or added item | Disposition and final label | Recommended default | Exact meaning / reason |
|---|---|---|---|
| **Added: Copy operations** | **Copy operations (menu and keyboard)** | **On** | F5 and ordinary command-based Copy review source, destination, and applicable options. Clipboard/drop use their own origin switches below. |
| **Added: Move operations** | **Move operations (menu and keyboard)** | **On** | F6 and ordinary command-based Move review destination and source disposition. A known Copy-only route always discloses source retention. |
| File or directory delete | **Delete files and directories** | **On** | Controls the routine Recycle confirmation. Help: “Permanent deletion always requires confirmation.” Applies to capable Local/provider Recycle, including Find results. |
| **Added: Permanent delete** | Explanatory **Permanent deletion always asks** | **Always** | Covers Shift+Delete, non-recyclable provider deletion, and explicit archive cleanup consent. No global auto-accept setting. |
| Non-empty directory delete | **Delete non-empty directories** | **On** | A separate reason for confirmation even if routine Delete is Off. Includes the contents of an explicitly selected directory; no repeated generic prompt if those contents were already disclosed. |
| File overwrite | **Overwrite existing files** | **On** | On asks at each eligible typed collision or consumes a valid task-scoped decision. Off explicitly means automatic replacement of ordinary files after exact qualification; the UI explains this consequence. No archive exemption. |
| System or hidden file delete | **Delete hidden or system files** | **On** | Uses the source object's no-follow attributes, including descendants discovered later. Does not infer attributes from pane visibility. |
| System or hidden directory delete | **Delete hidden or system directories** | **On** | Same protection for actual directories. A directory link remains a link object, not a directory target to inspect. |
| System or hidden file overwrite | **Overwrite hidden or system files** | **On** | Uses destination attributes. A stronger reason than ordinary overwrite suppression. No archive exemption where those attributes are known. |
| **Added: Read-only file delete** | **Delete read-only files** | **On** | Clearing read-only state is a separate disclosed consequence. Setting Off cannot authorize an unqualified mutation. |
| **Added: Read-only file overwrite** | **Overwrite read-only files** | **On** | Distinct from hidden/system; combines into one typed prompt when both apply. |
| Directories merge | **Merge directories** | **Off** | Existing directories merge; they are never replaced as whole folders. On asks before the first child mutation in each unapproved merge scope. Child collisions retain their own rules. |
| NTFS compress or uncompress | **Compress or uncompress files** — deferred | **On when supported** | Current Change Attributes UI disables compression. A future explicit attribute operation must identify affected scope/capability; normal copy compression inheritance does not invoke this setting. |
| NTFS encrypt or decrypt | **Encrypt or decrypt files** — deferred | **On when supported** | Current UI disables encryption. A future decrypt command needs explicit plaintext disclosure. EFS loss during Copy/Move is already a separate mandatory risk, not this switch. |
| Drag and Drop operations | **Drag-and-drop transfers** | **On** | Covers internal and inbound drops handled by RedSalamander, including modifier-resolved Copy/Move. Does not claim control of operations performed by an external drop target. |
| Paste Operation | **Paste files (menu or Ctrl+V)** | **On** | One policy for every file-paste entry point. Text paste in editors/fields is unaffected. Show captured clipboard operation and destination before consuming a cut list. |
| User Menu: always show dialog with names to compare | Replace with **Review User Menu commands before running** | **On** | Review the configured action, executable, expanded arguments, working directory, and selected-file summary. Do not invent a two-file comparator or change macro expansion. Required-input dialogs still appear when Off. |
| **Reviewed addition: Change attributes** | Keep the existing **Change Attributes** options dialog; no new preference | **Covered by its dialog** | Current user routes already require choosing values and accepting that dialog. A new checkbox would have no effect or introduce a redundant second prompt. Recursive scope remains explicit; future quick actions must define their own confirmation policy when implemented. |

### 3.2 Show message

| Supplied or added item | Disposition and final label | Recommended default | Exact meaning / reason |
|---|---|---|---|
| Archive is about to close, only if editing files | **Close an archive with edited files** — deferred | **On when supported** | Requires a tracked edit/write-back lifecycle. Pending changes must still be saved, kept, or explicitly discarded even if an advisory close message is disabled. Do not delete extracted edits on close. |
| Do you want to close Find? | **Close Find results** | **Off** | Applies to a visible Find window with retained results. Empty/idle Find closes directly. Closing an active search uses the combined rule below. |
| Do you want to stop searching? | **Stop an active search** | **On** | Explicit user Stop/Cancel or a user action replacing an active Find search. No prompt for cancellation after an already accepted close, teardown, or session shutdown. |
| Create target path in Copy/Move | Merge into **Create a missing destination or navigation folder** | **On** | One shared setting with operation-specific wording. Accept the qualified proposed path before creating missing parents. Reject malformed/unqualified provider paths. |
| Directory does not exist; create it? | Same setting as above | **On** | Navigation offers Create and open / Cancel only for a qualified writable provider. The explicit Make Directory dialog already supplies consent. No duplicate question. |
| Turn on Always on Top | Deferred; do not add a live row | **Off when supported** | No general command was established. Fullscreen's internal topmost window style must not trigger this warning. A future explicit toggle has a clear reversible effect. |
| Close Open Salamander | Rename to **Close RedSalamander** | **Off** | Idle main-window close only. Active operations, pending decisions, dirty settings, or pending archive edits retain their own required resolution. |
| Add files into existing archive | **Update an existing archive** — deferred | **On when supported** | Present Pack creates an archive; append/update is not established. Replacing a backing archive file is ordinary overwrite and is covered now. Future update summary does not grant every entry replacement. |
| **Added: Unsupported copy metadata/options** | **Some selected attributes or options cannot be preserved** | **On** | Explain known ordinary-metadata omissions before the affected item changes. Off suppresses this advisory only; task metadata policy and required risk decisions remain unchanged. See section 7. |
| **Added: Exit while operations run** | Explanatory **Running work and unsaved changes require a decision** | **Always** | Reuse the existing cancel-all/close workflow. A global idle-close switch must never silently cancel work or accept pending decisions. |
| Screenshot-only: compare differences exactly one/two hours | Defer both specialized rows | **On if such a feature is added** | No current exception-classification contract was established. Do not silently ignore clock/DST differences; retain current Compare results. The existing Compare feature owns future controls. |
| Screenshot-only: Email selected files | Exclude | **Not applicable** | No corresponding product command was established; external User Menu launch review covers configured email programs without a dedicated checkbox. |
| Screenshot-only: Copy/Move to archives/plugin FS ignores options | Replace with the metadata/options message above | **On** | Determine capability per route and feature; “plugin” or “archive” alone never means all options are ignored. Unsupported operations remain unavailable. |

### 3.3 Deliberately non-suppressible conditions

These are policy boundaries, not extra preference sections. Explain them in a short note beneath the two sections and link the help text to the relevant File Operations behavior:

- Permanent deletion and a fresh, item-scoped Recycle-to-permanent escalation.
- Exact destination links/type mismatch, same-object/self-descendant conflicts, missing identity/authority, and changed objects. No automatic mutation merely because a checkbox is Off.
- Known EFS plaintext exposure, loss of MOTW or named-stream/EA data, unexpected security weakening, and required verification failures. Allowed actions stay route/capability qualified; absence of a safe action means retain/skip/stop.
- Existing sparse-inflation, hydration, insufficient/unknown-space, overlap/interlock, and owned-artifact decisions. This plan adds no global “continue anyway” preference for these risks.
- Unknown commit outcomes and retained incomplete content. Message visibility cannot create Retry, deletion, or cleanup authority.
- Unsaved edits/settings and pending work during exit. An accepted owning dialog can cover the same decision once; a second generic dialog is unnecessary.

## 4. Page design and settings contract

### 4.1 Layout and interaction

- Category ID **Confirmation**, localized display **Confirmation**. The page uses the shared DxUi Preferences host, typography, theme, scroll container, and accessible controls. Do not clone the legacy screenshot's native checkbox layout.
- Two section headers only: **Confirm on**, **Show message**. Order the live rows as listed above, with Copy/Move first. Keep the required-decision explanations as note cards in these sections, not disabled interactive checkboxes.
- Every live choice is a labelled toggle with short help that defines what Off does. Especially disclose: “When off, ordinary existing files may be replaced automatically. Protected files still follow their own settings.”
- Do not display deferred rows. Document them here so future owners know the intended default and prerequisite. Do not permanently gray out options for unrelated features on this page.
- Integrate category/search indexing and retained page state. Search terms include confirmation, confirm, delete, recycle, overwrite, replace, hidden, system, read-only, merge, paste, drop, copy, move, attributes, Find, exit, metadata, and User Menu.
- Support keyboard Tab/Shift+Tab and Space, pointer activation, UIA TogglePattern, stable automation IDs, accessible names/help, high contrast, RTL, and 100/150/200% DPI. Read-only notes expose text, not TogglePattern.
- Restoring defaults updates only the draft. Do not add a second global preferences-reset flow. Changes do not preview mutations or prompt policies in an already running task.

### 4.2 Proposed persisted model

Add optional `Settings::confirmations` with typed defaults. The JSON shape below contains **only live first-release keys**; write default pruning consistently with existing optional settings. Missing object and missing members use these defaults. Do not serialize deferred controls.

```json
{
  "confirmations": {
    "confirmOn": {
      "copy": true,
      "move": true,
      "delete": true,
      "nonEmptyDirectoryDelete": true,
      "fileOverwrite": true,
      "hiddenSystemFileDelete": true,
      "hiddenSystemDirectoryDelete": true,
      "hiddenSystemFileOverwrite": true,
      "readOnlyFileDelete": true,
      "readOnlyFileOverwrite": true,
      "directoryMerge": false,
      "dragDrop": true,
      "paste": true,
      "userMenuCommand": true
    },
    "showMessage": {
      "closeFindResults": false,
      "stopSearch": true,
      "createMissingPath": true,
      "closeApplication": false,
      "metadataCompatibility": true
    }
  }
}
```

There is no `permanentDelete=false`, global `disableAllSafetyPrompts`, or persisted “Move anyway” receipt. A notification flag does not store an action grant.

Additionally add the transfer default `fileOperations.metadataPolicy` with values **`compatible`** (default) and **`requirePreservation`**, on the existing File Operations page. That setting defines requested preservation, not message frequency, and therefore does not belong under `confirmations`. The options surface may override it for one Copy/managed Move. Native Move presents it as not applicable. Section 7 defines both modes.

### 4.3 Complete settings path

Implement model -> parse/write/validate/prune -> schema/UI metadata -> `workingSettings` -> dirty projection -> save-merge projection -> Apply/OK -> runtime notification. Use `Common/SettingsStore.h`, `Common/Common/SettingsStore.cpp`, `RedSalamander/SettingsSave.h`, `Specs/SettingsStore.schema.json`, and the manual projections in `Preferences.Dialog.cpp`.

- Reject non-boolean known confirmation members and invalid metadata-policy enum values through the existing settings validation/recovery mechanism. Do not coerce strings/numbers to disable confirmations. Follow the repository's established unknown-field policy.
- Omitted/default-only nested groups compare equal to effective defaults; restoring defaults can remove the optional section. Preserve explicit `false` where the default is `true`.
- Control callbacks ignore programmatic synchronization, modify only `workingSettings`, and call `SetDirty` on the root Preferences dialog HWND. Never update the baseline in a toggle callback.
- Cancel discards unapplied changes. Apply commits atomically through the normal merge/stamp mechanism, preserves unrelated settings, advances baseline, and notifies consumers. Apply failure leaves the draft dirty.
- Include the new section in reload/merge projections and schema-driven search metadata. No registry side channel or session singleton.
- Existing installations with no section receive the new recommended defaults, including Copy/Move On. This is an intentional behavior change; explain it in release notes. Do not infer user opt-out from the old absence of a setting.
- Existing shortcuts and per-provider defaults remain intact. Do not import Open Salamander registry/settings or add a mandatory migration dialog.

## 5. Shared confirmation resolution

### 5.1 One owner and immutable inputs

Add a small typed host policy resolver beside the existing File Operations plan/issue policy, not a second prompt manager. Inspect `Specs/Core/Core_SharedHelpers.md`, `Common/`, and `Tests/TestSupport/` before creating helpers. Pure policy input includes:

- stable operation kind and **origin**: pane command, explicit options, destination picker, clipboard, internal/inbound drop, Find results, Compare sync, inline rename, Batch Rename, Change Attributes, Pack/Unpack, or User Menu;
- captured effective Confirmation settings and metadata policy, with an in-memory settings revision for diagnostics;
- immutable selected endpoints/paths and requested options;
- already accepted command-owned decision coverage, with operation, selection, destination, and disclosed consequences;
- exact per-item facts only when discovered by the existing worker path.

The output is typed: **Show start options**, **Covered by owning dialog**, **Use captured defaults**, or **Needs material decision** with a finite reason set. It is not an unqualified boolean permission to mutate.

No plugin can suppress a host decision by choosing an origin. Provider/ABI callers supply validated operation facts; the host owns user-origin classification. Old flags remain input compatibility, not destructive authority.

### 5.2 Origin selection and precedence

Avoid ambiguous OR combinations that make a toggle impossible to turn off. **Select exactly one ordinary start-confirmation switch by origin**, then independently evaluate additional/mandatory reasons:

| Origin | Ordinary start policy | Owning-dialog rule |
|---|---|---|
| F5/F6, menus, ordinary destination-picker transfer | `copy` or `move` | A path picker counts only if its final action also displays the operation summary and applicable options. Merely choosing a folder is not confirmation. |
| Shift+F5/Shift+F6 / explicit With Options | Always show | Never suppress or translate into a fast command. |
| Clipboard paste, any file-paste shortcut/menu | `paste` | The Copy/Move toggles do not additionally force this origin; mandatory source-retention/risk rules still do. |
| Internal or inbound drag/drop handled by this app | `dragDrop` | Right-drag Copy/Move menu chooses the action, but does not replace the configured options review unless extended to show equivalent information. |
| Find-result transfer | `copy` or `move` | Preserve selected result mappings and Find's special summary; extend that one summary to cover the options instead of stacking a pane dialog. |
| Compare sync | `copy` or `move` plus existing specialized requirements | Use the manifest summary as the start confirmation when equivalent; never regenerate from current pane selection. Required Move summary remains. |
| Inline F2 / Batch Rename | Editor commit / preview Run | Do not add generic Copy/Move confirmation. Exact collisions still ask. |
| Change Attributes | Existing options dialog | Acceptance covers the chosen values/recursive scope. No new toggle or quick-action command is added. |
| Pack / Unpack | Archive options dialog | Include relevant host options there; container/entry collisions and delete-after actions remain independently qualified. |
| User Menu | `userMenuCommand` | One launch review immediately before process creation. No blanket interception of viewer/editor actions. |

Within a typed item decision: hard validity constraints first; mandatory risk decisions next; protected-item settings next; exact compatible task-scoped grants next; ordinary switch fallback last. If hidden/system and read-only conditions both apply, **either enabled protection requires one combined prompt**. Unknown attributes never qualify a file for automatic ordinary overwrite.

An already accepted dialog covers only its captured selection, destination, operation, and disclosed choices. Revalidate on acceptance. Changing any of them invalidates coverage. An options confirmation never covers an undiscovered descendant collision just because the user clicked Copy.

### 5.3 Start timing and prompt behavior

- Routine Copy/Move confirmation occurs before task publication and clipboard consumption. Render cheap known facts first; label undiscovered counts/metadata as unknown. Do not recursively enumerate a tree to populate the dialog.
- Copy shows Links Preserve/Skip; Move never shows a link policy. Show Verify, Queue/Parallel, bandwidth, and the new metadata policy only where meaningful. Disabled options show a localized reason. Capability budget expiry means “checked during operation,” not assumed support.
- A start prompt has operation-labelled **Copy**, **Move**, **Paste**, or **Start**, plus **Cancel**. Use Cancel as initial/default focus for a destructive Move or known lossy operation. Explicitly focused affirmative actions remain keyboard-invokable. Never accept from timer expiry, closing the dialog, or a stale posted token.
- Selection, endpoints, option values, cut sequence, and effective policy are immutable after acceptance. Settings changes apply to newly submitted operations only. Children and late-discovered items inherit the same snapshot.
- Cancel publishes no transfer task, performs no mutation, consumes no cut list, and restores appropriate pane focus after prompt teardown. Clipboard sequence changes while a prompt is open refuse stale acceptance; do not consume the replacement clipboard.
- Permanent Delete and item-dependent questions remain on the existing task card after Preparing obtains the required authority, always before the affected mutation. “Before operation” means before the governed change, not before all non-mutating preparation.
- Combine reasons known at the same boundary into one prompt. Later newly discovered risks may legitimately ask later; no full-tree preflight or impossible promise of only one prompt per batch.
- Reuse the existing task arbiter, scope cache, foreground rules, posted-payload ownership/drain, cancellation and completion fences. One actionable prompt at a time per governed task.
- Closing the File Operations popup remains **hide**, not consent or cancellation. Explicitly closing a synchronous start confirmation means Cancel. Do not conflate the two surfaces.

### 5.4 Suppressed ordinary overwrite and merge

Off has to mean a deterministic behavior, not “whatever the provider happens to do”:

- `fileOverwrite=false` selects automatic **Overwrite** only for a qualified ordinary regular-file conflict whose stronger protections are also Off/inapplicable. Show the consequence in Preferences help. Bind/refine the current destination exactly as for a visible decision, build the same typed action set, then create a **policy-origin, item-scoped** decision receipt. Immediately revalidate/consume it through the existing conditional replacement boundary.
- The receipt records action, policy origin/revision, task/item, conflict class, endpoint/path-profile identities, protected-attribute facts, and exact destination authority/expectation. A copied boolean, path string, file hash, or diagnostic row is insufficient.
- If that route cannot produce/consume qualified replacement authority, show the existing reduced safe action set. Off does not add Overwrite to an unsupported provider or outer returned-error prompt.
- A changed destination requires fresh refinement. A newly protected/type-changed/link destination goes back through its stricter policy; it cannot inherit ordinary-file permission. Existing Retry remains proof-gated.
- `directoryMerge=false` allows ordinary same-kind merge with existing no-follow checks. On asks once for the selected merge subtree and eligible task-scoped merge class, without granting child overwrite/link replacement. A folder-on-file mismatch is never merge.
- `delete=false` suppresses routine Recycle confirmation only when the selected route is actually Recycle. Non-empty/protected/read-only enabled reasons may still ask. Failure to recycle can never fall through to permanent deletion without its fresh exact-item decision.

## 6. Specialized flows

### Delete and protected descendants

For selected roots, combine known reasons in the initial Delete/Recycle or Permanent Delete card. Preserve the existing permanent-root pin and post-confirmation identity check. For **host-walked** recursive work, discover non-empty/protected conditions on workers through the bounded walker, before mutating that object/subtree. Do not scan the whole selection on the UI thread.

Determine non-emptiness by bounded first-child discovery when needed, not by counting all descendants. An unknown result stays unknown and cannot auto-approve a setting requiring confirmation. A selected non-empty directory confirmation covers its disclosed contents; a later protected descendant not covered by the accepted protected-class scope still asks. Eligible Apply to all is task/class/profile scoped. A Recycle escalation remains exact-item only with no Apply to all.

**Opaque whole-tree providers and Shell Recycle are different.** The current Recycle path hands selected roots to the provider/Shell and does not expose every descendant before removal. When an enabled non-empty/protected-descendant rule cannot be resolved without recursively preflighting that tree, include an honest root-scoped disclosure before handoff: “This folder and its contents will be recycled/deleted. Its contents may include hidden, system, or read-only items.” Include only the relevant enabled/unknown classes and allow Cancel. This acceptance covers those explicitly disclosed classes within that exact selected root; it is not a claim that the host inspected each child. Without that coverage, do not hand off the root. Do not build a recursive preflight solely to suppress this one prompt. Apply the same rule to any provider-owned recursive deletion, retaining its documented native identity limits.

Delete toggles govern requested Delete/Recycle, not the already disclosed source cleanup of an accepted Move. Managed Move source removal and removal of an emptied rename-merge directory retain their existing exact cleanup gates; they do not acquire a second generic Delete prompt. Explicit delete-after-Pack/Unpack is destructive cleanup with its existing owning consent, not an implicit Move exemption.

Do not disable these host settings just because the Windows Recycle Bin has an independent confirmation option. Where a Shell-owned operation supplies an equivalent prompt, prove coverage before avoiding a duplicate; otherwise the host retains ownership. An entire mixed selection is not declared recyclable from one Local item.

### Find and application close

- `closeFindResults=true` asks only if closing would discard nonempty results. `stopSearch=true` asks only while a search is active. A close satisfying both conditions produces one **Stop searching and close Find?** decision, not two dialogs.
- Evaluate that decision before `_closeRequested`, `_cancelRequestedUi`, timer destruction, hiding, or `_session.Cancel()`. Cancel leaves the window/search intact. Accepted close proceeds through the existing asynchronous teardown once.
- Starting a new search while one is active uses one **Stop current search and start a new one?** decision if configured. If the old search completes while the question is open, reevaluate state without issuing a stale cancel or duplicate launch.
- User-triggered Stop and OS/application teardown are distinct origins. Already accepted shutdown and mandatory bounded session-end cleanup do not generate a new modal. Preserve results from an explicitly stopped search where the current feature supports them.
- `closeApplication` governs idle user exit only. When work/settings require resolution, integrate the idle-close question into that owning decision. Do not stack it after cancel-all, or terminate work from a suppressed message. Keep the current non-blocking drain contract.

### Missing paths

One key owns both supplied creation questions. Distinguish navigation and transfer wording. Off permits creating the exact qualified missing path explicitly requested by that command; it does not authorize creating arbitrary provider roots, crossing links, or reinterpreting malformed paths.

A transfer start/options dialog can include “Create missing destination folders” and cover that decision once when the condition is already known. If found later, ask on the task card before creation. Revalidate no-follow ancestors and provider naming rules at creation. If the name becomes occupied by a file/link, show the current conflict; do not delete it. Cancel before this creation leaves that target absent, but does not undo independently completed earlier items.

### Archives and User Menu

Separate archive **entry**, **container file**, and **extracted destination** identities. Apply overwrite/protected/merge rules to the actual mutation target. Extraction must honor the same policy as ordinary Copy; its existing Replace/Skip options are explicit task choices, not an archive-wide exemption. Show only format/provider-supported metadata choices and accurately disclose omissions. Read-only archive mounts never become writable from a confirmation setting.

Pack creating a new container retains its options dialog. An existing container is not an implicit append request; replacement must use a qualified exact overwrite route or fail/offer another name. Updating entries and edit/write-back remain deferred. Existing delete-after-Pack/Unpack consent and typed source-cleanup rules remain, and must not duplicate an already captured equivalent permanent-delete decision.

User Menu review reuses the exact launch preparation/macro snapshot. Show executable, arguments, working directory, and bounded selected-file summary; do not resolve a different selection or expand macros again after acceptance. Redact secret-bearing material from diagnostics; do not persist rendered command text as an authority receipt. Cancel starts no process and disposes only owned launch temporary files. This slice does not add a comparison tool, email command, or arbitrary command editor to the review dialog.

## 7. Metadata preservation and result truth

### 7.1 What this fixes and what it cannot promise

“Some attributes cannot be copied” is not sufficient to say a file's bytes were partially copied. Conversely, suppressing a warning cannot prove that required metadata or content succeeded.

Keep the existing publication, verification, source disposition, cancellation, and indeterminate axes. Add/extend typed metadata outcome information per item: feature, requested treatment, supported/unsupported/unknown capability, attempted/preserved/changed/omitted/failed/unassessed outcome, reason/error, and any applicable accepted risk scope. Derive the UI aggregate from these facts; never derive success from whether a dialog appeared or `HRESULT` alone.

This plan does **not** promise all-or-nothing recursive Copy. A later access failure, changed file, cancellation, exhausted space, or denied risk can leave earlier complete files at the destination. Such an aggregate remains partial/incomplete as appropriate.

### 7.2 Metadata policy

| Policy | Ordinary metadata | Required/sensitive metadata | Copy and Move consequence |
|---|---|---|---|
| **Compatible** (`compatible`, default) | Attempt supported ordinary attributes/timestamps; permit documented unsupported features and normal provider precision/inheritance differences. Record exact omissions. An unexpected failure is a warning, not a claim of preservation. | No standing authorization to lose named-stream/EA data, MOTW, EFS protection, or unexpectedly weaken security. Existing exact risk gates are retained/extended to Copy as specified below. | Complete content with only allowed ordinary omissions is **Completed with warnings/details**, never Partial solely for those omissions. Ordinary best-effort warnings alone do not block otherwise-qualified managed source cleanup. |
| **Require preservation** (`requirePreservation`) | Every applicable requested source ordinary feature must be preserved within the provider's documented semantics. Known absence of support refuses/skips that item before content when possible. Unknown support is checked on workers, not assumed success. | Sensitive features still use their stronger rule; this choice does not authorize decryption or invent unsupported ownership/ACL cloning. | Unmet required preservation is **Needs attention — required metadata not preserved**; keep a managed-Move source. A complete already-published file remains reported complete/published, with the failed requirement separate. No recopy or deletion merely to simplify status. |

“Require preservation” is a bounded transfer requirement, not a backup-mode claim. Its applicable feature set and exceptions are explicit below. No source feature is classified absent simply because a provider could not assess it; mark unassessed and enforce the chosen requirement. Never compare serialized local attribute bitfields as if they had the same meaning on every provider.

The `metadataCompatibility` message explains **known ordinary** unsupported/changed features. On: show before the affected change, coalesced into an existing start/options prompt if facts are already known, otherwise at the worker decision boundary. Off: follow the captured metadata policy with visible result details and no extra advisory. It never converts Require preservation into Compatible. Switching policy requires an explicit task choice; it cannot be inferred from dismissing a message.

When a late ordinary metadata failure is discovered after publication, show it in result details; do not ask a misleading “Continue copying?” question that could replay a committed mutation. Compatible records a warning; Require preservation records an unmet requirement. Expected compatible timestamp precision alone can be informational rather than a warning.

### 7.3 Feature-specific treatment

| Feature | Required treatment |
|---|---|
| Default data stream/content | Always copy the complete logical content. Truncation/short write is a content failure, not permitted metadata loss. Verification, if requested, remains an independent required result. |
| Ordinary attributes and timestamps | Best effort under Compatible; required within documented destination precision under Require preservation. Transient access/I/O failure is distinct from known unsupported capability. Do not transfer staging-only Hidden/Temporary state as user metadata. |
| Directory metadata | Restore supported timestamps deepest-first for newly created destination directories. Do not overwrite a pre-existing merge destination's attributes/timestamps under a generic preservation option. Its existing state is an intentional exception, not an unexplained loss. |
| NTFS compression | Preserve current inheritance semantics: Copy/managed Move destination inherits its parent; native same-volume Move preserves the existing object. Neither policy requests cloning the source compression bit. Do not call inheritance a loss or gate ordinary copying on it. Explicit compress/uncompress remains deferred. |
| ACL and owner | Preserve current destination-inheritance model; exact ACL/owner cloning is outside scope. Disclose observed differences. Unexpected security weakening beyond the declared inheritance contract needs a qualified risk decision; strict mode must not silently claim a security clone. |
| ADS / named streams and extended attributes | These may contain data; never relabel them as harmless ordinary attributes. Preserve or disclose the exact known feature loss. Copy requires item/task-scoped explicit acceptance of known loss; managed Move keeps source unless a qualified “Move anyway; data/metadata lost” decision is accepted. |
| MOTW | Security-significant. Known loss requires a separate explicit risk decision before publication; message Off is not acceptance. Combine with ADS loss when the same actual stream supplies MOTW; do not double count or prompt twice. |
| EFS | Never silently decrypt. A route without a qualified plaintext-consent publication contract fails before plaintext. This plan does not add plaintext output support to Local direct-final Copy, which currently fails closed. If another route already supports it, preserve its explicit pre-write consent and source-retention rules. |
| Sparse allocation | Preserve where supported. Logical inflation is a space/cost decision before inflated writing; a generic ordinary-attribute omission setting cannot approve it. |
| Placeholders | Recall only through the existing hydration gate when actually required. A fully hydrated non-name-surrogate placeholder copies through the normal file pipeline without a fake hydration question. Preserve the BR-4 parallel-worker regression. |
| Links / hard links | Existing Copy Preserve/Skip and literal Move semantics remain. A link is not an ordinary metadata bit. Link copying/omission cannot be hidden by Compatible; an explicitly selected Skip retains existing skipped-item/result semantics. |
| Remote/archive metadata | Use advertised and observed feature/precision support. Read-only destination means unsupported operation, not “copy while ignoring options.” Unknown metadata remains unassessed in the receipt. |

### 7.4 Publication and provider integration

The direct-final Local route currently records loss in telemetry and can return success. Implementing this plan cannot stop at filtering an error label in the popup. Deliver typed per-item feature outcomes to the host for Local native Copy, bridge, and qualified provider routes, and document any capability-limited routes.

Known data/security loss must be resolved before final publication. If loss can only be learned while applying metadata, use an already supported identity-owned stage and perform the decision there. The direct-final optimization is eligible only where it can prove that the accepted policy does not require a later pre-publication decision; otherwise choose the existing staged route for that item. **Unknown is not proof of eligibility.** Retain and measure the direct-final path for ordinary proven-compatible cases.

Extend existing typed metadata interfaces/results in source-tree lockstep only where necessary. Check `sizeBytes`/versions/enums at every ABI boundary and update all producers/consumers; no optional-interface success with null outputs. Missing metadata reporting cannot be called preserved. Do not make provider diagnostics or live path queries substitute for a typed exact-object result.

For a published file whose later final-attribute application fails, do not replay Copy or delete/recreate it. Report publication and complete content truth alongside the metadata result. Under Require preservation a managed Move retains its source; under Compatible ordinary warnings alone permit existing qualified cleanup. Known ADS/MOTW/EA losses still require their explicit receipt before cleanup. Indeterminate publication/cleanup never becomes success under either policy.

### 7.5 User-visible examples and aggregation

| Scenario | Required presentation |
|---|---|
| All content and requested metadata preserved | **Completed**; verification result displayed separately. |
| Content complete, destination cannot store an ordinary attribute, Compatible permits it | **Completed with warnings**; “Contents copied; [attribute] was not preserved.” No partial-file indicator. |
| Timestamp precision changed within documented destination semantics | **Completed** with informational detail; never claim identical timestamps. |
| Same unsupported ordinary attribute, Require preservation | **Needs attention — required metadata not preserved**; content/publication state remains exact; managed source kept. |
| Named stream loss explicitly accepted | **Completed with warnings — named stream not copied**; record the accepted scope. Do not erase the loss from details. |
| One file skipped, ten copied | **Partial — 10 copied, 1 skipped**, even if metadata warning messages are Off. |
| All bytes written but requested verification failed/unavailable | Existing verification failure/unavailable result; managed source retained. Never promote to clean success. |
| Copy complete but managed source cleanup fails | **Copied; source kept**, independently of metadata outcome. |
| Unknown provider mutation or incomplete retained final leaf | **Indeterminate** / **Retained incomplete** under existing rules; never “Completed with warnings.” |

If no current aggregate enum expresses Completed with warnings, add a presentation distinction derived from typed axes; do not silently change existing numeric/public enum meanings. Batch aggregation distinguishes failed requirements, skipped/unattempted items, allowed omissions, and true unknown outcomes. Warning counts are deduplicated by item/feature; details remain bounded/paged using the existing Issues/result infrastructure. Warnings requiring attention are not auto-dismissed by `autoDismissSuccess`.

## 8. Implementation scope and authoritative closeout

The executor may create `RedSalamander/Preferences.Confirmation.h/.cpp` and focused shared policy/tests where existing helpers cannot express these semantics. Likely integration files:

- Settings and schema: `Common/SettingsStore.h`, `Common/Common/SettingsStore.cpp`, `RedSalamander/SettingsSave.h`, `Specs/SettingsStore.schema.json`, `Tests/SettingsSchemaTests/SettingsSchemaTests.cpp`.
- Preferences: `Preferences.Internal.h/.cpp`, `Preferences.Dialog.h/.cpp`, `Preferences.h`, new Confirmation pane, `Preferences.FileOperations.*`, main project source/filter lists, embedded resources and all affected localization satellites.
- Entry points: `RedSalamander.cpp`, `FolderWindow.FileOperations.cpp`, `FolderView.FileOps.cpp`, `FolderView.DragDrop.cpp`, `FolderWindow.FileSystem.Commands.cpp`, `FindFilesWindow.cpp`, `CompareDirectoriesWindow.cpp`, and `FileActionLauncher.*`/actual User Menu caller as required by the shared launch boundary.
- Engine/results: `FolderWindow.FileOperationsInternal.h`, `FolderWindow.FileOperations.State.*`, `FileOperationConfirmation.*`, popup and Issues pane; `Common/PlugInterfaces/Host.h` and `FileSystem.h` only if versioned options/metadata evidence require it.
- Providers: Local `FileSystem.FileOps.cpp`/bound metadata implementation, Dummy fixtures, and current provider/extraction producers needed to honor the same typed result/decision contract. Do not broaden provider capabilities while wiring policy.
- Tests: existing Commands Preferences, file actions, ViewCommands/Find/archive cases; FileOperations Fairstream and phase cases; CompareDirectories and SettingsSchema suites. Use canonical registration and case-family ownership.

Authoritative files to update as each behavior lands:

1. `Specs/UI/UI_PreferencesDialog.md` — page, defaults, search/accessibility, draft/Apply semantics.
2. `Specs/Core/Core_SettingsStore.md` and settings schema — new section, metadata policy, parse/prune/reload migration.
3. `Specs/FileSystem/FileSystem_FileOperations.md` — admission/origin policy, policy-origin exact replacement decisions, delete invariants, metadata requirements and aggregate truth. Explicitly reconcile the current “only typed user decision” receipt wording with qualified policy-origin ordinary overwrites.
4. `Specs/UI/UI_FileOperationsPopup.md` — options, combined reason prompts, warning/result states, hide-vs-Cancel behavior.
5. `Specs/UI/UI_CommandMenuKeyboard.md` and `Specs/UI/UI_FolderView.md` — F5/F6 defaults, explicit options stability, clipboard/drop and Change Attributes/Pack/Unpack behavior. Routine labels must follow the menu's existing ellipsis convention when they now open a dialog.
6. `Specs/UI/UI_FindFilesWindow.md`, `Specs/Core/Core_CompareDirectories.md`, `Specs/UI/UI_FolderWindow.md` — specific summary/close/navigation/launch owners.
7. `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Plugins/Plugins_PluginAPI.md` and relevant provider specs — actual supported decisions/metadata evidence, including archive read-only limitations.
8. `Specs/Testing/Testing_SelfTests.md`, `Specs/Testing/Testing_PerformanceValidation.md`, reviewed consistency anchors, and curated `Specs/TestRuns/` evidence.

Out of scope: implementing archive write-back/update, NTFS compression/encryption commands, general Always on Top, email or two-file comparison tools, changing Compare timestamp semantics, exact ACL cloning/backup mode, new provider routes, Shell Copy/Move adoption, persistent history, deferred-conflict review, and whole-batch rollback. Deferred rows in section 3 stay owned here until explicitly transferred to a named capability plan.

## 9. Ordered implementation checklist

Each slice adds behavior tests before changing the corresponding runtime policy. Record its evidence in this plan without claiming the whole program complete. Commands below are repository-root commands; section 11 defines the common environment.

### C0 — Characterize current routes and add policy tests

- [ ] Enumerate every actual user ingress listed in section 5; identify which owning dialogs already cover the decision and which create work before confirmation.
- [ ] Add pure table-driven cases for origin selection, precedence, unknown attributes, immutable settings, and duplicate suppression under a governed `ConfirmationPolicy` FileOps family/case prefix. Include every live key and every mandatory exception.
- [ ] Capture pre-change focused evidence for current F5/F6 behavior, permanent-delete identity continuity, archive extraction conflicts, and direct-final metadata outcomes. Current expected behavior can be recorded as baseline without redefining a failure as a pass.
- [ ] Confirm producer coverage for metadata results before choosing any ABI changes; list missing typed routes explicitly.

Verify: `./Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter ConfirmationPolicy -TestRoot Z:\RedSalamander.Perf` — new policy-model cases pass; baseline behavior differences remain documented until C2/C5. Run the same filter after every policy change.

### C1 — Persist settings and build the two-section page

- [ ] Add all live confirmation settings, defaults, pruning, validation, schema/search metadata and `fileOperations.metadataPolicy`.
- [ ] Implement the Confirmation pane following the File Operations pane's shared DxUi pattern; complete category creation/layout/visibility/destroy paths and both Preferences projections.
- [ ] Add localized resources, accessible IDs/help, deferred-row exclusion, explanatory mandatory notes, and metadata-policy control on File Operations.
- [ ] Add Commands cases with prefix `cmd_confirmation_preferences_` for actual toggle input, dirty/save/cancel/reload, scrolling/search, accessibility, and teardown.

Verify: `./Tools/Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_confirmation_preferences_ -TestRoot Z:\RedSalamander.Perf` — all new cases pass. SettingsSchema, resource localization and source contracts must pass through the relevant Full/Affected runner entries before continuing; Affected remains iteration-only.

### C2 — Wire start confirmation and specialized command owners

- [ ] Capture origin and settings once at host entry. Apply origin-specific start policy to F5/F6, explicit options, paste, drops, Find and Compare; upgrade owning summaries where needed.
- [ ] Preserve cancel/no-task/no-clipboard-consumption and stale-sequence checks. Capture all task options once for every child plan.
- [ ] Integrate missing-path creation, User Menu launch review and Change Attributes owning-dialog coverage.
- [ ] Add `cmd_confirmation_ingress_` cases and `ConfirmationPolicy` task snapshots. Assert exact prompt count and provider calls, not only a returned success code.

Verify: `./Tools/Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_confirmation_ingress_ -TestRoot Z:\RedSalamander.Perf` and the C0 filter — all pass. Compare sync fixture proves unchanged manifest/identity/source-cleanup semantics.

### C3 — Wire exact conflicts, protected items, delete and archive extraction

- [ ] Resolve suppression only after typed metadata refinement; add policy-origin exact-item receipts without a broad overwrite flag/cache shortcut.
- [ ] Combine ordinary/protected/read-only reasons; use worker discovery for host-walked mutations and explicit root-scope unknown-content disclosure before opaque recursive-provider/Shell handoff.
- [ ] Retain Permanent Delete/Recycling escalation and all identity/Retry/interlock boundaries. Avoid duplicate permanent-delete questions when archive cleanup already owns exact consent.
- [ ] Integrate extraction destination conflicts and backing-container replacement without implementing archive update. Unsupported routes fail honestly.
- [ ] Add `ConfirmationPolicy` adversarial fixtures for replaced occupants, protected attributes, links, mixed providers, recursive trees, archive extraction, and every On/Off combination that changes mutation.

Verify: `./Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter ConfirmationPolicy -AllowAlternateVolumeTestRoot -TestRoot Z:\RedSalamander.Perf` — all applicable cases pass, mandatory decisions never auto-accept. Existing permanent-delete and archive case families also remain green.

### C4 — Find and application messages

- [ ] Apply close/stop policy before irreversible UI state transitions; combine Find questions and handle a search finishing while the prompt is open.
- [ ] Apply idle application-close preference without bypassing active-work/dirty-state decisions or session shutdown behavior.
- [ ] Add `cmd_confirmation_messages_` cases with active/idle/resultless/nonempty states, Cancel, accepted close, repeated WM_CLOSE, race and owner teardown.

Verify: `./Tools/Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_confirmation_messages_ -TestRoot Z:\RedSalamander.Perf` — pass with no leaked HWNDs, running searches, queued payloads, or extra cancellations.

### C5 — Metadata requirements, publication and results

- [ ] Add per-item metadata policy/outcome propagation and user-visible receipts for Local, bridge and qualified provider routes; remove diagnostic-only success claims for features now requiring a visible decision/result.
- [ ] Extend sensitive-loss decisions to Copy, preserving existing Move/EFS/space/hydration rules and bounded receipt ownership.
- [ ] Qualify direct-final eligibility against policy; use existing owned staging when a pre-publication decision may be necessary. Preserve ordinary proven-compatible direct-final performance.
- [ ] Add `ConfirmationMetadata` fixtures for allowed ordinary loss, strict failure, named streams/MOTW/EA, unsupported/unassessed capabilities, native Move, late metadata failure, and aggregation.
- [ ] Update popup/Issues results and automatic dismissal so complete bytes, warnings, required failures, skipped items, verification and retained source remain distinct.

Verify: `./Tools/Run-AllTests.ps1 -Suite FileOps -CaseFilter ConfirmationMetadata -AllowAlternateVolumeTestRoot -TestRoot Z:\RedSalamander.Perf` — all deterministic cases pass with byte checks and exact publication/source assertions. Perf evidence must show which items selected direct-final versus staging.

### C6 — Broad qualification and closeout

- [ ] Run all new `cmd_confirmation_` cases in repeat/shuffle and the full Commands process, with Preferences Keyboard/native-focus and popup-close predecessor coverage.
- [ ] Run complete FileOps and Compare suites and final Fresh Full on frozen code. Isolated passes do not clear a broad failure.
- [ ] Archive same-machine baseline/candidate perf and final Full receipts; document capability skips with exact reasons.
- [ ] Merge durable behavior into every applicable authority in section 8 and update affected consistency records. Move this plan to Done only after every live slice is complete; transfer deferred capability rows to a named WIP owner first.
- [ ] Remove I18 from the active index, update H3 routing, and add the exact Done-path admission plus focused positive/neighbor-negative guard required by `Specs/README.md`.

Verify: final Fresh Full and spec inventory exit 0; all live settings have actual consumers and passing behavior coverage; no unresolved original failure is cleared by an isolated rerun.

## 10. Test matrix and performance requirements

Use the existing shared UIA, current-page native input wait, settings-artifact backup, TestSandbox, fileops fixture/provider, conflict pause and exact-ownership helpers. Never use an unrelated live app HWND or real user files. New test case names in this plan are proposed names, not claims that those cases already exist.

| Area | Required assertions |
|---|---|
| Settings | Missing/default/partial objects; true-default explicit false survives round-trip; invalid type/enum; default pruning; reload; unrelated sections survive; failed Apply; multiple settings writers/stamps. |
| Preferences | Exactly two headers; every live toggle's visible/accessible state; no deferred controls; Tab/Shift+Tab, Space, pointer and UIA; dirty/restore/Apply/Cancel; category/search round trips; stable native focus; RTL and DPI. |
| Start | Each ingress in section 5 with On/Off; force-options command; one prompt when reasons overlap; Cancel before publication; changed destination/selection; copied/cut clipboard sequences and menu/shortcut parity. |
| Destructive policy | Every protected/ordinary combination; hidden+read-only combined; unknown attrs; no-follow link; same-object/self-descendant; exact replacement changed after automatic decision; policy revision changed while task waits; mixed Local/remote selections. |
| Delete | Empty/non-empty/unknown; new protected descendant; host-walked vs opaque root handoff; unknown protected classes disclosed before Shell Recycle with zero recursive preflight; Recycle confirmed vs Off; mixed recycle capability; permanent root pin/replacement race; failed Recycle escalation exactly once; archive-owned cleanup consent; accepted Move cleanup does not get a second Delete prompt. |
| Archives | Read-only mount stays read-only; extracted ordinary/protected collisions follow policy; no “ignore in archives”; duplicate/type mismatch entries; backing container replacement vs nonexistent append command; capability-limited metadata. |
| Messages | Find close+stop combinations and empty results; search completion race; new-search replacement; idle exit vs active work/dirty state; OS shutdown does not prompt; missing-path race and cancellation. |
| Metadata | Content bytes equal in all successful fixtures; unsupported vs transient failure vs unknown capability; Compatible vs Require; loss notice On/Off changes prompt count only; security loss always gated; compressed-source inheritance; timestamp precision; existing merge-directory metadata untouched; source kept when required. |
| Results | Complete+allowed loss is not Partial; skipped child remains Partial; strict unmet metadata remains attention; failed/unavailable Verify unchanged; Unknown mutation unchanged; source cleanup failure remains source kept; no auto-dismiss of warning/attention. |
| Lifetime | Cancel during prompt metadata loading, queued work, owner shutdown and provider callback; stale payload ignored; no repeated mutation after metadata failure; exact owned-stage cleanup bounded; no leaked providers/HWNDs. |

Performance is part of C0 onward, not a final optional pass:

- Reuse `fileops.plan.construct_us`, `fileops.plan.admit_us`, existing Preparing/conflict timings, Preferences render/input metrics and provider-call counters. Add bounded `fileops.confirmation.resolve_us`, `fileops.confirmation.shown`, `fileops.confirmation.covered`, `fileops.confirmation.policy_action`, and `fileops.metadata.outcome` only where existing metrics do not express the requirement.
- Metric dimensions are finite origin/reason/action/feature enums and counts. No paths, credentials, arguments, clipboard contents, or per-frame/per-descendant logging storm. Human think time is excluded from engine timing and reported separately from prompt publication.
- Pure settings/origin resolution performs **zero provider calls** and retains a fixed-size policy snapshot, independent of selected-tree size. Measure 1, 100 and 1,024 selected roots; do not benchmark arbitrary recursive enumeration as policy cost.
- Proposed reference-profile gate: non-UI policy resolution p95 <= 1 ms for a fixed input, no extra recursive preflight, and no throughput regression greater than 5% in median of five matched no-prompt compatible Local Copy runs without an explained, approved capability-route change. Capture build/profile/dataset/cache mode and baseline spread; do not claim improvement from noisy single samples.
- New required staging scenarios are measured separately against their correctness baseline. Report staging bytes, extra I/O, queue high-water, retained receipts and cancellation latency; never compare them as if they were the unchanged direct-final route.
- Preferences page and prompt input-to-visible update p95 <= 50 ms on the reference machine after content is ready, excluding provider wait; no synchronous provider metadata/recursive calls from page rendering or policy resolution. Reuse stricter existing relevant budgets where present.
- Repeated/shuffled Commands tests must restore page, focus, runtime settings, pane providers/navigation and window state and prove a quiet baseline; preserve the preceding Keyboard search/focus fixes.

## 11. Verified command surface and environment

Use `build.ps1` and the unified runner, never ad hoc writes to `.build` as a test sandbox. The documented commands below were checked against the current runner/build guidance while authoring; the proposed new filters become executable as their cases are added.

```powershell
# Optional focused compilation during implementation; the runner normally builds.
.\build.ps1 -ProjectName RedSalamander -Configuration Debug

# Focused new UI cases, then exercise their shared process state.
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_confirmation_ -SelfTestRepeat 3 -SelfTestShuffleSeed 90818 -TestRoot Z:\RedSalamander.Perf
.\Tools\Run-AllTests.ps1 -Suite Commands -TestRoot Z:\RedSalamander.Perf

# Complete affected behavior suites.
.\Tools\Run-AllTests.ps1 -Suite FileOps -AllowAlternateVolumeTestRoot -TestRoot Z:\RedSalamander.Perf
.\Tools\Run-AllTests.ps1 -Suite Compare -TestRoot Z:\RedSalamander.Perf

# Required final application validation on frozen candidate code.
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -AllowAlternateVolumeTestRoot -TestRoot Z:\RedSalamander.Perf

# Documentation-only authoring / closeout checks.
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
git diff --check
```

Expected: runner commands exit 0, all mandatory cases execute and pass, repository/change verdict **PASSED** from Fresh Full, no unresolved quarantines or reused/provisional evidence replacing a fresh entry. `-SkipBuild` is permitted only with the runner's valid full-solution receipt; Affected/Commands-family evidence never substitutes for final Fresh Full. Archive evidence under the current TestRuns contract and validate it with `Tools/Test-TestRunArchive.ps1`.

Local test resources already authorized by the owner: **`Z:\RedSalamander.Perf` (ReFS, primary)** and **`D:\RedSalamander.Perf` (NTFS, alternate)**. Both require the ownership marker; the alternate resource is declared in `Z:\RedSalamander.Perf\config\machine-resources.json`, with explanatory `config\local-volumes.md`. Verify actual filesystem/capabilities at run time. Do not assume ReFS lacks every NTFS feature or manufacture capability skips; use Dummy fault/capability fixtures where physical FAT/exFAT/cloud capabilities are absent. No format/mount/new-drive work is authorized by this plan.

Before any Tools change, follow the tooling-governance skill and inventory/help/caller tests. `Get-SpecInventory.ps1`'s generated repository inventory is a tooling report under its governed `.build/SpecInventory` path; runtime test data stays under the owned Perf roots. Never kill an independently launched app to replace its build output.

## 12. Completion and stop conditions

The application implementation is complete only when:

- [ ] Every live row/default in section 3 is implemented through the full settings and UI path, with no inert switches.
- [ ] Every operation origin resolves through one policy; no accidental bypass or duplicate generic confirmation remains.
- [ ] Suppressed ordinary overwrite uses qualified policy-origin exact-item authority, and mandatory/risk decisions remain intact.
- [ ] Known sensitive loss cannot use direct-final publication before a required decision; all published outcomes have typed metadata truth.
- [ ] Compatible ordinary omissions do not cause false Partial labels; required preservation/content/verification/source-retention failures remain visible.
- [ ] Tests and same-machine perf evidence meet sections 9–11, final Fresh Full is green, and domain specs/schema/resources agree.
- [ ] Deferred capabilities have a named remaining WIP owner and are not claimed delivered; exact Done admission and index/H3 routing are updated.

Stop only the affected slice and report the concrete issue if an assumed provider route cannot deliver exact conflict/publication/metadata evidence, a safe action set cannot express the proposed automatic action, or implementing a deferred capability would be necessary for a live toggle. Do not improvise path-based destructive fallback, silently weaken strict policy, enable dead controls, or relabel an indeterminate outcome. Reconcile ordinary source drift locally; unrelated work can proceed while a truly blocked capability is documented.

This planning pass reviewed the listed user workflows, relevant command/settings/metadata paths, and current specs. It did not run application mutations, perform a repository-wide security audit, validate live remote providers, or prove runtime performance. Those are implementation validation duties, not evidence supplied by this WIP document.
