#include "FolderWindow.FileOperations.Tests.VisualGallerySurfaces.cpp"

// Included by the FileOperations selftest dispatcher. These are presentation fixtures,
// not simulated I/O successes: the real popup, engine action policy and renderer are used.
struct FileOpsVisualScenario final
{
    std::wstring name;
    std::vector<FileOperationsPopupInternal::TaskSnapshot> tasks;
    bool collapsed                                       = false;
    bool footerOnly                                      = false;
    bool groupExpanded                                   = true;
    bool compactDensity                                  = false;
    FileOperationsPopupInternal::PopupHitTest::Kind menu = FileOperationsPopupInternal::PopupHitTest::Kind::None;
    bool graphicsFailure                                 = false;
};

template <typename VerifyProvider>
[[nodiscard]] bool VerifyGalleryHostedProvider(HWND popup, std::wstring_view expectedId, bool expectPresent, VerifyProvider&& verify, std::wstring& failure)
{
    constexpr size_t kMaxNodes             = 512u;
    constexpr size_t kMaxSiblingsPerParent = 128u;
    wil::com_ptr_nothrow<IRawElementProviderFragmentRoot> root;
    root.attach(DxUi::CreateWindowHostAccessibilityProvider(popup));
    wil::com_ptr_nothrow<IRawElementProviderFragment> rootFragment;
    if (! root || FAILED(root.query_to(rootFragment.put())) || ! rootFragment)
    {
        failure = L"Hosted UIA fragment root is unavailable.";
        return false;
    }

    std::vector<wil::com_ptr_nothrow<IRawElementProviderFragment>> pending;
    pending.push_back(std::move(rootFragment));
    size_t visited           = 0u;
    size_t namedNodes        = 0u;
    size_t matchingProviders = 0u;
    while (! pending.empty())
    {
        auto fragment = std::move(pending.back());
        pending.pop_back();
        if (++visited > kMaxNodes)
        {
            failure = L"Hosted UIA fragment traversal exceeded its node bound.";
            return false;
        }
        wil::com_ptr_nothrow<IRawElementProviderSimple> simple;
        if (! fragment || FAILED(fragment.query_to(simple.put())) || ! simple)
        {
            failure = L"Hosted UIA fragment could not be read as a simple provider.";
            return false;
        }

        wil::unique_variant name;
        if (FAILED(simple->GetPropertyValue(UIA_NamePropertyId, &name)))
        {
            failure = L"Hosted UIA Name property read failed.";
            return false;
        }
        if (name.vt == VT_BSTR && name.bstrVal && SysStringLen(name.bstrVal) != 0u)
            ++namedNodes;
        else if (name.vt != VT_EMPTY && name.vt != VT_BSTR)
        {
            failure = L"Hosted UIA Name property returned an unexpected value type.";
            return false;
        }

        wil::unique_variant automationId;
        if (FAILED(simple->GetPropertyValue(UIA_AutomationIdPropertyId, &automationId)))
        {
            failure = L"Hosted UIA AutomationId property read failed.";
            return false;
        }
        if (automationId.vt != VT_EMPTY && automationId.vt != VT_BSTR)
        {
            failure = L"Hosted UIA AutomationId property returned an unexpected value type.";
            return false;
        }
        const bool exactMatch =
            automationId.vt == VT_BSTR && automationId.bstrVal && expectedId == std::wstring_view(automationId.bstrVal, SysStringLen(automationId.bstrVal));
        if (exactMatch)
        {
            ++matchingProviders;
            if (expectPresent && ! verify(simple.get()))
                return false;
        }

        wil::com_ptr_nothrow<IRawElementProviderFragment> child;
        const HRESULT childHr = fragment->Navigate(NavigateDirection_FirstChild, child.put());
        if (FAILED(childHr))
        {
            failure = L"Hosted UIA FirstChild traversal failed.";
            return false;
        }
        size_t siblings = 0u;
        while (child)
        {
            if (++siblings > kMaxSiblingsPerParent || visited + pending.size() >= kMaxNodes)
            {
                failure = L"Hosted UIA sibling traversal exceeded its bound.";
                return false;
            }
            wil::com_ptr_nothrow<IRawElementProviderFragment> next;
            const HRESULT siblingHr = child->Navigate(NavigateDirection_NextSibling, next.put());
            if (FAILED(siblingHr))
            {
                failure = L"Hosted UIA NextSibling traversal failed.";
                return false;
            }
            pending.push_back(std::move(child));
            child = std::move(next);
        }
    }
    if (visited < 2u || namedNodes == 0u)
    {
        failure = L"Hosted UIA traversal did not prove a nonempty named tree.";
        return false;
    }
    const size_t expectedCount = expectPresent ? 1u : 0u;
    if (matchingProviders != expectedCount)
    {
        failure =
            std::format(L"Hosted AutomationId {} count was {}; expected {} (offscreen fragments included).", expectedId, matchingProviders, expectedCount);
        return false;
    }
    return true;
}

[[nodiscard]] bool VerifyApplyToAllScopeProvider(HWND popup, uint64_t taskId, bool expectPresent, bool expectChecked, std::wstring& failure)
{
    FileOperationsPopupInternal::TaskSnapshot snapshot{};
    if (! DebugGetFileOperationsPopupTaskSnapshot(popup, taskId, snapshot) || ! snapshot.conflict.active || snapshot.conflict.generation == 0u)
    {
        failure = std::format(L"Apply-to-all scope provider has no current conflict generation for task {}.", taskId);
        return false;
    }
    const std::wstring expectedId = std::format(L"FileOperations.Action.{}.{}.0.{}",
                                                static_cast<uint32_t>(FileOperationsPopupInternal::PopupHitTest::Kind::TaskConflictToggleApplyToAll),
                                                taskId,
                                                snapshot.conflict.generation);
    return VerifyGalleryHostedProvider(popup,
                                       expectedId,
                                       expectPresent,
                                       [&](IRawElementProviderSimple* simple)
    {
        wil::unique_variant controlType;
        if (FAILED(simple->GetPropertyValue(UIA_ControlTypePropertyId, &controlType)) || controlType.vt != VT_I4 ||
            controlType.lVal != UIA_CheckBoxControlTypeId)
        {
            failure = L"Apply-to-all scope provider is not a UIA CheckBox.";
            return false;
        }
        wil::com_ptr_nothrow<IUnknown> pattern;
        wil::com_ptr_nothrow<IToggleProvider> toggle;
        if (FAILED(simple->GetPatternProvider(UIA_TogglePatternId, pattern.put())) || ! pattern || FAILED(pattern.query_to(toggle.put())) || ! toggle)
        {
            failure = L"Apply-to-all CheckBox does not expose IToggleProvider.";
            return false;
        }
        ToggleState state{};
        if (FAILED(toggle->get_ToggleState(&state)) || state != (expectChecked ? ToggleState_On : ToggleState_Off))
        {
            failure = L"Apply-to-all CheckBox toggle state does not match the gallery scenario.";
            return false;
        }
        return true;
    },
                                       failure);
}

[[nodiscard]] std::vector<FileOpsVisualScenario> BuildFileOpsVisualScenarios()
{
    using Snapshot = FileOperationsPopupInternal::TaskSnapshot;
    using Phase    = FileOperations::TaskLifecyclePhase;
    Snapshot base{};
    base.taskId                     = 1u;
    base.operation                  = FILESYSTEM_COPY;
    base.lifecyclePhase             = static_cast<uint8_t>(Phase::Running);
    base.started                    = true;
    base.discoveryClosed            = true;
    base.hasProgressCallbacks       = true;
    base.lastProgressCallbackTick   = GetTickCount64();
    base.operationStartTick         = base.lastProgressCallbackTick - 25'000u;
    base.currentSourcePath          = L"C:\\Photographies\\Vacances en famille\\Photographie retouchée — septembre 2026.raw";
    base.currentDestinationPath     = L"D:\\Sauvegardes\\Vacances en famille\\Photographie retouchée — septembre 2026.raw";
    base.destinationFolder          = L"D:\\Backups\\Summer trip";
    base.destinationPluginId        = L"builtin/file-system";
    base.destinationPane            = FolderWindow::Pane::Right;
    base.destinationPluginShortId   = L"file";
    base.totalBytes                 = 8ull * 1024ull * 1024ull * 1024ull;
    base.completedBytes             = 3ull * 1024ull * 1024ull * 1024ull;
    base.totalItems                 = 128u;
    base.plannedItems               = 128u;
    base.completedItems             = 48u;
    base.itemTotalBytes             = 512ull * 1024ull * 1024ull;
    base.itemCompletedBytes         = 192ull * 1024ull * 1024ull;
    base.effectiveConcurrencyBudget = 4u;
    std::vector<FileOpsVisualScenario> result;
    const auto add = [&](std::wstring name, Snapshot task) { result.push_back({std::move(name), {std::move(task)}}); };
    result.push_back({L"01-empty", {}});
    auto task                 = base;
    task.started              = false;
    task.operationStartTick   = 0u;
    task.lifecyclePhase       = static_cast<uint8_t>(Phase::Preparing);
    task.totalBytes           = 0u;
    task.totalItems           = 0u;
    task.completedBytes       = 0u;
    task.completedItems       = 0u;
    task.hasProgressCallbacks = false;
    add(L"02-preparing", task);
    task.lifecyclePhase = static_cast<uint8_t>(Phase::AwaitingAcceptance);
    add(L"03-awaiting-acceptance", task);
    task                  = base;
    task.started          = false;
    task.waitingInQueue   = true;
    task.canMoveQueueDown = true;
    add(L"04-queued-start-now", task);
    task.overlapQueued = true;
    add(L"05-queued-overlap-no-start-now", task);
    task                          = base;
    task.discoveryClosed          = false;
    task.discoveryAheadActive     = true;
    task.operationStartTick       = 0u;
    task.started                  = false;
    task.discoveredFileCount      = 128u;
    task.discoveredDirectoryCount = 12u;
    task.discoveredTotalBytes     = base.totalBytes;
    task.discoveryElapsedMs       = 3200u;
    task.completedBytes           = 0u;
    task.completedItems           = 0u;
    add(L"06-discovery-before-transfer", task);
    task.started            = true;
    task.operationStartTick = base.operationStartTick;
    task.completedBytes     = base.completedBytes;
    task.completedItems     = 48u;
    add(L"07-discovery-and-transfer", task);
    task.discoveryAheadActive = false;
    task.discoverySkipped     = true;
    add(L"08-discovering-as-needed", task);
    add(L"09-copy-single-stream", base);
    task = base;
    for (size_t i = 0u; i < 4u; ++i)
    {
        auto& file            = task.inFlightFiles[i];
        file.progressStreamId = i + 1u;
        file.sourcePath       = std::format(L"C:\\Photos\\Summer trip\\IMG_{:04}.raw", i + 1u);
        file.totalBytes       = base.itemTotalBytes;
        file.completedBytes   = (i + 1u) * 64ull * 1024ull * 1024ull;
        file.lastUpdateTick   = base.lastProgressCallbackTick;
    }
    task.inFlightFileCount = 4u;
    add(L"10-copy-four-streams", task);
    auto parallel                          = task;
    task.desiredSpeedLimitBytesPerSecond   = 80ull * 1024ull * 1024ull;
    task.effectiveSpeedLimitBytesPerSecond = task.desiredSpeedLimitBytesPerSecond;
    add(L"11-speed-limited", task);
    task           = base;
    task.operation = FILESYSTEM_MOVE;
    add(L"12-move", task);
    task.operation      = FILESYSTEM_DELETE;
    task.totalBytes     = 0u;
    task.completedBytes = 0u;
    add(L"13-delete", task);
    task        = base;
    task.paused = true;
    add(L"14-paused", task);
    task                       = base;
    task.lifecyclePhase        = static_cast<uint8_t>(Phase::Stopping);
    task.cancellationRequested = true;
    add(L"15-stopping", task);
    task = base;
    task.lastProgressCallbackTick -= 60'000u;
    add(L"16-callback-silence", task);
    task                            = base;
    task.verificationRequested      = true;
    task.verificationActive         = true;
    task.completedBytes             = task.totalBytes;
    task.verificationTotalBytes     = task.totalBytes;
    task.verificationCompletedBytes = task.totalBytes / 2u;
    task.verifiedItemCount          = 64u;
    add(L"17-verifying", task);
    task                = base;
    task.finished       = true;
    task.lifecyclePhase = static_cast<uint8_t>(Phase::Terminal);
    task.completedBytes = task.totalBytes;
    task.completedItems = task.totalItems;
    add(L"18-completed", task);
    auto done                  = task;
    task.verificationRequested = true;
    task.verificationState     = static_cast<uint8_t>(FileOperations::VerificationState::Verified);
    add(L"19-completed-verified", task);
    task                       = done;
    task.resultHr              = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    task.warningCount          = 3u;
    task.resultSummary         = L"125 fichiers copiés ; 3 ignorés. Les fichiers ignorés sont conservés à leur emplacement source.";
    task.lastDiagnosticMessage = L"Trois fichiers ont été ignorés après un conflit de noms.";
    add(L"20-partial-skipped", task);
    auto partial                            = task;
    task.operation                          = FILESYSTEM_MOVE;
    task.retainedSourceCount                = 3u;
    task.hasSourceActionPaths               = true;
    task.resultSummary                      = L"Copie effectuée ; source conservée. Trois fichiers restent à leur emplacement d’origine.";
    task.exactRetainedSourceActionAvailable = true;
    add(L"21-move-source-kept", task);
    task.clipboardMoveAdmission = true;
    task.clipboardMoveConsumed  = true;
    add(L"22-clipboard-move-source-kept", task);
    task.unknownSourceCount                 = 1u;
    task.exactRetainedSourceActionAvailable = false;
    task.hasIndeterminateResult             = true;
    task.resultSummary                      = L"L’état de certaines sources est inconnu. La liste des fichiers coupés a été vidée et n’a pas été restaurée.";
    add(L"23-move-source-unknown", task);
    task                       = done;
    task.resultHr              = E_ACCESSDENIED;
    task.errorCount            = 1u;
    task.lastDiagnosticMessage = L"Accès refusé. Le dossier de destination n’a pas pu être ouvert.";
    task.resultSummary         = L"Aucun fichier n’a été copié.";
    add(L"24-failed", task);
    auto failed        = task;
    task               = base;
    task.finished      = true;
    task.resultHr      = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    task.resultSummary = L"Opération annulée. 48 fichiers ont été copiés ; 80 fichiers n’ont pas été copiés.";
    add(L"25-canceled-after-progress", task);
    auto canceled               = task;
    task                        = done;
    task.resultHr               = HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE);
    task.hasIndeterminateResult = true;
    task.interruptedMoveNotice  = true;
    task.operation              = FILESYSTEM_MOVE;
    task.interruptedOperationId = 314u;
    task.resultSummary = L"Résultat inconnu. Le déplacement a été interrompu pendant son exécution. Vérifiez la source et la destination avant de recommencer.";
    add(L"26-interrupted-move", task);
    task = base;
    // Synthetic presentation-only UNC value; never passed to filesystem admission.
    // >4 KiB with French accents, emoji and RTL text, matching the path-detail interaction fixture.
    std::wstring longUncSource = L"\\\\serveur\\archives\\";
    for (size_t index = 0u; index < 100u; ++index)
        longUncSource += L"Collection familiale conservée pendant plusieurs générations\\";
    longUncSource += L"école-e\u0301-שלום-\U0001F4F7\\Dernier fichier.raw";
    task.currentSourcePath      = std::move(longUncSource);
    task.currentDestinationPath = L"D:\\Backup collections\\International archive\\Photographs\\東京 — été — München\\IMG_2026_0916_final_edited_version.raw";
    add(L"27-long-unicode-paths", task);
    result.push_back({L"28-compact-progress", {parallel}, true});
    result.push_back({L"29-footer-only", {parallel}, false, true});
    auto waiting           = base;
    waiting.taskId         = 2u;
    waiting.waitingInQueue = true;
    waiting.started        = false;
    waiting.canMoveQueueUp = true;
    result.push_back({L"30-running-and-queued", {parallel, waiting}});
    done.taskId     = 1u;
    partial.taskId  = 2u;
    failed.taskId   = 3u;
    canceled.taskId = 4u;
    result.push_back({L"31-completed-group-expanded", {done, partial, failed, canceled}, true});
    result.push_back({L"32-completed-group-collapsed", {done, partial, failed, canceled}, true, false, false});
    constexpr std::array<std::wstring_view, 26> names{{L"file-exists",
                                                       L"read-only-file",
                                                       L"read-only-exists",
                                                       L"type-mismatch",
                                                       L"destination-link",
                                                       L"name-not-representable",
                                                       L"target-conflict",
                                                       L"sharing-violation",
                                                       L"access-denied",
                                                       L"disk-full",
                                                       L"path-too-long",
                                                       L"recycle-failed",
                                                       L"insufficient-space",
                                                       L"space-unknown",
                                                       L"efs-plaintext",
                                                       L"sparse-inflation",
                                                       L"placeholder-hydration",
                                                       L"metadata-loss",
                                                       L"network-offline",
                                                       L"unsupported-reparse",
                                                       L"same-host-overlap",
                                                       L"same-host-live-output",
                                                       L"permanent-delete",
                                                       L"unknown-error",
                                                       L"permanent-delete-by-name",
                                                       L"recycle-failed-by-name"}};
    static_assert(names.size() == static_cast<size_t>(FileOperations::ConflictClass::Count));
    // I35 S4: the by-name classes follow Unknown in the enumeration. Their scenarios are added after the
    // whole existing gallery (150, 151) so every earlier scenario keeps its index and name.
    constexpr size_t kFirstByNameClass = static_cast<size_t>(FileOperations::ConflictClass::PermanentDeleteByNameConfirmation);
    const auto makeDecision            = [&](size_t i)
    {
        Snapshot decision        = base;
        decision.conflict.active = true;
        // Synthetic prompts still need a real presentation identity for hosted
        // control bindings and UIA; zero is the production stale/invalid sentinel.
        decision.conflict.generation = static_cast<uint64_t>(i) + 1u;
        decision.conflict.bucket     = static_cast<uint8_t>(i);
        constexpr std::array<DWORD, 26> errors{{ERROR_FILE_EXISTS,
                                                ERROR_ACCESS_DENIED,
                                                ERROR_FILE_EXISTS,
                                                ERROR_ALREADY_EXISTS,
                                                ERROR_REPARSE_TAG_MISMATCH,
                                                ERROR_INVALID_NAME,
                                                ERROR_ALREADY_EXISTS,
                                                ERROR_SHARING_VIOLATION,
                                                ERROR_ACCESS_DENIED,
                                                ERROR_DISK_FULL,
                                                ERROR_FILENAME_EXCED_RANGE,
                                                ERROR_NOT_SUPPORTED,
                                                ERROR_DISK_FULL,
                                                ERROR_SUCCESS,
                                                ERROR_SUCCESS,
                                                ERROR_SUCCESS,
                                                ERROR_SUCCESS,
                                                ERROR_NOT_SUPPORTED,
                                                ERROR_NETWORK_UNREACHABLE,
                                                ERROR_NOT_SUPPORTED,
                                                ERROR_SUCCESS,
                                                ERROR_SUCCESS,
                                                ERROR_SUCCESS,
                                                ERROR_GEN_FAILURE,
                                                ERROR_SUCCESS,
                                                ERROR_NOT_SUPPORTED}};
        decision.conflict.status          = HRESULT_FROM_WIN32(errors[i]);
        decision.conflict.sourcePath      = base.currentSourcePath;
        decision.conflict.destinationPath = base.currentDestinationPath;
        // These fixture paths deliberately use the local Windows namespace.
        decision.conflict.sourceName                    = std::filesystem::path(decision.conflict.sourcePath).filename().wstring();
        decision.conflict.sourceParentPath              = std::filesystem::path(decision.conflict.sourcePath).parent_path().wstring();
        decision.conflict.destinationName               = std::filesystem::path(decision.conflict.destinationPath).filename().wstring();
        decision.conflict.destinationParentPath         = std::filesystem::path(decision.conflict.destinationPath).parent_path().wstring();
        decision.conflict.sourceMetadata.available      = true;
        decision.conflict.sourceMetadata.sizeKnown      = true;
        decision.conflict.sourceMetadata.sizeBytes      = 512ull * 1024ull * 1024ull;
        decision.conflict.sourceMetadata.lastWriteTime  = 134'338'680'000'000'000ll;
        decision.conflict.destinationMetadata           = decision.conflict.sourceMetadata;
        decision.conflict.destinationMetadata.sizeBytes = 480ull * 1024ull * 1024ull;
        decision.conflict.destinationMetadata.lastWriteTime -= 864'000'000'000ll;
        decision.conflict.overlapTaskId      = 2u;
        decision.conflict.overlapProblem     = static_cast<uint8_t>(FileOperations::SameHostOverlapProblem::RemovesRead);
        decision.conflict.consentDetail      = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILE, 128ull, std::wstring_view(L"s"));
        decision.conflict.factItemCountKnown = i >= 11u;
        decision.conflict.factItemCount      = 128u;
        decision.conflict.factBytesKnown     = i >= 11u;
        decision.conflict.factBytes          = base.totalBytes;
        decision.conflict.deferredConsent    = i >= 11u && i != 18u && i != 19u && i != 23u;
        if (i == 11u || i == 22u || i >= kFirstByNameClass)
            decision.operation = FILESYSTEM_DELETE;
        if (i == 17u)
            decision.operation = FILESYSTEM_MOVE;
        if (i >= kFirstByNameClass)
        {
            // The by-name consents end with how the location deletes and the leftover warning.
            decision.conflict.consentNote =
                LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_BY_NAME_CHECKED) + L" " + LoadStringResource(nullptr, IDS_FILEOPS_CONSENT_POSSIBLE_LEFTOVER);
        }
        return decision;
    };
    for (size_t i = 0u; i < kFirstByNameClass; ++i)
    {
        add(std::format(L"{:02}-decision-{}", 33u + i, names[i]), makeDecision(i));
    }
    task                          = result[32].tasks.front();
    task.conflict.metadataLoading = true;
    add(L"57-decision-metadata-loading", task);
    task.conflict.metadataLoading               = false;
    task.conflict.destinationMetadata.available = false;
    add(L"58-decision-replacement-withheld", task);
    task.conflict.attemptCount = 2u;
    task.conflict.retryFailed  = true;
    task.conflict.bucket       = static_cast<uint8_t>(FileOperations::ConflictClass::AccessDenied);
    add(L"59-decision-retry-failed", task);
    task                            = result[32].tasks.front();
    task.conflict.applyToAllChecked = true;
    add(L"60-decision-apply-to-all-checked", task);
    const std::filesystem::path scanStressPath = L"C:\\Photographies\\Archives familiales de plusieurs générations\\Sélection définitive 📷\\Photographie "
                                                 L"originale — re\u0301union familiale de septembre 2026.raw";
    for (size_t i = 0u; i < 4u; ++i)
    {
        Snapshot info{};
        info.kind                             = Snapshot::Kind::Informational;
        info.taskId                           = 1u;
        auto& detail                          = info.informational;
        detail.taskId                         = 1u;
        detail.kind                           = static_cast<FolderWindow::InformationalTaskUpdate::Kind>(i);
        detail.title                          = i == 0u   ? L"Comparer les dossiers"
                                                : i == 1u ? L"Modifier les majuscules et minuscules"
                                                : i == 2u ? L"Modifier les attributs"
                                                          : L"Créer une liste de fichiers";
        detail.leftRoot                       = L"C:\\Photos";
        detail.rightRoot                      = L"D:\\Backups";
        detail.scanActive                     = true;
        detail.scanFolderCount                = 12u;
        detail.scanEntryCount                 = 128u;
        detail.scanCurrentRelative            = L"Summer trip\\IMG_2026_0916.raw";
        detail.changeCaseEnumerating          = true;
        detail.changeCaseScannedEntries       = 128u;
        detail.changeCasePlannedRenames       = 42u;
        detail.changeCaseCurrentPath          = scanStressPath;
        detail.changeAttributesApplying       = true;
        detail.changeAttributesPlannedItems   = 128u;
        detail.changeAttributesCompletedItems = 48u;
        detail.changeAttributesCurrentPath    = scanStressPath;
        detail.makeFileListRendering          = true;
        detail.makeFileListTotalEntries       = 128u;
        detail.makeFileListRenderedEntries    = 48u;
        detail.makeFileListCurrentPath        = scanStressPath;
        add(std::format(L"{:02}-{}",
                        61u + i,
                        i == 0u   ? L"compare-scanning"
                        : i == 1u ? L"change-case-enumerating"
                        : i == 2u ? L"change-attributes-applying"
                                  : L"file-list-rendering"),
            info);
    }
    const auto addVariant = [&](std::wstring_view label, Snapshot snapshot) { add(std::format(L"{:02}-{}", result.size() + 1u, label), std::move(snapshot)); };
    for (const auto verification :
         {FileOperations::VerificationState::Failed, FileOperations::VerificationState::Unavailable, FileOperations::VerificationState::Canceled})
    {
        task                              = done;
        task.taskId                       = 1u;
        task.verificationRequested        = true;
        task.verificationState            = static_cast<uint8_t>(verification);
        task.verificationProblemItemCount = 3u;
        addVariant(verification == FileOperations::VerificationState::Failed        ? L"verification-failed"
                   : verification == FileOperations::VerificationState::Unavailable ? L"verification-unavailable"
                                                                                    : L"verification-canceled",
                   task);
    }
    task                    = base;
    task.totalBytes         = 0u;
    task.completedBytes     = 0u;
    task.itemTotalBytes     = 0u;
    task.itemCompletedBytes = 0u;
    addVariant(L"zero-byte-items", task);
    task                = partial;
    task.taskId         = 1u;
    task.completedBytes = 0u;
    task.completedItems = 0u;
    task.warningCount   = 128u;
    task.resultSummary  = L"Les 128 fichiers ont été ignorés. Aucun fichier n’a été copié.";
    addVariant(L"all-items-skipped", task);
    task             = base;
    task.taskId      = 1u;
    task.queuePaused = true;
    addVariant(L"queue-paused", task);
    task                    = waiting;
    task.taskId             = 1u;
    task.started            = false;
    task.operationStartTick = 0u;
    task.waitingInQueue     = false;
    task.waitingForOthers   = true;
    task.canMoveQueueUp     = false;
    task.canMoveQueueDown   = false;
    addVariant(L"waiting-for-other-task", task);
    task                      = parallel;
    task.autoConcurrencyUsed  = true;
    task.autoTunedConcurrency = 4u;
    addVariant(L"automatic-concurrency", task);
    addVariant(L"compact-density", parallel);
    result.back().compactDensity = true;
    task                         = parallel;
    for (size_t i = 4u; i < task.inFlightFiles.size(); ++i)
    {
        task.inFlightFiles[i]                  = task.inFlightFiles[i % 4u];
        task.inFlightFiles[i].progressStreamId = i + 1u;
        task.inFlightFiles[i].sourcePath       = std::format(L"C:\\Photos\\Summer trip\\IMG_{:04}.raw", i + 1u);
    }
    task.inFlightFileCount = task.inFlightFiles.size();
    addVariant(L"sixteen-streams", task);
    task                                 = result[60].tasks.front();
    auto& compare                        = task.informational;
    compare.scanActive                   = false;
    compare.contentActive                = true;
    compare.contentCurrentRelative       = L"Summer trip\\IMG_2026_0916.raw";
    compare.contentTotalBytes            = base.totalBytes;
    compare.contentCompletedBytes        = base.completedBytes;
    compare.contentCurrentTotalBytes     = base.itemTotalBytes;
    compare.contentCurrentCompletedBytes = base.itemCompletedBytes;
    compare.contentPendingCount          = 80u;
    compare.contentCompletedCount        = 48u;
    compare.contentEtaSeconds            = 42u;
    addVariant(L"compare-content", task);
    task.informational.scanActive = true;
    addVariant(L"compare-scan-and-content", task);
    task                                          = result[61].tasks.front();
    task.informational.changeCaseEnumerating      = false;
    task.informational.changeCaseRenaming         = true;
    task.informational.changeCaseCompletedRenames = 16u;
    task.informational.changeCaseCurrentPath      = scanStressPath;
    addVariant(L"change-case-renaming", task);
    task                                              = result[62].tasks.front();
    task.informational.changeAttributesApplying       = false;
    task.informational.changeAttributesEnumerating    = true;
    task.informational.changeAttributesScannedEntries = 128u;
    addVariant(L"change-attributes-enumerating", task);
    task                                          = result[63].tasks.front();
    task.informational.makeFileListRendering      = false;
    task.informational.makeFileListCollecting     = true;
    task.informational.makeFileListScannedEntries = 128u;
    addVariant(L"file-list-collecting", task);
    task.informational.makeFileListCollecting = false;
    task.informational.makeFileListWriting    = true;
    addVariant(L"file-list-writing", task);
    for (size_t i = 0u; i < 4u; ++i)
    {
        for (const HRESULT outcome : {S_OK, E_ACCESSDENIED, HRESULT_FROM_WIN32(ERROR_CANCELLED)})
        {
            task                           = result[60u + i].tasks.front();
            task.finished                  = true;
            task.resultHr                  = outcome;
            task.informational.finished    = true;
            task.informational.resultHr    = outcome;
            task.informational.doneSummary = outcome == S_OK             ? L"Completed: 128 items processed."
                                             : outcome == E_ACCESSDENIED ? L"Accès refusé. Trois éléments n’ont pas pu être traités."
                                                                         : L"Opération annulée. 48 éléments ont été traités.";
            addVariant(std::format(L"{}-{}",
                                   i == 0u   ? L"compare"
                                   : i == 1u ? L"change-case"
                                   : i == 2u ? L"change-attributes"
                                             : L"file-list",
                                   outcome == S_OK             ? L"done"
                                   : outcome == E_ACCESSDENIED ? L"failed"
                                                               : L"canceled"),
                       task);
        }
    }
    using Hit = FileOperationsPopupInternal::PopupHitTest::Kind;
    for (const auto menu : {Hit::FooterOptions, Hit::FooterQueueMode, Hit::TaskCompletedMore, Hit::TaskSpeedLimit, Hit::TaskConflictMore})
    {
        const std::wstring_view name = menu == Hit::FooterOptions       ? L"menu-options"
                                       : menu == Hit::FooterQueueMode   ? L"menu-queue"
                                       : menu == Hit::TaskCompletedMore ? L"menu-completed-more"
                                       : menu == Hit::TaskSpeedLimit    ? L"menu-speed"
                                                                        : L"menu-conflict-more";
        task                         = menu == Hit::TaskCompletedMore ? partial : menu == Hit::TaskConflictMore ? result[32].tasks.front() : parallel;
        task.taskId                  = 1u;
        addVariant(name, task);
        result.back().menu = menu;
    }
    addVariant(L"renderer-failure", base);
    result.back().graphicsFailure = true;
    // Keep fixture facts coherent. Presentation fixtures are not permission to
    // manufacture impossible operation receipts or progress before admission.
    for (auto& scenario : result)
    {
        if (scenario.compactDensity)
            scenario.collapsed = true;
        for (auto& value : scenario.tasks)
        {
            if (! value.started && value.kind == Snapshot::Kind::FileOperation)
            {
                value.completedBytes     = 0u;
                value.completedItems     = 0u;
                value.itemCompletedBytes = 0u;
            }
            if (value.verificationActive)
            {
                value.completedItems                 = value.totalItems;
                value.verificationItemTotalBytes     = value.itemTotalBytes;
                value.verificationItemCompletedBytes = value.itemTotalBytes / 2u;
            }
            if (value.verificationRequested && value.finished)
            {
                value.verificationTotalBytes = value.totalBytes;
                if (value.verificationState == static_cast<uint8_t>(FileOperations::VerificationState::Verified))
                {
                    value.verificationCompletedBytes = value.totalBytes;
                    value.verifiedItemCount          = value.totalItems;
                }
                else
                {
                    value.resultHr                   = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
                    value.warningCount               = 3u;
                    value.verifiedItemCount          = 125u;
                    value.verificationCompletedBytes = value.totalBytes * 125u / 128u;
                    value.resultSummary              = L"128 fichiers copiés. 125 vérifiés ; la vérification de 3 fichiers n’a pas été terminée.";
                }
            }
            if (value.finished && value.resultHr == E_ACCESSDENIED && value.kind == Snapshot::Kind::FileOperation)
            {
                value.completedBytes     = 0u;
                value.completedItems     = 0u;
                value.itemCompletedBytes = 0u;
            }
            if (value.conflict.active)
            {
                value.conflict.destinationMetadata.isDirectory = value.conflict.bucket == static_cast<uint8_t>(FileOperations::ConflictClass::TypeMismatch);
                value.conflict.destinationMetadata.isLink      = value.conflict.bucket == static_cast<uint8_t>(FileOperations::ConflictClass::DestinationLink);
            }
            if (value.kind == Snapshot::Kind::Informational && value.finished)
            {
                auto& info                       = value.informational;
                info.scanActive                  = false;
                info.contentActive               = false;
                info.changeCaseEnumerating       = false;
                info.changeCaseRenaming          = false;
                info.changeAttributesEnumerating = false;
                info.changeAttributesApplying    = false;
                info.makeFileListCollecting      = false;
                info.makeFileListRendering       = false;
                info.makeFileListWriting         = false;
                if (SUCCEEDED(info.resultHr))
                {
                    info.changeCaseCompletedRenames     = info.changeCasePlannedRenames;
                    info.changeAttributesCompletedItems = info.changeAttributesPlannedItems;
                    info.makeFileListRenderedEntries    = info.makeFileListTotalEntries;
                    info.doneSummary                    = L"Toutes les opérations prévues ont été effectuées.";
                }
            }
        }
    }
    auto longDecision = result[32];
    longDecision.name = L"108-decision-long-french-paths";
    const std::wstring longName =
        L"Photographie panoramique de la réunion familiale — version définitive retouchée pour impression grand format et archivage patrimonial.raw";
    longDecision.tasks.front().conflict.sourcePath =
        L"C:\\Photographies\\Vacances en famille — septembre 2026\\Sélection définitive pour impression et archivage\\" + longName;
    longDecision.tasks.front().conflict.destinationPath =
        L"D:\\Sauvegardes\\Archives photographiques personnelles\\Exposition annuelle de la médiathèque\\" + longName;
    result.push_back(std::move(longDecision));
    auto& longTask                          = result.back().tasks.front();
    longTask.currentSourcePath              = longTask.conflict.sourcePath;
    longTask.currentDestinationPath         = longTask.conflict.destinationPath;
    longTask.destinationFolder              = std::filesystem::path(longTask.conflict.destinationPath).parent_path();
    longTask.conflict.sourceName            = longName;
    longTask.conflict.destinationName       = longName;
    longTask.conflict.sourceParentPath      = std::filesystem::path(longTask.conflict.sourcePath).parent_path().wstring();
    longTask.conflict.destinationParentPath = longTask.destinationFolder.wstring();
    for (const size_t index : {21u, 22u})
    {
        auto recovery = result[index];
        recovery.name = index == 21u ? L"114-recovery-known-source-menu" : L"115-recovery-unknown-source-menu";
        recovery.menu = FileOperationsPopupInternal::PopupHitTest::Kind::TaskCompletedMore;
        result.push_back(std::move(recovery));
    }
    // Exercise the longest ordinary footer vocabulary together, including the
    // known-work qualifier while another task is still discovering.
    auto discovering            = base;
    discovering.discoveryClosed = false;
    auto attention              = result[32].tasks.front();
    std::vector<Snapshot> summaryTasks{base, waiting, discovering, attention};
    for (size_t i = 0u; i < summaryTasks.size(); ++i)
        summaryTasks[i].taskId = i + 1u;
    result.push_back({L"116-footer-wrapped-summary", summaryTasks, false, true});
    result.push_back({L"117-expanded-wrapped-summary", std::move(summaryTasks)});
    auto longLeaf                   = base;
    longLeaf.currentSourcePath      = L"C:\\Photographies\\Archives familiales\\" + longName;
    longLeaf.currentDestinationPath = L"D:\\Sauvegardes\\Collection patrimoniale\\" + longName;
    result.push_back({L"118-progress-long-filename", {longLeaf}});
    const std::wstring unicodeName  = L"Photographie de la famille \U0001F468\u200D\U0001F469\u200D\U0001F467 — se\u0301lection de\u0301finitive pour "
                                      L"archivage — e\u0301te\u0301 en re\u0301gion parisienne \U0001F4F7.raw";
    longLeaf.currentSourcePath      = L"C:\\Photographies\\Archives familiales\\" + unicodeName;
    longLeaf.currentDestinationPath = L"D:\\Sauvegardes\\Collection patrimoniale\\" + unicodeName;
    result.push_back({L"119-progress-unicode-filename", {std::move(longLeaf)}});
    task                    = waiting;
    task.taskId             = 1u;
    task.started            = false;
    task.operationStartTick = 0u;
    task.waitingInQueue     = true;
    task.discoveryClosed    = false;
    task.totalBytes         = 0u;
    task.canMoveQueueUp     = false;
    task.canMoveQueueDown   = false;
    add(L"120-waiting-unknown-total", task);
    task.operation            = FILESYSTEM_MOVE;
    task.autoConcurrencyUsed  = true;
    task.autoTunedConcurrency = 4u;
    add(L"121-waiting-move-auto", task);
    // Defensive retained-row fixture: measurement must agree with painting even with stale streams.
    task.inFlightFileCount = 3u;
    for (size_t index = 0u; index < task.inFlightFileCount; ++index)
    {
        auto& stream          = task.inFlightFiles[index];
        stream.sourcePath     = std::format(L"C:\\Photographies\\Souvenir de famille numéro {}.raw", index + 1u);
        stream.totalBytes     = 1024u;
        stream.completedBytes = index == 1u ? 1024u : 512u;
        stream.lastUpdateTick = GetTickCount64();
    }
    add(L"122-waiting-retained-rows", task);
    auto limitedWaiting                                          = result[3];
    limitedWaiting.name                                          = L"133-speed-limit-waiting";
    limitedWaiting.tasks.front().desiredSpeedLimitBytesPerSecond = 999ull * 1024ull * 1024ull * 1024ull;
    result.push_back(limitedWaiting);
    auto limitedDiscovery                                          = result[5];
    limitedDiscovery.name                                          = L"134-speed-limit-discovery";
    limitedDiscovery.tasks.front().desiredSpeedLimitBytesPerSecond = limitedWaiting.tasks.front().desiredSpeedLimitBytesPerSecond;
    result.push_back(std::move(limitedDiscovery));
    limitedWaiting.name                               = L"135-speed-limit-discovery-waiting";
    limitedWaiting.tasks.front().discoveryAheadActive = true;
    limitedWaiting.tasks.front().discoveryClosed      = false;
    result.push_back(std::move(limitedWaiting));
    auto allResults      = result[30];
    allResults.name      = L"136-completed-group-all-results-expanded";
    auto unknownResult   = result[25].tasks.front();
    unknownResult.taskId = 5u;
    // An interrupted-Move restart notice stays outside the Completed group, so the group's
    // Unknown member is an ordinary finished task with an indeterminate result.
    unknownResult.interruptedMoveNotice  = false;
    unknownResult.interruptedOperationId = 0u;
    unknownResult.resultSummary          = L"Résultat inconnu. Vérifiez la source et la destination.";
    allResults.tasks.push_back(std::move(unknownResult));
    result.push_back(allResults);
    allResults.name          = L"137-completed-group-all-results-collapsed";
    allResults.groupExpanded = false;
    result.push_back(std::move(allResults));
    auto maximumLimit                                          = result[10];
    maximumLimit.name                                          = L"138-speed-limit-maximum";
    maximumLimit.tasks.front().desiredSpeedLimitBytesPerSecond = (std::numeric_limits<uint64_t>::max)();
    result.push_back(std::move(maximumLimit));
    // Synthetic text-fit stress: the stable engine does not keep an elapsed Preparing status after operation admission.
    auto preparingElapsed                             = result[1];
    preparingElapsed.name                             = L"139-synthetic-preparing-elapsed-chip";
    preparingElapsed.tasks.front().operationStartTick = GetTickCount64() - 3200u;
    result.push_back(std::move(preparingElapsed));
    // Closed discovery with neither byte nor item denominator must remain indeterminate.
    task                    = base;
    task.totalBytes         = 0u;
    task.completedBytes     = 0u;
    task.totalItems         = 0u;
    task.plannedItems       = 0u;
    task.completedItems     = 0u;
    task.itemTotalBytes     = 0u;
    task.itemCompletedBytes = 0u;
    task.inFlightFileCount  = 0u;
    add(L"140-closed-zero-byte-unknown-items", task);
    // Delete progress uses the closed work denominator, including exact zero progress.
    task                    = base;
    task.operation          = FILESYSTEM_DELETE;
    task.totalBytes         = 0u;
    task.completedBytes     = 0u;
    task.totalItems         = 1u;
    task.plannedItems       = 1u;
    task.completedItems     = 0u;
    task.itemTotalBytes     = 0u;
    task.itemCompletedBytes = 0u;
    add(L"141-delete-closed-known-items-zero", task);
    task.totalBytes   = base.totalBytes;
    task.totalItems   = base.totalItems;
    task.plannedItems = base.plannedItems;
    add(L"142-delete-closed-known-bytes-zero", task);
    task.totalBytes   = 0u;
    task.totalItems   = 0u;
    task.plannedItems = 0u;
    add(L"143-delete-closed-unknown", task);
    // Mixed terminal Move: every independent publication, source, and verification count must be painted.
    task                                  = done;
    task.operation                        = FILESYSTEM_MOVE;
    task.resultHr                         = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    task.hasIndeterminateResult           = true;
    task.resultSummary                    = L"Déplacement partiel : vérifiez les éléments publiés, les sources et la vérification.";
    task.totalItems                       = 10u;
    task.plannedItems                     = 10u;
    task.completedItems                   = 9u;
    task.publishedItemCount               = 9u;
    task.unknownPublicationItemCount      = 1u;
    task.removedSourceCount               = 7u;
    task.retainedSourceCount              = 2u;
    task.unknownSourceCount               = 1u;
    task.verificationRequested            = true;
    task.verificationState                = static_cast<uint8_t>(FileOperations::VerificationState::Failed);
    task.verifiedItemCount                = 6u;
    task.verificationProblemItemCount     = 3u;
    task.verificationFailedItemCount      = 1u;
    task.verificationUnavailableItemCount = 1u;
    task.verificationCanceledItemCount    = 1u;
    add(L"144-terminal-move-mixed-facts", task);
    // Direct Issues follows retained projectable rows, independently of the
    // warning/error counters. The two diagnostic cards keep destination recovery
    // in More; the clean Delete has only Dismiss.
    task                      = done;
    task.warningCount         = 2u;
    task.errorCount           = 1u;
    task.resultHr             = HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
    task.resultSummary        = L"Trois problèmes nécessitent une vérification.";
    task.hasProjectableIssues = true;
    add(L"145-direct-issues-row", task);
    task.hasProjectableIssues = false;
    add(L"146-diagnostic-counters-no-issues-row", task);
    task           = done;
    task.operation = FILESYSTEM_DELETE;
    task.currentDestinationPath.clear();
    task.destinationFolder.clear();
    task.destinationPluginId.clear();
    task.destinationPluginShortId.clear();
    task.resultSummary = L"Suppression terminée.";
    add(L"147-clean-no-issues", task);
    task = result[60].tasks.front();
    task.informational.scanCurrentRelative.clear();
    add(L"148-compare-scan-placeholder", task);
    // I35 R9.4 (P3): a Delete canceled before its root crossed the mutation gate states one fact.
    task           = done;
    task.operation = FILESYSTEM_DELETE;
    task.currentDestinationPath.clear();
    task.destinationFolder.clear();
    task.destinationPluginId.clear();
    task.destinationPluginShortId.clear();
    task.resultHr                    = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    task.totalBytes                  = 0u;
    task.completedBytes              = 0u;
    task.totalItems                  = 1u;
    task.plannedItems                = 1u;
    task.completedItems              = 0u;
    task.retainedSourceCount         = 1u;
    task.canceledRetainedSourceCount = 1u;
    task.startedMutation             = false;
    task.completedSummaryProjected   = true;
    task.resultSummary               = LoadStringResource(nullptr, IDS_FILEOPS_RESULT_NOT_STARTED);
    add(L"149-delete-not-started", task);
    // I35 R7.3: the by-name consents of a location without file identities (Delete by name).
    for (size_t i = kFirstByNameClass; i < names.size(); ++i)
    {
        add(std::format(L"{:03}-decision-{}", 150u + (i - kFirstByNameClass), names[i]), makeDecision(i));
    }
    // I35 (owner report 2026-10-13): finished Verify On Copies whose verification did not cover every byte (a
    // destination without an exact proof finishes with a warning). The verification half stays visible beside the full
    // transfer half, and a compact row names the share verified instead of 100%.
    task                                  = done;
    task.resultHr                         = S_FALSE;
    task.warningCount                     = 1u;
    task.verificationRequested            = true;
    task.verificationState                = static_cast<uint8_t>(FileOperations::VerificationState::Unavailable);
    task.verificationTotalBytes           = task.totalBytes;
    task.verificationCompletedBytes       = 0u;
    task.verifiedItemCount                = 0u;
    task.verificationProblemItemCount     = 128u;
    task.verificationUnavailableItemCount = 128u;
    add(L"152-completed-not-verified", task);
    auto notVerified                                 = task;
    notVerified.taskId                               = 3u;
    auto quarterVerified                             = task;
    quarterVerified.taskId                           = 2u;
    quarterVerified.verificationCompletedBytes       = task.totalBytes / 4u; // exactly a quarter: the label reads 25%
    quarterVerified.verifiedItemCount                = 32u;
    quarterVerified.verificationProblemItemCount     = 96u;
    quarterVerified.verificationUnavailableItemCount = 96u;
    auto allVerified                                 = done;
    allVerified.verificationRequested                = true;
    allVerified.verificationState                    = static_cast<uint8_t>(FileOperations::VerificationState::Verified);
    allVerified.verificationTotalBytes               = allVerified.totalBytes;
    allVerified.verificationCompletedBytes           = allVerified.totalBytes;
    allVerified.verifiedItemCount                    = 128u;
    result.push_back({L"153-compact-verification-shares", {allVerified, quarterVerified, notVerified}, true});
    return result;
}

[[nodiscard]] bool CaptureFileOpsGalleryInteraction(SelfTestState& state,
                                                    const std::filesystem::path& output,
                                                    std::ofstream& manifest,
                                                    const std::vector<FileOpsGalleryMonitor>& monitors)
{
    using namespace FileOperationsPopupInternal;
    if (RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread())
        return false;
    auto* folder                = TryGetFolderWindow(state.mainWindow);
    auto* operations            = TryGetFileOps(folder);
    auto theme                  = ResolveAppTheme(ThemeMode::Dark, L"FileOperationsGallery");
    theme.reducedMotionOverride = true;
    folder->ApplyTheme(theme);
    POINT previousCursor{};
    if (! GetCursorPos(&previousCursor))
        return false;
    auto restoreCursor   = wil::scope_exit([&] { SetCursorPos(previousCursor.x, previousCursor.y); });
    auto lifetime        = std::make_shared<int>(0);
    const auto scenarios = BuildFileOpsVisualScenarios();
    std::ofstream interactionEvidence(output / L"interaction.tsv", std::ios::binary);
    interactionEvidence << "display\tdpi\tfocusedAutomationId\tmouseCaptureOwned\n";
    wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, lifetime));
    if (! popup)
        return false;
    auto* renderer = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA));
    renderer->DebugSetVisualScenario(scenarios[32].tasks, false, true);
    RedSalamander::TestSupport::MessagePumpWaitOptions ready{};
    ready.timeout = std::chrono::milliseconds(1500);
    std::ofstream dpiFocusEvidence(output / L"dpi-focus.tsv", std::ios::binary);
    dpiFocusEvidence << "display\tbeforeDpi\tafterDpi\tbeforeAutomationId\tafterAutomationId\thostedHasFocus\tnativeFocusOwned\n";
    bool observed144To96        = false;
    bool observed96To144        = false;
    const auto nativeFocusOwned = [&]() noexcept
    {
        const HWND focus = GetFocus();
        return focus && (focus == popup.get() || IsChild(popup.get(), focus) != FALSE);
    };
    const auto verifyFocusAfterMove = [&](size_t display, const std::wstring& beforeFocus, UINT beforeDpi, bool recordActionDpi = true)
    {
        PopupLayoutDebugSnapshot moved{};
        moved.taskId       = 1u;
        const auto settled = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            MONITORINFOEXW actual{};
            actual.cbSize = sizeof(actual);
            return GetMonitorInfoW(MonitorFromWindow(popup.get(), MONITOR_DEFAULTTONEAREST), &actual) &&
                   std::wstring_view(actual.szDevice) == monitors[display].device && DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), moved) &&
                   moved.hostedFocusedAutomationId == beforeFocus && moved.hostedFocusedHasFocus && nativeFocusOwned() && GetForegroundWindow() == popup.get();
        },
            ready);
        const UINT afterDpi = GetDpiForWindow(popup.get());
        if (recordActionDpi)
            dpiFocusEvidence << display + 1u << '\t' << beforeDpi << '\t' << afterDpi << '\t' << Common::Strings::Utf8FromUtf16ReplacingInvalid(beforeFocus)
                             << '\t' << Common::Strings::Utf8FromUtf16ReplacingInvalid(moved.hostedFocusedAutomationId) << '\t' << moved.hostedFocusedHasFocus
                             << '\t' << nativeFocusOwned() << '\n';
        if (! settled.conditionMet)
            Fail(L"The same popup lost its exact focused control or native focus during a physical monitor move.");
        if (settled.conditionMet && recordActionDpi)
        {
            observed144To96 = observed144To96 || (beforeDpi == 144u && afterDpi == 96u);
            observed96To144 = observed96To144 || (beforeDpi == 96u && afterDpi == 144u);
        }
        return settled.conditionMet;
    };
    for (size_t display = 0u; display < monitors.size(); ++display)
    {
        PopupLayoutDebugSnapshot beforeMove{};
        beforeMove.taskId = 1u;
        if (display > 0u && (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), beforeMove) || beforeMove.hostedFocusedAutomationId.empty()))
            return false;
        const UINT beforeMoveDpi = GetDpiForWindow(popup.get());
        PlaceFileOpsGalleryWindow(popup.get(), monitors[display], 760, 560);
        if (display > 0u &&
            (! verifyFocusAfterMove(display, beforeMove.hostedFocusedAutomationId, beforeMoveDpi) ||
             ! SaveFileOpsGallerySurface(popup.get(), output, manifest, L"114-focused-action-dpi-hop", display, L"same-HWND-same-focused-action-DPI")))
            return false;
        if (! RedSalamander::TestSupport::DirectedSelfTestInputWarning::TryActivateOwnedWindow(popup.get()))
        {
            Fail(L"Interactive gallery could not acquire its own foreground window; no input was sent.");
            return false;
        }
        const auto key = [&](WORD virtualKey, bool shift)
        {
            if (GetForegroundWindow() != popup.get() || ! nativeFocusOwned())
            {
                Fail(L"Interactive gallery lost its foreground window before keyboard input; no input was sent.");
                return false;
            }
            std::array<INPUT, 4> input{};
            const size_t offset = shift ? 1u : 0u;
            for (auto& item : input)
                item.type = INPUT_KEYBOARD;
            input[0].ki.wVk               = VK_SHIFT;
            input[offset].ki.wVk          = virtualKey;
            input[offset + 1u].ki.wVk     = virtualKey;
            input[offset + 1u].ki.dwFlags = KEYEVENTF_KEYUP;
            input[3].ki.wVk               = VK_SHIFT;
            input[3].ki.dwFlags           = KEYEVENTF_KEYUP;
            const UINT count              = shift ? 4u : 2u;
            const UINT sent               = SendInput(count, input.data(), sizeof(INPUT));
            if (sent == count)
                return true;
            // A short SendInput can stop between a down/up pair. Best-effort
            // release only keys left down by this batch, without reactivating.
            if (GetForegroundWindow() == popup.get() && nativeFocusOwned())
            {
                std::array<INPUT, 2> release{};
                UINT releases = 0u;
                if (sent == offset + 1u)
                {
                    release[releases].type         = INPUT_KEYBOARD;
                    release[releases].ki.wVk       = virtualKey;
                    release[releases++].ki.dwFlags = KEYEVENTF_KEYUP;
                }
                if (shift && sent > 0u)
                {
                    release[releases].type         = INPUT_KEYBOARD;
                    release[releases].ki.wVk       = VK_SHIFT;
                    release[releases++].ki.dwFlags = KEYEVENTF_KEYUP;
                }
                if (releases != 0u)
                    static_cast<void>(SendInput(releases, release.data(), sizeof(INPUT)));
            }
            Fail(std::format(L"Interactive gallery keyboard injection was partial: sent {} of {} events.", sent, count));
            return false;
        };
        if (display == 0u)
        {
            // Protected phase: actual routed text input and actual text focus.
            // Keep its native DPI receipt separate from the action-focus receipt.
            PopupLayoutDebugSnapshot actionContext{};
            actionContext.taskId = 1u;
            if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), actionContext) || actionContext.hostedFocusedAutomationId.empty())
                return false;
            const std::wstring actionFocus = actionContext.hostedFocusedAutomationId;
            const auto longPaths           = std::ranges::find(scenarios, std::wstring(L"27-long-unicode-paths"), &FileOpsVisualScenario::name);
            if (longPaths == scenarios.end() || longPaths->tasks.empty())
                return false;
            const auto& task         = longPaths->tasks.front();
            const std::wstring path  = task.currentSourcePath;
            const std::wstring rowId = std::format(L"FileOperations.Task.{}.Path.Source.0", task.taskId);
            const HWND detailHwnd    = popup.get();
            renderer->DebugSetVisualScenario(longPaths->tasks, false, true);
            PlaceFileOpsGalleryWindow(detailHwnd, monitors.front(), 480, 720);
            RedrawWindow(detailHwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            renderer->DebugOpenPathDetail(rowId);
            RedrawWindow(detailHwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            PopupLayoutDebugSnapshot detailFocus{};
            detailFocus.taskId   = task.taskId;
            const auto textReady = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                const auto* field = renderer->DebugPathDetailField();
                return field && field->IsReadOnly() && field->IsMultiline() && field->GetText() == path &&
                       DebugGetFileOperationsPopupLayoutSnapshot(detailHwnd, detailFocus) && detailFocus.hostedFocusedHasFocus &&
                       detailFocus.hostedFocusedAutomationId == L"FileOperations.PathDetail.Text" && nativeFocusOwned() && GetForegroundWindow() == detailHwnd;
            },
                ready);
            DxUi::TextFieldDebugMultilineState top{};
            if (! textReady.conditionMet || path.size() <= 4096u || ! path.starts_with(L"\\\\") || path.find(L"école") == std::wstring::npos ||
                path.find(L"שלום") == std::wstring::npos || path.find(L"\U0001F4F7") == std::wstring::npos || ! renderer->DebugPathDetailMultilineState(top) ||
                ! top.canScrollVertically || top.firstVisibleLine != 0u)
            {
                Fail(L"Protected long-path detail must own actual text/native focus and expose the exact read-only multiline Unicode document.");
                return false;
            }
            if (! key(VK_NEXT, false))
                return false;
            DxUi::TextFieldDebugMultilineState scrolled{};
            const auto pageReady = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                const auto* field = renderer->DebugPathDetailField();
                return field && field->IsReadOnly() && field->IsMultiline() && field->GetText() == path && renderer->DebugPathDetailMultilineState(scrolled) &&
                       scrolled.firstVisibleLine > top.firstVisibleLine && scrolled.firstVisibleLine + scrolled.visibleLineCount <= scrolled.totalLineCount &&
                       DebugGetFileOperationsPopupLayoutSnapshot(detailHwnd, detailFocus) && detailFocus.hostedFocusedHasFocus &&
                       detailFocus.hostedFocusedAutomationId == L"FileOperations.PathDetail.Text" && nativeFocusOwned() && GetForegroundWindow() == detailHwnd;
            },
                ready);
            AppendLog(
                std::format(L"VisualGalleryPathDetailPageDown: settled={} first={}->{} visible={} total={} actualFocus='{}' native=0x{:X} foreground=0x{:X}",
                            pageReady.conditionMet,
                            top.firstVisibleLine,
                            scrolled.firstVisibleLine,
                            scrolled.visibleLineCount,
                            scrolled.totalLineCount,
                            detailFocus.hostedFocusedAutomationId,
                            reinterpret_cast<uintptr_t>(GetFocus()),
                            reinterpret_cast<uintptr_t>(GetForegroundWindow())));
            if (! pageReady.conditionMet ||
                ! SaveFileOpsGallerySurface(detailHwnd, output, manifest, L"115-path-detail-routed-page-down", 0u, L"actual-text-focus-real-routed-PageDown"))
            {
                Fail(L"Protected routed PageDown must scroll the exact read-only document while retaining actual text/native focus.");
                return false;
            }
            auto* field = renderer->DebugPathDetailField();
            if (! field)
                return false;
            // Exercise a nonempty selection, rather than only a collapsed caret.
            // Opening the workspace selects the whole path; use an interior range
            // before reflow so either endpoint being lost fails the DPI check.
            field->SetSelectionRange(5u, 19u);
            RedrawWindow(detailHwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            field = renderer->DebugPathDetailField();
            if (! field)
                return false;
            const size_t caret   = field->GetCaretIndex();
            const auto selection = field->GetSelectionRange();
            if (caret != 19u || ! selection.has_value() || selection.value() != std::pair<size_t, size_t>{5u, 19u})
                return false;
            bool text144To96 = false;
            bool text96To144 = false;
            std::ofstream textDpi(output / L"path-detail-native-dpi.tsv", std::ios::binary);
            textDpi << "display\tbeforeDpi\tafterDpi\tsameHwnd\texactTextCaretSelection\tactualTextFocus\tnativeFocusOwned\n";
            for (size_t hop = 1u; hop <= monitors.size(); ++hop)
            {
                const size_t targetDisplay = hop % monitors.size();
                const UINT beforeDpi       = GetDpiForWindow(detailHwnd);
                if (GetForegroundWindow() != detailHwnd || ! nativeFocusOwned())
                    return false;
                PlaceFileOpsGalleryWindow(detailHwnd, monitors[targetDisplay], 480, 720);
                if (! verifyFocusAfterMove(targetDisplay, L"FileOperations.PathDetail.Text", beforeDpi, false))
                    return false;
                RedrawWindow(detailHwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                const auto* moved      = renderer->DebugPathDetailField();
                const bool exact       = moved && moved->IsReadOnly() && moved->IsMultiline() && moved->GetText() == path && moved->GetCaretIndex() == caret &&
                                         moved->GetSelectionRange() == selection;
                const bool actualFocus = DebugGetFileOperationsPopupLayoutSnapshot(detailHwnd, detailFocus) &&
                                         detailFocus.hostedFocusedAutomationId == L"FileOperations.PathDetail.Text" && detailFocus.hostedFocusedHasFocus;
                const UINT afterDpi    = GetDpiForWindow(detailHwnd);
                textDpi << targetDisplay + 1u << '\t' << beforeDpi << '\t' << afterDpi << '\t' << (detailHwnd == popup.get()) << '\t' << exact << '\t'
                        << actualFocus << '\t' << nativeFocusOwned() << '\n';
                textDpi.flush();
                AppendLog(std::format(L"VisualGalleryPathDetailDpi: display={} dpi={}->{} sameHwnd={} textCaretSelection={} caret={} selection={}:{} "
                                      L"actualFocus='{}' retainedFocus={} native=0x{:X} foreground=0x{:X}",
                                      targetDisplay + 1u,
                                      beforeDpi,
                                      afterDpi,
                                      detailHwnd == popup.get(),
                                      exact,
                                      moved ? moved->GetCaretIndex() : 0u,
                                      selection.value().first,
                                      selection.value().second,
                                      detailFocus.hostedFocusedAutomationId,
                                      actualFocus,
                                      reinterpret_cast<uintptr_t>(GetFocus()),
                                      reinterpret_cast<uintptr_t>(GetForegroundWindow())));
                if (detailHwnd != popup.get() || ! exact || ! actualFocus || ! nativeFocusOwned() || GetForegroundWindow() != detailHwnd)
                {
                    Fail(L"Physical monitor moves must preserve the same HWND, exact path/caret/selection and actual hosted/native text focus.");
                    return false;
                }
                text144To96 = text144To96 || (beforeDpi == 144u && afterDpi == 96u);
                text96To144 = text96To144 || (beforeDpi == 96u && afterDpi == 144u);
                if (! SaveFileOpsGallerySurface(detailHwnd,
                                                output,
                                                manifest,
                                                L"116-path-detail-text-focus-dpi-hop",
                                                targetDisplay,
                                                L"same-HWND-exact-text-selection-actual-text-native-focus"))
                    return false;
            }
            std::ofstream textQualification(output / L"path-detail-native-dpi-qualification.tsv", std::ios::binary);
            textQualification << "status\tobserved144To96\tobserved96To144\n"
                              << (text144To96 && text96To144 ? "complete" : "pending") << '\t' << text144To96 << '\t' << text96To144 << '\n';
            textQualification.flush();
            if (! textDpi.good() || ! textQualification.good() || ! key(VK_ESCAPE, false))
                return false;
            const auto rowReady = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                return ! renderer->DebugPathDetailField() && DebugGetFileOperationsPopupLayoutSnapshot(detailHwnd, detailFocus) &&
                       detailFocus.hostedFocusedAutomationId == rowId && detailFocus.hostedFocusedHasFocus && nativeFocusOwned() &&
                       GetForegroundWindow() == detailHwnd;
            },
                ready);
            if (! rowReady.conditionMet)
            {
                Fail(L"Closing protected path detail must restore its exact originating row and owned native focus.");
                return false;
            }
            renderer->DebugSetVisualScenario(scenarios[32].tasks, false, true);
            PlaceFileOpsGalleryWindow(detailHwnd, monitors.front(), 760, 560);
            // Static placement deliberately lowers the window. The text-focus
            // witnesses above are already complete; restore the owned window's
            // z-order once before the pointer phase, without rescuing lost focus.
            if (GetForegroundWindow() != detailHwnd || ! nativeFocusOwned() ||
                ! RedSalamander::TestSupport::DirectedSelfTestInputWarning::TryActivateOwnedWindow(detailHwnd))
            {
                Fail(L"Protected action restoration lost native ownership; no further input was sent.");
                return false;
            }
            RedrawWindow(detailHwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
            if (! renderer->DebugSetHostedButtonFocusForSelfTest(detailHwnd, actionFocus))
                return false;
            const auto actionRestored = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                return DebugGetFileOperationsPopupLayoutSnapshot(detailHwnd, actionContext) && actionContext.hostedFocusedHasFocus &&
                       actionContext.hostedFocusedAutomationId == actionFocus && nativeFocusOwned() && GetForegroundWindow() == detailHwnd;
            },
                ready);
            if (! actionRestored.conditionMet)
            {
                Fail(L"Protected path detail must restore the original action context before the existing interaction checks.");
                return false;
            }
        }
        PopupLayoutDebugSnapshot focused{};
        focused.taskId = 1u;
        if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), focused))
            return false;
        const std::wstring focusBeforeTab = focused.hostedFocusedAutomationId;
        const uint64_t nativeTabsBefore   = focused.hostedNativeTabCount;
        const uint64_t focusChangesBefore = focused.hostedFocusChangeCount;
        if (! key(VK_TAB, false))
            return false;
        const auto focusReady = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            // A retained default already has an identity before SendInput is
            // dispatched. Require the actual Tab transition before reversing it.
            return DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), focused) && focused.hostedFocusedHasFocus &&
                   ! focused.hostedFocusedAutomationId.empty() && focused.hostedFocusedAutomationId != focusBeforeTab && nativeFocusOwned() &&
                   GetForegroundWindow() == popup.get();
        },
            ready);
        AppendLog(std::format(L"VisualGalleryKeyboard: display={} direction=forward before='{}' after='{}' retainedFocus={} nativeFocus=0x{:X} "
                              L"foreground=0x{:X} popup=0x{:X} settled={}",
                              display + 1u,
                              focusBeforeTab,
                              focused.hostedFocusedAutomationId,
                              focused.hostedFocusedHasFocus,
                              reinterpret_cast<uintptr_t>(GetFocus()),
                              reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                              reinterpret_cast<uintptr_t>(popup.get()),
                              focusReady.conditionMet));
        AppendLog(
            std::format(L"VisualGalleryKeyboardDispatch: nativeTabs={}->{} focusChanges={}->{} previousTarget='{}' lastTarget='{}' messages={} elapsedMs={}",
                        nativeTabsBefore,
                        focused.hostedNativeTabCount,
                        focusChangesBefore,
                        focused.hostedFocusChangeCount,
                        focused.hostedPreviousFocusTargetId,
                        focused.hostedLastFocusTargetId,
                        focusReady.dispatchedMessageCount,
                        focusReady.elapsed.count()));
        if (! focusReady.conditionMet)
            Fail(L"Interactive gallery Tab did not move actual focus to a different retained control.");
        if (! focusReady.conditionMet ||
            ! SaveFileOpsGallerySurface(popup.get(), output, manifest, L"109-keyboard-tab-focus", display, L"real-foreground-keyboard-input"))
            return false;
        const std::wstring firstFocus = focused.hostedFocusedAutomationId;
        if (! key(VK_TAB, true))
            return false;
        const auto reverseReady = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            return DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), focused) && ! focused.hostedFocusedAutomationId.empty() &&
                   focused.hostedFocusedAutomationId != firstFocus && focused.hostedFocusedHasFocus && nativeFocusOwned() &&
                   GetForegroundWindow() == popup.get();
        },
            ready);
        AppendLog(std::format(L"VisualGalleryKeyboard: display={} direction=reverse before='{}' after='{}' retainedFocus={} nativeFocus=0x{:X} "
                              L"foreground=0x{:X} popup=0x{:X} settled={}",
                              display + 1u,
                              firstFocus,
                              focused.hostedFocusedAutomationId,
                              focused.hostedFocusedHasFocus,
                              reinterpret_cast<uintptr_t>(GetFocus()),
                              reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                              reinterpret_cast<uintptr_t>(popup.get()),
                              reverseReady.conditionMet));
        if (! reverseReady.conditionMet)
            Fail(L"Interactive gallery Shift+Tab did not move actual focus to a different retained control.");
        if (! reverseReady.conditionMet ||
            ! SaveFileOpsGallerySurface(popup.get(), output, manifest, L"110-keyboard-shift-tab-focus", display, L"real-foreground-keyboard-input"))
            return false;
        // Window placement and the synthetic scenario can queue a later auto-size.
        // Wait for real layout quiescence before recording physical pointer targets.
        // Do not reactivate or freeze the product's sizing/animation policy.
        RECT stableWindow{};
        D2D1_RECT_F stableButton{};
        unsigned int stableSamples = 0u;
        bool pointerOwnershipLost  = false;
        PopupLayoutDebugSnapshot pointerLayout{};
        pointerLayout.taskId            = 1u;
        const auto pointerGeometryReady = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            if (GetForegroundWindow() != popup.get() || ! nativeFocusOwned())
            {
                pointerOwnershipLost = true;
                return true; // End this attempt; later focus recovery cannot admit input.
            }
            if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), pointerLayout) || pointerLayout.layoutMotionActive)
            {
                stableSamples = 0u;
                return false;
            }
            RECT window{};
            if (! GetWindowRect(popup.get(), &window))
            {
                stableSamples = 0u;
                return false;
            }
            for (const auto& button : renderer->DebugVisualButtons())
                if (button.hit.kind == PopupHitTest::Kind::FooterOptions)
                {
                    const auto current = button.bounds;
                    const bool same    = window.left == stableWindow.left && window.top == stableWindow.top && window.right == stableWindow.right &&
                                         window.bottom == stableWindow.bottom && current.left == stableButton.left && current.top == stableButton.top &&
                                         current.right == stableButton.right && current.bottom == stableButton.bottom;
                    stableSamples      = same ? stableSamples + 1u : 1u;
                    stableWindow       = window;
                    stableButton       = current;
                    return stableSamples >= 3u;
                }
            stableSamples = 0u;
            return false;
        },
            ready);
        AppendLog(std::format(
            L"VisualGalleryPointerGeometry: display={} settled={} ownershipLost={} samples={} layoutMotion={} window=({},{},{},{}) button=({},{},{},{})",
            display + 1u,
            pointerGeometryReady.conditionMet,
            pointerOwnershipLost,
            stableSamples,
            pointerLayout.layoutMotionActive,
            stableWindow.left,
            stableWindow.top,
            stableWindow.right,
            stableWindow.bottom,
            stableButton.left,
            stableButton.top,
            stableButton.right,
            stableButton.bottom));
        if (pointerOwnershipLost || ! pointerGeometryReady.conditionMet || stableSamples < 3u)
        {
            Fail(L"Interactive gallery pointer geometry did not settle while native focus remained owned; no pointer input was sent.");
            return false;
        }
        std::optional<PopupButton> target;
        for (const auto& button : renderer->DebugVisualButtons())
            if (button.hit.kind == PopupHitTest::Kind::FooterOptions)
                target = button;
        if (! target.has_value())
            return false;
        const auto bounds = target.value().bounds;
        // PopupButton bounds are physical client pixels. SyncHostedControls alone
        // converts them to DIPs; applying DPI again would target another window.
        POINT point{static_cast<LONG>((bounds.left + bounds.right) * 0.5f), static_cast<LONG>((bounds.top + bounds.bottom) * 0.5f)};
        RECT clientBounds{};
        RECT windowBounds{};
        const bool clientKnown    = GetClientRect(popup.get(), &clientBounds) != FALSE;
        const bool windowKnown    = GetWindowRect(popup.get(), &windowBounds) != FALSE;
        const bool pointKnown     = ClientToScreen(popup.get(), &point) != FALSE;
        const HWND pointerTarget  = WindowFromPoint(point);
        DWORD pointerProcess      = 0u;
        const DWORD pointerThread = GetWindowThreadProcessId(pointerTarget, &pointerProcess);
        std::array<wchar_t, 256> pointerClass{};
        GetClassNameW(pointerTarget, pointerClass.data(), static_cast<int>(pointerClass.size()));
        AppendLog(std::format(L"VisualGalleryPointer: display={} dpi={} button=({},{},{},{}) clientKnown={} client=({},{},{},{}) "
                              L"windowKnown={} window=({},{},{},{}) pointKnown={} screen=({},{}) hit=0x{:X} process={} thread={} class='{}' "
                              L"warning={} popup=0x{:X} focus=0x{:X} foreground=0x{:X}",
                              display + 1u,
                              GetDpiForWindow(popup.get()),
                              bounds.left,
                              bounds.top,
                              bounds.right,
                              bounds.bottom,
                              clientKnown,
                              clientBounds.left,
                              clientBounds.top,
                              clientBounds.right,
                              clientBounds.bottom,
                              windowKnown,
                              windowBounds.left,
                              windowBounds.top,
                              windowBounds.right,
                              windowBounds.bottom,
                              pointKnown,
                              point.x,
                              point.y,
                              reinterpret_cast<uintptr_t>(pointerTarget),
                              pointerProcess,
                              pointerThread,
                              pointerClass.data(),
                              RedSalamander::TestSupport::DirectedSelfTestInputWarning::IsWarningWindow(pointerTarget),
                              reinterpret_cast<uintptr_t>(popup.get()),
                              reinterpret_cast<uintptr_t>(GetFocus()),
                              reinterpret_cast<uintptr_t>(GetForegroundWindow())));
        if (! clientKnown || ! windowKnown || ! pointKnown || GetForegroundWindow() != popup.get() || ! nativeFocusOwned() ||
            (pointerTarget != popup.get() && ! IsChild(popup.get(), pointerTarget)))
        {
            Fail(L"Interactive gallery pointer target is not owned by the popup; no pointer input was sent.");
            return false;
        }
        const auto tracePointer = [&](std::wstring_view phase, POINT requested, POINT before)
        {
            POINT cursor{};
            const bool cursorKnown      = GetCursorPos(&cursor) != FALSE;
            POINT client                = cursor;
            const bool clientPointKnown = cursorKnown && ScreenToClient(popup.get(), &client) != FALSE;
            const auto state            = renderer->DebugPointerSnapshotForSelfTest(client, PopupHitTest::Kind::FooterOptions);
            AppendLog(std::format(L"VisualGalleryPointerState: phase='{}' requested=({},{}) before=({},{}) cursorKnown={} cursor=({},{}) "
                                  L"clientKnown={} client=({},{}) mouseMoves={} lastMove=({},{}) host=0x{:X} hostDpi={} "
                                  L"targetDip=({},{},{},{}) visible={} enabled={} hovered={} hit='{}' hitDip=({},{},{},{}) "
                                  L"hitVisible={} hitEnabled={} capture=0x{:X} captured='{}' focus=0x{:X} foreground=0x{:X}",
                                  phase,
                                  requested.x,
                                  requested.y,
                                  before.x,
                                  before.y,
                                  cursorKnown,
                                  cursor.x,
                                  cursor.y,
                                  clientPointKnown,
                                  client.x,
                                  client.y,
                                  state.mouseMoveCount,
                                  state.lastMouseMove.x,
                                  state.lastMouseMove.y,
                                  reinterpret_cast<uintptr_t>(state.hostHwnd),
                                  state.hostDpi,
                                  state.targetBoundsDip.left,
                                  state.targetBoundsDip.top,
                                  state.targetBoundsDip.right,
                                  state.targetBoundsDip.bottom,
                                  state.targetVisible,
                                  state.targetEnabled,
                                  state.targetHovered,
                                  state.hitAutomationId,
                                  state.hitBoundsDip.left,
                                  state.hitBoundsDip.top,
                                  state.hitBoundsDip.right,
                                  state.hitBoundsDip.bottom,
                                  state.hitVisible,
                                  state.hitEnabled,
                                  reinterpret_cast<uintptr_t>(state.nativeCapture),
                                  state.capturedAutomationId,
                                  reinterpret_cast<uintptr_t>(GetFocus()),
                                  reinterpret_cast<uintptr_t>(GetForegroundWindow())));
        };
        POINT pointerBefore{};
        if (! GetCursorPos(&pointerBefore) || bounds.right - bounds.left < 4.0f || bounds.bottom - bounds.top < 4.0f)
        {
            Fail(L"Interactive gallery cannot establish two distinct interior pointer positions; no pointer input was sent.");
            return false;
        }
        POINT stagingPoint = point;
        --stagingPoint.x;
        if (stagingPoint.x == pointerBefore.x && stagingPoint.y == pointerBefore.y)
            stagingPoint.x = point.x + 1;
        const auto positionOwnedPointer = [&](POINT requested)
        {
            const HWND hit = WindowFromPoint(requested);
            if (GetForegroundWindow() != popup.get() || ! nativeFocusOwned() || (hit != popup.get() && ! IsChild(popup.get(), hit)))
                return false;
            if (! SetCursorPos(requested.x, requested.y))
                return false;
            POINT actual{};
            return GetCursorPos(&actual) && actual.x == requested.x && actual.y == requested.y;
        };
        const auto movesBefore = renderer->DebugNativeMouseMoveCountForSelfTest();
        if (! positionOwnedPointer(stagingPoint))
        {
            tracePointer(L"staging-refused", stagingPoint, pointerBefore);
            Fail(L"Interactive gallery could not establish its owned interior pointer position.");
            return false;
        }
        tracePointer(L"staging-positioned", stagingPoint, pointerBefore);
        const auto stagingMoved =
            RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept { return renderer->DebugNativeMouseMoveCountForSelfTest() > movesBefore; }, ready);
        tracePointer(L"staging-pumped", stagingPoint, pointerBefore);
        const auto movesBeforeTarget = renderer->DebugNativeMouseMoveCountForSelfTest();
        if (! stagingMoved.conditionMet || ! positionOwnedPointer(point))
        {
            tracePointer(L"final-position-refused", point, stagingPoint);
            Fail(L"Interactive gallery did not receive its real owned pointer movement; no mouse-down was sent.");
            return false;
        }
        tracePointer(L"target-positioned", point, stagingPoint);
        const auto hovered = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return renderer->DebugNativeMouseMoveCountForSelfTest() > movesBeforeTarget &&
                   renderer->DebugVisualButtonHovered(PopupHitTest::Kind::FooterOptions);
        }, ready);
        tracePointer(hovered.conditionMet ? L"hovered" : L"hover-timeout", point, stagingPoint);
        if (! hovered.conditionMet)
        {
            Fail(L"Interactive gallery did not reach the target button's real hover state.");
            return false;
        }
        if (! SaveFileOpsGallerySurface(popup.get(), output, manifest, L"111-pointer-hover", display, L"real-pointer-hover"))
            return false;
        POINT currentPointer{};
        const HWND currentTarget     = WindowFromPoint(point);
        const bool cursorKnown       = GetCursorPos(&currentPointer) != FALSE;
        POINT currentClient          = currentPointer;
        const bool cursorClientKnown = ScreenToClient(popup.get(), &currentClient) != FALSE;
        std::optional<PopupButton> clickTarget;
        for (const auto& button : renderer->DebugVisualButtons())
            if (button.hit.kind == PopupHitTest::Kind::FooterOptions)
                clickTarget = button;
        const float currentClientX   = static_cast<float>(currentClient.x);
        const float currentClientY   = static_cast<float>(currentClient.y);
        const bool stillInsideTarget = clickTarget.has_value() && cursorClientKnown && currentClientX >= clickTarget.value().bounds.left &&
                                       currentClientX < clickTarget.value().bounds.right && currentClientY >= clickTarget.value().bounds.top &&
                                       currentClientY < clickTarget.value().bounds.bottom;
        if (GetForegroundWindow() != popup.get() || ! nativeFocusOwned() || ! cursorKnown || currentPointer.x != point.x || currentPointer.y != point.y ||
            ! stillInsideTarget || ! renderer->DebugVisualButtonHovered(PopupHitTest::Kind::FooterOptions) ||
            (currentTarget != popup.get() && ! IsChild(popup.get(), currentTarget)))
        {
            Fail(L"Interactive gallery lost its owned pointer or focus before mouse-down; no pointer input was sent.");
            return false;
        }
        INPUT mouse{};
        mouse.type       = INPUT_MOUSE;
        mouse.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
        if (SendInput(1u, &mouse, sizeof(mouse)) != 1u)
            return false;
        bool mouseReleased        = false;
        const auto releasePointer = [&]() noexcept
        {
            SetCursorPos(monitors[display].work.left + 2, monitors[display].work.top + 2);
            INPUT release{};
            release.type       = INPUT_MOUSE;
            release.mi.dwFlags = MOUSEEVENTF_LEFTUP;
            return SendInput(1u, &release, sizeof(release)) == 1u;
        };
        auto releaseMouse  = wil::scope_exit([&]() noexcept
        {
            if (! mouseReleased)
                static_cast<void>(releasePointer());
        });
        const auto pressed = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            const HWND capture = GetCapture();
            return capture && (capture == popup.get() || IsChild(popup.get(), capture));
        },
            ready);
        interactionEvidence << display + 1u << '\t' << GetDpiForWindow(popup.get()) << '\t'
                            << Common::Strings::Utf8FromUtf16ReplacingInvalid(focused.hostedFocusedAutomationId) << '\t' << pressed.conditionMet << '\n';
        if (! pressed.conditionMet ||
            ! SaveFileOpsGallerySurface(popup.get(), output, manifest, L"112-pointer-pressed", display, L"real-pointer-down-capture-owned"))
            return false;
        if (! releasePointer())
            return false;
        const auto released = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept { return GetCapture() == nullptr && (GetAsyncKeyState(VK_LBUTTON) & 0x8000) == 0; }, ready);
        if (! released.conditionMet)
            return false;
        mouseReleased = true;
    }
    PopupLayoutDebugSnapshot beforeReturn{};
    beforeReturn.taskId = 1u;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), beforeReturn) || beforeReturn.hostedFocusedAutomationId.empty())
        return false;
    const UINT beforeReturnDpi = GetDpiForWindow(popup.get());
    PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), 760, 560);
    if (! verifyFocusAfterMove(0u, beforeReturn.hostedFocusedAutomationId, beforeReturnDpi) ||
        ! SaveFileOpsGallerySurface(popup.get(), output, manifest, L"113-dpi-return-to-first-display", 0u, L"same-HWND-same-focused-action-round-trip-DPI"))
        return false;
    // The gallery remains usable on other monitor arrangements. I26 acceptance
    // checks this separate receipt and requires both native DPI transitions.
    std::ofstream dpiQualification(output / L"native-dpi-qualification.tsv", std::ios::binary);
    dpiQualification << "status\tobserved144To96\tobserved96To144\n"
                     << (observed144To96 && observed96To144 ? "complete" : "pending") << '\t' << observed144To96 << '\t' << observed96To144 << '\n';
    interactionEvidence.flush();
    dpiFocusEvidence.flush();
    dpiQualification.flush();
    return interactionEvidence.good() && dpiFocusEvidence.good() && dpiQualification.good();
}

struct FileOpsGalleryPromptCapture final
{
    std::filesystem::path output;
    std::ofstream* manifest                            = nullptr;
    bool succeeded                                     = false;
    size_t display                                     = 0u;
    const std::vector<FileOpsGalleryMonitor>* monitors = nullptr;
};
FileOpsGalleryPromptCapture* g_fileOpsGalleryPromptCapture = nullptr;

void CALLBACK CaptureFileOpsGalleryPrompt(HWND, UINT, UINT_PTR timerId, DWORD) noexcept
{
    const HWND prompt = GetFileOperationsSpeedLimitPromptHandle();
    if (! prompt || ! g_fileOpsGalleryPromptCapture)
        return;
    KillTimer(nullptr, timerId);
    auto& capture = *g_fileOpsGalleryPromptCapture;
    const std::array<std::wstring_view, 3> names{L"99-speed-custom", L"100-speed-invalid", L"101-speed-valid"};
    capture.succeeded = true;
    for (size_t index = 0u; index < names.size(); ++index)
    {
        if (index == 1u)
        {
            capture.succeeded = DebugSetFileOperationsSpeedLimitPromptText(L"invalid-limit") && DebugConfirmFileOperationsSpeedLimitPrompt();
        }
        else if (index == 2u)
        {
            capture.succeeded = DebugSetFileOperationsSpeedLimitPromptText(L"64MB");
        }
        RedSalamander::TestSupport::MessagePumpWaitOptions ready{};
        ready.timeout       = std::chrono::milliseconds(2500);
        const auto rendered = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            FileOperationsSpeedLimitPromptDebugSnapshot snapshot{};
            return DebugGetFileOperationsSpeedLimitPromptSnapshot(snapshot) &&
                   (index == 1u ? ! snapshot.validationText.empty() : snapshot.validationText.empty());
        },
            ready);
        capture.succeeded = capture.succeeded && rendered.conditionMet;
        RedrawWindow(prompt, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
        const auto file   = std::format(L"{}__dark__fr-FR__display{}__{}dpi.png", names[index], capture.display + 1u, GetDpiForWindow(prompt));
        capture.succeeded = capture.succeeded && SUCCEEDED(RedSalamander::TestSupport::SaveWindowScreenshot(prompt, capture.output / file));
        if (! capture.succeeded)
            break;
        *capture.manifest << Common::Strings::Utf8FromUtf16ReplacingInvalid(file) << '\t' << Common::Strings::Utf8FromUtf16ReplacingInvalid(names[index])
                          << "\tdark\t0\t" << GetDpiForWindow(prompt) << "\tproduction-dialog-synthetic-input\n";
        if (index == 1u && capture.display == 0u && capture.monitors && capture.monitors->size() > 1u)
        {
            // Move this same live, invalid form across physical monitors and back.
            // Reuse the gallery monitor inventory; never synthesize WM_DPICHANGED or
            // change desktop settings. This static route does not take foreground.
            FileOperationsSpeedLimitPromptDebugSnapshot before{};
            capture.succeeded     = DebugSetFileOperationsSpeedLimitPromptSelection(1u, 4u) && DebugGetFileOperationsSpeedLimitPromptSnapshot(before) &&
                                    ! before.validationText.empty();
            const HWND focus      = GetFocus();
            const HWND foreground = GetForegroundWindow();
            for (size_t hop = 1u; capture.succeeded && hop <= capture.monitors->size(); ++hop)
            {
                const size_t display = hop % capture.monitors->size();
                const auto& monitor  = (*capture.monitors)[display];
                capture.succeeded =
                    SetWindowPos(prompt, nullptr, monitor.work.left + 16, monitor.work.top + 16, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE) != FALSE;
                FileOperationsSpeedLimitPromptDebugSnapshot moved{};
                RECT window{};
                MONITORINFOEXW actual{};
                actual.cbSize      = sizeof(actual);
                const auto settled = RedSalamander::TestSupport::PumpMessagesUntil(
                    [&]() noexcept
                {
                    RedrawWindow(prompt, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    return GetFileOperationsSpeedLimitPromptHandle() == prompt && DebugGetFileOperationsSpeedLimitPromptSnapshot(moved) &&
                           GetWindowRect(prompt, &window) && GetMonitorInfoW(MonitorFromWindow(prompt, MONITOR_DEFAULTTONEAREST), &actual) &&
                           std::wstring_view(actual.szDevice) == monitor.device && window.left >= actual.rcWork.left && window.top >= actual.rcWork.top &&
                           window.right <= actual.rcWork.right && window.bottom <= actual.rcWork.bottom;
                },
                    ready);
                capture.succeeded = capture.succeeded && settled.conditionMet && moved.fieldIdentity == before.fieldIdentity && moved.text == before.text &&
                                    moved.validationText == before.validationText && moved.caretIndex == before.caretIndex &&
                                    moved.selectionAnchorIndex == before.selectionAnchorIndex && GetFocus() == focus && GetForegroundWindow() == foreground &&
                                    moved.labelBounds.bottom - moved.labelBounds.top + 0.5f >= moved.requiredLabelHeightDip &&
                                    moved.hintBounds.bottom - moved.hintBounds.top + 0.5f >= moved.requiredHintHeightDip &&
                                    moved.bodyBounds.bottom + 9.5f <= moved.okBounds.top && moved.okBounds.right + 7.5f <= moved.cancelBounds.left &&
                                    moved.cancelBounds.bottom <= moved.clientBounds.bottom + 0.5f;
                AppendLog(std::format(L"SpeedPromptNativeDpi: hop={} display={} device={} dpi={} sameField={} text={} validation={} selection={} "
                                      L"focus={} workAreaFit={} passed={}",
                                      hop,
                                      display + 1u,
                                      monitor.device,
                                      GetDpiForWindow(prompt),
                                      moved.fieldIdentity == before.fieldIdentity,
                                      moved.text == before.text,
                                      moved.validationText == before.validationText,
                                      moved.caretIndex == before.caretIndex && moved.selectionAnchorIndex == before.selectionAnchorIndex,
                                      GetFocus() == focus && GetForegroundWindow() == foreground,
                                      settled.conditionMet,
                                      capture.succeeded));
                capture.succeeded =
                    capture.succeeded && SaveFileOpsGallerySurface(prompt,
                                                                   capture.output,
                                                                   *capture.manifest,
                                                                   std::format(L"148-speed-live-validation-dpi-hop{}", hop),
                                                                   display,
                                                                   L"same-HWND-native-monitor-move; invalid French form; selection and work-area fit");
            }
            if (! capture.succeeded)
                break;
        }
    }
    static_cast<void>(DebugCancelFileOperationsSpeedLimitPrompt());
}

// Review-remediation capture deliberately inspects only its own HWND and retained tree.
// It needs neither desktop-global UIA/focus observation nor real foreground input.
[[nodiscard]] bool CaptureFileOpsDocumentation(SelfTestState& state)
{
    using namespace FileOperationsPopupInternal;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread())
        return false;
    auto* folder           = TryGetFolderWindow(state.mainWindow);
    const auto monitors    = GetFileOpsGalleryMonitors();
    const auto output      = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::FileOperations, L"documentation");
    const auto breadcrumbs = output / L"empty-breadcrumbs";
    if (! folder || monitors.empty() || output.empty() || ! SelfTest::EnsureDirectory(output) || ! SelfTest::EnsureDirectory(breadcrumbs))
        return false;

    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui.value().language.empty() && g_settings.ui.value().language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui.value().language};
    const auto previousTheme = folder->GetTheme();
    const auto restore       = wil::scope_exit([&]
    {
        static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage));
        folder->ApplyTheme(previousTheme);
    });
    if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, L"en-US"})))
        return false;
    auto theme                  = ResolveAppTheme(ThemeMode::Dark, L"Documentation");
    theme.reducedMotionOverride = true;
    folder->ApplyTheme(theme);

    // These invented, display-only identities are never opened. Private state and
    // an empty durable store keep actual tasks and personal data out of the frames.
    constexpr std::wstring_view subject           = L"Workshop";
    constexpr std::wstring_view destinationFolder = LR"(D:\Demo\Archive\Workshop)";
    constexpr std::wstring_view sourcePath        = LR"(C:\Demo\Media\Workshop\session-01.mp4)";
    constexpr std::wstring_view destinationPath   = LR"(D:\Demo\Archive\Workshop\session-01.mp4)";
    constexpr std::wstring_view issueName         = L"session-02.mp4";
    constexpr std::wstring_view issueSource       = LR"(C:\Demo\Media\Workshop\session-02.mp4)";
    constexpr std::wstring_view issueDestination  = LR"(D:\Demo\Archive\Workshop\session-02.mp4)";
    FileOperationMoveBreadcrumb::SetRootForSelfTest(breadcrumbs);
    const auto restoreBreadcrumbs = wil::scope_exit([] { FileOperationMoveBreadcrumb::ClearRootForSelfTest(); });
    Common::Settings::Settings settings{};
    settings.fileOperations.emplace();
    auto owner = std::make_unique<FolderWindow>();
    owner->SetSettings(&settings);
    FolderWindow::FileOperationState operations(*owner);
    auto lifetime = std::make_shared<int>(0);
    wil::unique_hwnd popup(FileOperationsPopup::Create(&operations, folder, state.mainWindow, lifetime));
    wil::unique_hwnd issues;
    const auto closeWindows = wil::scope_exit([&]
    {
        lifetime.reset();
        issues.reset();
        popup.reset();
    });
    auto* renderer          = popup ? reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA)) : nullptr;
    if (! renderer)
        return false;

    TaskSnapshot task{};
    task.taskId         = 0xD0C1u;
    task.operation      = FILESYSTEM_COPY;
    task.lifecyclePhase = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Running);
    task.started = task.discoveryClosed = task.discoveredBytesKnown = task.hasProgressCallbacks = true;
    task.selection                                                                              = {std::wstring(subject), 1u, 0u, 1u, true};
    task.currentSourcePath                                                                      = sourcePath;
    task.currentDestinationPath                                                                 = destinationPath;
    task.destinationFolder                                                                      = destinationFolder;
    task.plannedItems = task.totalItems = 1u;
    task.totalBytes = task.itemTotalBytes = 19'864'223'744ull;
    task.completedBytes = task.itemCompletedBytes = 10'490'457'620ull;
    task.discoveredFileCount                      = 240u;
    task.discoveredDirectoryCount                 = 8u;
    task.copiedFileCount                          = 12u;
    task.copiedFolderCount                        = 2u;
    task.copiedFileCountKnown = task.copiedFolderCountKnown = task.autoConcurrencyUsed = true;
    task.autoTunedConcurrency = task.effectiveConcurrencyBudget = 4u;
    task.lastProgressCallbackTick                               = GetTickCount64();
    task.operationStartTick                                     = task.lastProgressCallbackTick - 200'000u;
    renderer->DebugSetVisualScenario({task}, false, true);
    PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), 900, 460);
    const auto subjectId = std::format(L"FileOperations.Task.{}.Subject", task.taskId);
    std::wstring name;
    std::wstring help;
    bool valid = renderer->DebugHostedRegionAccessibleTextForSelfTest(subjectId, name, help) &&
                 name == FormatStringResource(nullptr, IDS_FMT_FILEOPS_OPERATION_SUBJECT, LoadStringResource(nullptr, IDS_FILEOP_OPERATION_COPY), subject);
    for (const auto& [suffix, label] :
         std::array<std::pair<std::wstring_view, UINT>, 2u>{{{L"Source.0", IDS_FILEOPS_LABEL_FROM}, {L"Destination", IDS_FILEOPS_LABEL_TO}}})
    {
        const auto expected = std::format(L"{} {}", LoadStringResource(nullptr, label), suffix == L"Source.0" ? sourcePath : destinationFolder);
        valid = renderer->DebugHostedRegionAccessibleTextForSelfTest(std::format(L"FileOperations.Task.{}.Path.{}", task.taskId, suffix), name, help) &&
                name == expected && help == expected && valid;
    }
    PopupLayoutDebugSnapshot popupSnapshot{};
    popupSnapshot.taskId = task.taskId;
    valid                = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), popupSnapshot) && ! popupSnapshot.taskHasVisibleButtonOverlap &&
                           popupSnapshot.taskCountsText == L"Folders: 2/8 · Files: 12/240" && valid;
    if (! valid || FAILED(RedSalamander::TestSupport::SaveWindowScreenshot(popup.get(), output / L"file-operations-popup.png")))
        return false;

    auto counting                 = task;
    counting.discoveryClosed      = false;
    counting.discoveredBytesKnown = false;
    counting.totalBytes           = 0u;
    counting.discoveredFileCount = counting.discoveredDirectoryCount = 0u;
    renderer->DebugSetVisualScenario({counting}, false, true);
    PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), 900, 460);
    popupSnapshot        = {};
    popupSnapshot.taskId = counting.taskId;
    if (! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), popupSnapshot) || popupSnapshot.taskCountsText != L"Counting…")
        return false;
    counting.discoveredFileCount      = task.discoveredFileCount;
    counting.discoveredDirectoryCount = task.discoveredDirectoryCount;
    renderer->DebugSetVisualScenario({counting}, false, true);
    PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), 900, 460);
    popupSnapshot        = {};
    popupSnapshot.taskId = counting.taskId;
    valid                = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), popupSnapshot) && ! popupSnapshot.taskHasVisibleButtonOverlap &&
                           popupSnapshot.taskCountsText == L"Counting… Folders: 8 · Files: 240";
    if (! valid || FAILED(RedSalamander::TestSupport::SaveWindowScreenshot(popup.get(), output / L"file-operations-counting.png")))
        return false;

    // These private scenes exercise different count authorities and lifecycle
    // states. Literal English expectations catch copy regressions independently
    // of the resources used by the renderer; the broader gallery covers satellites.
    constexpr std::array<std::wstring_view, 6u> additionalNames{{L"file-operations-move-totals.png",
                                                                 L"file-operations-file-progress.png",
                                                                 L"file-operations-paused.png",
                                                                 L"file-operations-queued.png",
                                                                 L"file-operations-completed.png",
                                                                 L"file-operations-conflict.png"}};
    for (size_t scene = 0u; scene < additionalNames.size(); ++scene)
    {
        auto example                = task;
        std::wstring expectedCounts = L"Folders: 2/8 · Files: 12/240";
        UINT expectedStatus         = IDS_FILEOPS_STATUS_RUNNING;
        if (scene == 0u)
        {
            example.operation    = FILESYSTEM_MOVE;
            example.selection    = {std::wstring(subject), 2u, 0u, 2u, true};
            example.plannedItems = example.totalItems = 2u;
            example.discoveredDirectoryCount          = 2u;
            example.discoveredFileCount               = 31u;
            example.copiedFileCountKnown = example.copiedFolderCountKnown = false;
            expectedCounts                                                = L"Folders: 2 · Files: 31";
        }
        else if (scene == 1u)
        {
            example.copiedFolderCountKnown = false;
            expectedCounts                 = L"Folders: 8 · Files: 12/240";
        }
        else if (scene == 2u)
        {
            example.paused = true;
            expectedStatus = IDS_FILEOPS_GRAPH_PAUSED;
        }
        else if (scene == 3u)
        {
            example.started        = false;
            example.waitingInQueue = true;
            example.lifecyclePhase = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Waiting);
            example.completedBytes = example.itemCompletedBytes = 0u;
            example.copiedFileCount = example.copiedFolderCount = 0u;
            expectedCounts.clear(); // Waiting cards expose no work counters or activity.
            expectedStatus = IDS_FILEOPS_STATUS_WAITING;
        }
        else if (scene == 4u)
        {
            example.finished       = true;
            example.lifecyclePhase = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Terminal);
            example.completedBytes = example.itemCompletedBytes = example.totalBytes;
            example.completedItems                              = example.totalItems;
            example.copiedFileCount                             = example.discoveredFileCount;
            example.copiedFolderCount                           = example.discoveredDirectoryCount;
            expectedCounts.clear();
            expectedStatus = IDS_FILEOPS_STATUS_COMPLETED;
        }
        else
        {
            auto& prompt           = example.conflict;
            prompt.active          = true;
            prompt.generation      = 1u;
            prompt.bucket          = static_cast<uint8_t>(FileOperations::ConflictClass::RegularFileExists);
            prompt.status          = HRESULT_FROM_WIN32(ERROR_FILE_EXISTS);
            prompt.sourcePath      = sourcePath;
            prompt.destinationPath = destinationPath;
            prompt.sourceName = prompt.destinationName = L"session-01.mp4";
            prompt.sourceParentPath                    = LR"(C:\Demo\Media\Workshop)";
            prompt.destinationParentPath               = destinationFolder;
            prompt.sourceMetadata.available = prompt.sourceMetadata.sizeKnown = true;
            prompt.sourceMetadata.sizeBytes                                   = 512ull * 1024ull * 1024ull;
            prompt.sourceMetadata.lastWriteTime                               = 134'338'680'000'000'000ll;
            prompt.destinationMetadata                                        = prompt.sourceMetadata;
            prompt.destinationMetadata.sizeBytes                              = 480ull * 1024ull * 1024ull;
            prompt.destinationMetadata.lastWriteTime -= 864'000'000'000ll;
            expectedStatus = IDS_FILEOPS_STATUS_WAITING_FOR_DECISION;
        }
        renderer->DebugSetVisualScenario({example}, false, true);
        PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), 900, scene == 5u ? 660 : 460);
        popupSnapshot        = {};
        popupSnapshot.taskId = example.taskId;
        valid                = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), popupSnapshot) && ! popupSnapshot.taskHasVisibleButtonOverlap;
        if (scene == 5u)
        {
            // Conflict cards communicate status through the decision body, with
            // no duplicate chip. Assert the surface the product actually paints.
            valid = popupSnapshot.conflictWaitingForDecisionVisible && popupSnapshot.conflictPromptVisibleWithoutClipping &&
                    ! popupSnapshot.taskStatusChipActuallyPainted && valid;
        }
        else
            valid = popupSnapshot.taskCountsText == expectedCounts && popupSnapshot.taskStatusChipPaintedTextFits &&
                    popupSnapshot.taskStatusChipPaintedText == LoadStringResource(nullptr, expectedStatus) && valid;
        if (! valid)
            AppendLog(std::format(L"DocumentationPrivacy: scene={} counts='{}' status='{}' chipFits={} conflictFits={}",
                                  additionalNames[scene],
                                  popupSnapshot.taskCountsText,
                                  popupSnapshot.taskStatusChipPaintedText,
                                  popupSnapshot.taskStatusChipPaintedTextFits,
                                  popupSnapshot.conflictPromptVisibleWithoutClipping));
        if (! valid || FAILED(RedSalamander::TestSupport::SaveWindowScreenshot(popup.get(), output / additionalNames[scene])))
            return false;
    }

    auto failed                  = task;
    failed.operation             = FILESYSTEM_MOVE;
    failed.finished              = true;
    failed.lifecyclePhase        = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Terminal);
    failed.resultHr              = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    failed.errorCount            = 2u;
    failed.retainedSourceCount   = 1u;
    failed.completedBytes        = 0u;
    failed.itemCompletedBytes    = 0u;
    failed.copiedFileCount       = 0u;
    failed.copiedFolderCount     = 0u;
    failed.lastDiagnosticMessage = LoadStringResource(nullptr, IDS_FILEOPS_DIAG_DESTINATION_EXISTS_UNVERIFIABLE);
    renderer->DebugSetVisualScenario({failed}, false, true);
    PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), 760, 460);
    popupSnapshot        = {};
    popupSnapshot.taskId = failed.taskId;
    valid                = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), popupSnapshot) && ! popupSnapshot.taskHasVisibleButtonOverlap &&
                           popupSnapshot.taskLastNoteText == FormatStringResource(nullptr, IDS_FMT_FILEOPS_LAST_NOTE, failed.lastDiagnosticMessage) &&
                           popupSnapshot.taskLastNoteLineCount >= 2u && popupSnapshot.taskLastNoteFits && popupSnapshot.taskLastNotePrecedesProgress;
    if (! valid)
        AppendLog(std::format(L"DocumentationPrivacy: wrapped note invalid; lines={} fits={} precedesProgress={} overlap={}",
                              popupSnapshot.taskLastNoteLineCount,
                              popupSnapshot.taskLastNoteFits,
                              popupSnapshot.taskLastNotePrecedesProgress,
                              popupSnapshot.taskHasVisibleButtonOverlap));
    if (! valid || FAILED(RedSalamander::TestSupport::SaveWindowScreenshot(popup.get(), output / L"file-operations-wrapped-note.png")))
        return false;

    FolderWindow::FileOperationState::CompletedTaskSummary summary{};
    summary.taskId        = 0xD0C2u;
    summary.operation     = FILESYSTEM_COPY;
    summary.selection     = {std::wstring(issueName), 1u, 1u, 0u, true};
    summary.resultHr      = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    summary.errorCount    = 1u;
    summary.totalItems    = 1u;
    summary.completedTick = GetTickCount64();
    FolderWindow::FileOperationState::TaskDiagnosticEntry issue{};
    issue.localTime       = SYSTEMTIME{2026, 10, 1, 19, 10, 0, 0, 0};
    issue.taskId          = summary.taskId;
    issue.operation       = FILESYSTEM_COPY;
    issue.severity        = FolderWindow::FileOperationState::DiagnosticSeverity::Error;
    issue.status          = summary.resultHr;
    issue.category        = L"item.continueOnError";
    issue.message         = LoadStringResource(nullptr, IDS_FILEOPS_DIAG_SOURCE_MISSING);
    issue.sourcePath      = issueSource;
    issue.destinationPath = issueDestination;
    summary.issueDiagnostics.push_back(std::move(issue));
    operations.DebugAppendCompletedTaskForSelfTest(std::move(summary));
    issues.reset(FileOperationsIssuesPane::Create(&operations, folder, state.mainWindow, lifetime));
    if (! issues)
        return false;
    PlaceFileOpsGalleryWindow(issues.get(), monitors.front(), 900, 540);
    FileOperationsIssuesPane::SelfTestSnapshot issuesSnapshot{};
    valid = FileOperationsIssuesPane::SelfTestRefresh(issues.get(), true) && FileOperationsIssuesPane::SelfTestSelectTask(issues.get(), 0xD0C2u) &&
            FileOperationsIssuesPane::TryGetSelfTestSnapshot(issues.get(), issuesSnapshot, true) && issuesSnapshot.rowCount == 1u &&
            issuesSnapshot.selectionCount == 1u && issuesSnapshot.detailPresent &&
            issuesSnapshot.detailTexts[4] == std::format(L"{}\n{}", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_SOURCE), issueSource) &&
            issuesSnapshot.detailTexts[5] == std::format(L"{}\n{}", LoadStringResource(nullptr, IDS_FILEOPS_ISSUES_COL_DESTINATION), issueDestination) &&
            issuesSnapshot.detailTexts[8].empty() && ! issuesSnapshot.selectedCells.empty() && issuesSnapshot.selectedCells.front() == issueSource;
    for (const auto& text : issuesSnapshot.detailTexts)
        valid = text.find(L"[path hidden]") == std::wstring::npos && text.find(L"[name hidden]") == std::wstring::npos &&
                text.find(L"C:\\Users\\") == std::wstring::npos && text.find(L"\\\\") == std::wstring::npos && text.find(L'@') == std::wstring::npos && valid;
    for (const auto& text : issuesSnapshot.selectedCells)
        valid = text.find(L"[path hidden]") == std::wstring::npos && text.find(L"[name hidden]") == std::wstring::npos &&
                text.find(L"C:\\Users\\") == std::wstring::npos && text.find(L"\\\\") == std::wstring::npos && text.find(L'@') == std::wstring::npos && valid;
    if (! valid)
        AppendLog(std::format(L"DocumentationPrivacy: Issues projection failed; rows={} selected={} detail={} source={} destination={} subject={} file={}",
                              issuesSnapshot.rowCount,
                              issuesSnapshot.selectionCount,
                              issuesSnapshot.detailPresent,
                              issuesSnapshot.detailTexts[4],
                              issuesSnapshot.detailTexts[5],
                              issuesSnapshot.detailTexts[8],
                              issuesSnapshot.selectedCells.empty() ? std::wstring_view{} : std::wstring_view(issuesSnapshot.selectedCells.front())));
    if (! valid || FAILED(RedSalamander::TestSupport::SaveWindowScreenshot(issues.get(), output / L"file-operations-issues.png")))
        return false;
    std::ofstream manifest(output / L"scenarios.tsv", std::ios::binary);
    manifest << "file\tscenario\tlanguage\ttheme\tclientWidthDip\tdpi\tidentity\n"
             << "file-operations-popup.png\trunning-copy\ten-US\tdark\t900\t" << GetDpiForWindow(popup.get()) << "\tinvented Workshop media paths\n"
             << "file-operations-counting.png\topen-counting\ten-US\tdark\t900\t" << GetDpiForWindow(popup.get()) << "\tinvented Workshop media paths\n"
             << "file-operations-move-totals.png\ttotals-only-move\ten-US\tdark\t900\t" << GetDpiForWindow(popup.get()) << "\tinvented Workshop media paths\n"
             << "file-operations-file-progress.png\tfile-progress-only\ten-US\tdark\t900\t" << GetDpiForWindow(popup.get())
             << "\tinvented Workshop media paths\n"
             << "file-operations-paused.png\tpaused-copy\ten-US\tdark\t900\t" << GetDpiForWindow(popup.get()) << "\tinvented Workshop media paths\n"
             << "file-operations-queued.png\tqueued-copy\ten-US\tdark\t900\t" << GetDpiForWindow(popup.get()) << "\tinvented Workshop media paths\n"
             << "file-operations-completed.png\tcompleted-copy\ten-US\tdark\t900\t" << GetDpiForWindow(popup.get()) << "\tinvented Workshop media paths\n"
             << "file-operations-conflict.png\tdestination-conflict\ten-US\tdark\t900\t" << GetDpiForWindow(popup.get()) << "\tinvented Workshop media paths\n"
             << "file-operations-wrapped-note.png\twrapped-terminal-note\ten-US\tdark\t760\t" << GetDpiForWindow(popup.get())
             << "\tinvented Workshop media paths\n"
             << "file-operations-issues.png\tselected-source-missing-issue\ten-US\tdark\t900\t" << GetDpiForWindow(issues.get())
             << "\tinvented Workshop media paths\n";
    AppendLog(L"DocumentationPrivacy: ten private synthetic frames; exact invented names and paths; count authorities; lifecycle states; wrapped note; "
              L"conflict; one selected issue; no "
              L"redaction or personal data.");
    return manifest.good();
}

// Exercise the production painter and hosted accessibility regions: a correct aggregate
// snapshot alone cannot detect stray counting text or the source-row mini marquee.
[[nodiscard]] bool CaptureFileOpsWaitingCards(SelfTestState& state)
{
    using namespace FileOperationsPopupInternal;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread())
        return false;
    auto* folder        = TryGetFolderWindow(state.mainWindow);
    const auto output   = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::FileOperations, L"waiting-cards");
    const auto monitors = GetFileOpsGalleryMonitors();
    if (! folder || monitors.empty() || output.empty() || ! SelfTest::EnsureDirectory(output))
        return false;
    const HWND foreground    = GetForegroundWindow();
    const auto previousTheme = folder->GetTheme();
    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui.value().language.empty() && g_settings.ui.value().language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui.value().language};
    const auto restore = wil::scope_exit([&]
    {
        static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage));
        folder->ApplyTheme(previousTheme);
    });
    Common::Settings::Settings settings{};
    settings.fileOperations.emplace();
    auto owner = std::make_unique<FolderWindow>();
    owner->SetSettings(&settings);
    FolderWindow::FileOperationState operations(*owner);
    const auto lifetime = std::make_shared<int>(0);
    std::ofstream manifest(output / L"scenarios.tsv", std::ios::binary);
    manifest << "file\tlanguage\tclientWidthDip\tdpi\toperation\twaitKind\tfixture\n";
    bool passed = true;
    for (const auto language : {L"en-US", L"fr-FR"})
    {
        if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, language})))
            return false;
        for (const int width : {480, 960})
        {
            auto theme                  = ResolveAppTheme(width == 480 ? ThemeMode::Light : ThemeMode::Dark, L"WaitingCards");
            theme.reducedMotionOverride = true;
            folder->ApplyTheme(theme);
            for (const auto operation : {FILESYSTEM_COPY, FILESYSTEM_MOVE})
                for (unsigned wait = 0u; wait < 3u; ++wait)
                {
                    TaskSnapshot task{};
                    task.taskId               = 0xCA4Eu;
                    task.operation            = operation;
                    task.lifecyclePhase       = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Ready);
                    task.selection            = {L"Moonrise.2026.2160p.mkv", 1u, 1u, 0u, true};
                    task.currentSourcePath    = L"C:\\Studio\\Exports\\Moonrise.2026.2160p.mkv";
                    task.destinationFolder    = L"\\\\studio-nas\\Media\\Movies";
                    task.destinationPane      = FolderWindow::Pane::Right;
                    task.waitingInQueue       = wait == 0u;
                    task.waitingForOthers     = wait == 1u;
                    task.queuePaused          = wait == 2u;
                    task.started              = wait == 2u;
                    task.discoveryClosed      = wait == 1u;
                    task.discoveredBytesKnown = task.discoveryClosed;
                    task.totalBytes           = task.discoveryClosed ? 8ull << 30 : 0u;
                    // Stale stream/scan facts must never paint activity while the task waits.
                    task.discoveryAheadActive       = true;
                    task.operationStartTick         = task.started ? GetTickCount64() : 0u;
                    task.autoConcurrencyUsed        = true;
                    task.effectiveConcurrencyBudget = 4u;
                    task.inFlightFileCount          = wait == 2u ? 1u : 0u;
                    if (task.inFlightFileCount)
                        task.inFlightFiles[0].sourcePath = task.currentSourcePath;
                    wil::unique_hwnd popup(FileOperationsPopup::Create(&operations, folder, state.mainWindow, lifetime));
                    auto* renderer = popup ? reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA)) : nullptr;
                    if (! renderer)
                        return false;
                    auto running           = task;
                    running.started        = true;
                    running.waitingInQueue = running.waitingForOthers = running.queuePaused = false;
                    running.operationStartTick                                              = GetTickCount64();
                    renderer->DebugSetVisualScenario({running}, false, true);
                    PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), width, 400);
                    renderer->DebugUpdateLiveVisualScenario({task});
                    PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), width, 400);
                    for (unsigned frame = 0u; frame < 12u; ++frame)
                    {
                        Debug::Perf::Scope repaint(L"FileOps.Popup.WaitingCard.FrameUs");
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    }
                    PopupLayoutDebugSnapshot layout{};
                    layout.taskId = task.taskId;
                    bool correct = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), layout) && layout.found && layout.taskPaintedTransferText.empty() &&
                                   layout.taskPaintedDiscoveryText.empty() && layout.taskPaintedFileProgressCount == 0u &&
                                   ! layout.taskTransferProgressVisible && ! layout.taskDiscoveryIndicatorVisible &&
                                   ! layout.taskTransferProgressMarqueeVisible && ! layout.taskThroughputSummaryVisible && layout.taskConcurrencyText.empty() &&
                                   layout.taskPaintedPathCount == 2u && layout.taskCancelVisible && layout.taskSpeedLimitVisible &&
                                   (wait == 2u || layout.taskStartNowVisible) && ! layout.taskHasVisibleButtonOverlap &&
                                   layout.taskStatusChipPaintedText == LoadStringResource(nullptr, IDS_FILEOPS_STATUS_WAITING);
                    std::wstring text;
                    std::wstring help;
                    for (const auto suffix : {L"Progress.Amount", L"Progress.Percent", L"Progress.Counts", L"Throughput.Eta", L"Concurrency"})
                        correct = ! renderer->DebugHostedRegionAccessibleTextForSelfTest(
                                      std::format(L"FileOperations.Task.{}.{}", task.taskId, suffix), text, help) &&
                                  correct;
                    const auto file = std::format(L"waiting-{}-{}-{}-{}.png", language, width, static_cast<unsigned>(operation), wait);
                    if (wait == 0u)
                    {
                        const HRESULT captured        = RedSalamander::TestSupport::SaveWindowScreenshot(popup.get(), output / file);
                        correct                       = SUCCEEDED(captured) && correct;
                        const std::string manifestRow = std::format("{}\t{}\t{}\t{}\t{}\t{}\tsynthetic-studio-export-no-operation-admitted\n",
                                                                    Common::Strings::Utf8FromUtf16ReplacingInvalid(file),
                                                                    Common::Strings::Utf8FromUtf16ReplacingInvalid(language),
                                                                    width,
                                                                    GetDpiForWindow(popup.get()),
                                                                    static_cast<unsigned>(operation),
                                                                    wait);
                        manifest.write(manifestRow.data(), static_cast<std::streamsize>(manifestRow.size()));
                    }
                    AppendLog(std::format(L"WaitingCard: {} correct={} amount='{}' counts='{}' miniBars={} paths={} concurrency='{}'",
                                          file,
                                          correct,
                                          layout.taskPaintedTransferText,
                                          layout.taskPaintedDiscoveryText,
                                          layout.taskPaintedFileProgressCount,
                                          layout.taskPaintedPathCount,
                                          layout.taskConcurrencyText));
                    passed = correct && passed;
                }
        }
    }
    return passed && manifest.good() && GetForegroundWindow() == foreground;
}

[[nodiscard]] bool CaptureFileOpsCardHierarchy(SelfTestState& state)
{
    using namespace FileOperationsPopupInternal;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread())
        return false;
    auto* folder        = TryGetFolderWindow(state.mainWindow);
    auto* operations    = folder ? TryGetFileOps(folder) : nullptr;
    const auto output   = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::FileOperations, L"card-hierarchy");
    const auto monitors = GetFileOpsGalleryMonitors();
    if (! folder || ! operations || monitors.empty() || output.empty() || ! SelfTest::EnsureDirectory(output))
        return false;
    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui.value().language.empty() && g_settings.ui.value().language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui.value().language};
    const auto previousTheme      = folder->GetTheme();
    const bool previousFooter     = operations->GetPopupFooterOnly();
    const bool previousDensity    = operations->GetPopupCompactDensity();
    const bool previousThroughput = operations->GetPopupThroughputExpanded();
    const auto restore            = wil::scope_exit([&]
    {
        static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage));
        folder->ApplyTheme(previousTheme);
        operations->SetPopupFooterOnly(previousFooter);
        operations->SetPopupCompactDensity(previousDensity);
        operations->SetPopupThroughputExpanded(previousThroughput);
    });
    operations->SetPopupFooterOnly(false);
    operations->SetPopupCompactDensity(false);
    operations->SetPopupThroughputExpanded(true);
    const auto lifetime = std::make_shared<int>(0);
    std::ofstream manifest(output / L"scenarios.tsv", std::ios::binary);
    manifest << "file\tlanguage\ttheme\tclientWidthDip\tdpi\tscenario\n";
    bool passed = true;
    for (const auto language : {L"en-US", L"fr-FR"})
    {
        if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, language})))
            return false;
        TaskSnapshot identity{};
        identity.operation         = FILESYSTEM_MOVE;
        identity.selection         = {L"original-folder", 1u, 0u, 1u, true};
        identity.currentSourcePath = L"C:\\original-folder\\new-child.bin";
        const auto originalHeader =
            FormatStringResource(nullptr, IDS_FMT_FILEOPS_OPERATION_SUBJECT, LoadStringResource(nullptr, IDS_FILEOP_OPERATION_MOVE), L"original-folder");
        for (const auto status : {TaskSnapshot::StatusKind::Waiting,
                                  TaskSnapshot::StatusKind::Preparing,
                                  TaskSnapshot::StatusKind::Discovering,
                                  TaskSnapshot::StatusKind::Running,
                                  TaskSnapshot::StatusKind::Paused,
                                  TaskSnapshot::StatusKind::Verifying,
                                  TaskSnapshot::StatusKind::Conflict,
                                  TaskSnapshot::StatusKind::Stopping,
                                  TaskSnapshot::StatusKind::Done,
                                  TaskSnapshot::StatusKind::Failed,
                                  TaskSnapshot::StatusKind::Canceled})
        {
            identity.statusKind = status;
            identity.finished =
                status == TaskSnapshot::StatusKind::Done || status == TaskSnapshot::StatusKind::Failed || status == TaskSnapshot::StatusKind::Canceled;
            identity.totalItems += 1'000u;
            passed = DebugBuildFileOperationsTaskHeaderText(identity, GetTickCount64()) == originalHeader && passed;
        }
        for (const int width : {480, 960})
        {
            auto theme                  = ResolveAppTheme(width == 480 ? ThemeMode::Light : ThemeMode::Dark, L"CardHierarchy");
            theme.reducedMotionOverride = true;
            folder->ApplyTheme(theme);
            for (unsigned scene = 0u; scene < 31u; ++scene)
            {
                TaskSnapshot task{};
                task.taskId                 = 0xCA4Du;
                task.operation              = FILESYSTEM_MOVE;
                task.lifecyclePhase         = static_cast<uint8_t>(FileOperations::TaskLifecyclePhase::Running);
                task.started                = true;
                task.discoveryClosed        = true;
                task.hasProgressCallbacks   = true;
                task.currentSourcePath      = L"C:\\Movies\\practical.magic.2.2026.hdr.2160p.web.h265-poke.mkv";
                task.currentDestinationPath = L"D:\\Movies\\practical.magic.2.2026.hdr.2160p.web.h265-poke.mkv";
                task.destinationFolder      = L"D:\\Movies";
                task.plannedItems = task.totalItems = 1u;
                task.selection                      = {L"practical.magic.2.2026.hdr.2160p.web.h265-poke.mkv", 1u, 1u, 0u, true};
                task.totalBytes                     = 19'864'223'744ull; // 18.5 GiB, displayed by the existing compact byte formatter.
                task.completedBytes                 = 10'490'457'620ull; // 9.77 GiB.
                task.itemTotalBytes                 = task.totalBytes;
                task.itemCompletedBytes             = task.completedBytes;
                task.autoConcurrencyUsed            = true;
                task.autoTunedConcurrency = task.effectiveConcurrencyBudget = 4u;
                task.lastProgressCallbackTick                               = GetTickCount64();
                task.operationStartTick                                     = task.lastProgressCallbackTick - 200'000u;
                if (scene == 2u || scene == 5u || scene == 6u)
                {
                    task.currentSourcePath = L"C:\\Photographs\\Family archive\\2026\\child.raw";
                    task.totalItems        = 500u;
                    task.plannedItems      = 10u;
                    task.selection         = {{}, 10u, 8u, 2u, scene != 6u};
                }
                if (scene == 4u)
                {
                    task.selection         = {L"Family archive", 1u, 0u, 1u, true};
                    task.currentSourcePath = L"C:\\Photographs\\Family archive\\2026\\child.raw";
                    task.totalItems        = 500u;
                }
                if (scene == 3u || scene == 7u)
                {
                    // The original source disappeared before mutation; this is not a partly written file.
                    task.totalBytes = task.completedBytes = task.itemTotalBytes = task.itemCompletedBytes = 0u;
                    task.finished                                                                         = true;
                    task.resultHr                                                                         = HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
                    task.errorCount                                                                       = 1u;
                    task.missingSourceCount                                                               = 1u;
                    task.lastDiagnosticMessage = LoadStringResource(nullptr, IDS_FILEOPS_DIAG_SOURCE_MISSING);
                }
                wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, lifetime));
                auto* renderer = popup ? reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA)) : nullptr;
                if (! renderer)
                    return false;
                if (scene == 8u)
                    task.conflict.active = true;
                if (scene == 9u)
                {
                    task.operation            = FILESYSTEM_DELETE;
                    task.selection.singleName = L"Photographie retouchée — 日本語 — été.mkv";
                }
                if (scene == 10u || scene == 11u)
                {
                    task.operation                = FILESYSTEM_COPY;
                    task.selection                = {scene == 10u ? L"Unknown sizes" : L"Zero-byte files", 1u, 0u, 1u, true};
                    task.currentSourcePath        = std::format(L"C:\\Photos\\{}\\child.mkv", task.selection.singleName);
                    task.discoveredDirectoryCount = 1u;
                    task.totalBytes = task.itemTotalBytes = task.itemCompletedBytes = 0u;
                    task.discoveredBytesKnown                                       = scene == 11u;
                    if (scene == 11u)
                        task.completedBytes = 0u;
                    task.copiedFileCountKnown = true;
                    task.copiedFileCount      = scene == 10u ? 32u : 10u;
                    task.discoveredFileCount  = scene == 10u ? 128u : 20u;
                }
                if (scene == 12u)
                {
                    task.operation         = FILESYSTEM_DELETE;
                    task.selection         = {L"Cleanup folder", 1u, 0u, 1u, true};
                    task.currentSourcePath = L"C:\\Cleanup folder\\child.bin";
                    task.totalBytes = task.completedBytes = task.itemTotalBytes = task.itemCompletedBytes = 0u;
                    task.totalItems                                                                       = 11u;
                    task.completedItems                                                                   = 4u;
                }
                std::wstring expectedCounts;
                if ((scene >= 13u && scene < 24u) || scene == 30u)
                {
                    task.operation                = FILESYSTEM_COPY;
                    task.discoveredFileCount      = 240u;
                    task.discoveredDirectoryCount = 8u;
                    task.copiedFileCount          = 12u;
                    task.copiedFolderCount        = 2u;
                    task.copiedFileCountKnown = task.copiedFolderCountKnown = task.discoveredBytesKnown = true;
                    task.selection                                                                      = {L"Family archive", 1u, 0u, 1u, true};
                    expectedCounts = FormatStringResource(nullptr, IDS_FMT_FILEOPS_FOLDER_FILE_PROGRESS, 2u, 8u, 12u, 240u);
                    if (scene == 14u || scene == 15u)
                    {
                        task.discoveryClosed      = false;
                        task.discoveryAheadActive = true;
                        if (scene == 14u)
                        {
                            task.discoveredFileCount = task.discoveredDirectoryCount = 0u;
                            expectedCounts                                           = LoadStringResource(nullptr, IDS_FILEOPS_COUNTING_FILES_FOLDERS);
                        }
                        else
                            expectedCounts = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNTING_FOUND, 8u, 240u);
                    }
                    if (scene == 16u)
                    {
                        task.discoveryClosed  = false;
                        task.discoverySkipped = true;
                        expectedCounts.clear(); // The existing completed-work summary owns the skipped-counting counts.
                    }
                    if (scene == 17u)
                    {
                        task.operation            = FILESYSTEM_MOVE;
                        task.copiedFileCountKnown = task.copiedFolderCountKnown = false;
                        expectedCounts                                          = FormatStringResource(nullptr, IDS_FMT_FILEOPS_FOLDER_FILE_TOTALS, 8u, 240u);
                    }
                    if (scene == 18u)
                    {
                        task.autoTunedConcurrency       = 8u;
                        task.effectiveConcurrencyBudget = 2u; // Display the effective limit, never the requested/resolved value or active file count.
                    }
                    if (scene == 19u)
                    {
                        task.operation           = FILESYSTEM_MOVE;
                        task.discoveredFileCount = task.discoveredDirectoryCount = 0u;
                        expectedCounts                                           = LoadStringResource(nullptr, IDS_FILEOPS_FILE_FOLDER_COUNTS_UNAVAILABLE);
                    }
                    if (scene == 20u)
                        task.autoConcurrencyUsed = false;
                    if (scene == 30u)
                        task.effectiveConcurrencyBudget = 0u;
                    if (scene == 21u)
                    {
                        task.discoveredFileCount = task.discoveredDirectoryCount = 0u;
                        task.copiedFileCount = task.copiedFolderCount = task.totalBytes = task.completedBytes = 0u;
                        expectedCounts = FormatStringResource(nullptr, IDS_FMT_FILEOPS_FOLDER_FILE_PROGRESS, 0u, 0u, 0u, 0u);
                    }
                    if (scene == 22u)
                    {
                        task.copiedFolderCountKnown = false;
                        expectedCounts              = FormatStringResource(nullptr, IDS_FMT_FILEOPS_FOUND_FOLDERS_FILE_PROGRESS, 8u, 12u, 240u);
                    }
                    if (scene == 23u)
                    {
                        renderer->debugGraphLabelText[0] = FormatStringResource(nullptr,
                                                                                IDS_FMT_FILEOPS_RATE_WITH_STALL,
                                                                                FormatStringResource(nullptr, IDS_FMT_FILEOP_SPEED_BYTES, L"57.6 MB"),
                                                                                LoadStringResource(nullptr, IDS_FILEOPS_STALL_WAITING_DESTINATION));
                        renderer->debugGraphLabelText[1] = LoadStringResource(nullptr, IDS_FILEOPS_ETA_UNAVAILABLE);
                    }
                }
                std::wstring expectedNote;
                if (scene >= 24u && scene < 30u)
                {
                    task.finished              = true;
                    task.resultHr              = HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                    task.errorCount            = 2u;
                    task.retainedSourceCount   = 1u;
                    task.lastDiagnosticMessage = LoadStringResource(nullptr, IDS_FILEOPS_DIAG_DESTINATION_EXISTS_UNVERIFIABLE);
                    if (scene == 25u)
                        task.lastDiagnosticMessage += L"\n" + std::wstring(160u, L'W') + L"\n日本語 — été — final paragraph.";
                    if (scene == 29u)
                    {
                        for (unsigned line = 0u; line < 260u; ++line)
                            task.lastDiagnosticMessage += L"\n" + LoadStringResource(nullptr, IDS_FILEOPS_DIAG_IDENTITY_UNAVAILABLE);
                        task.lastDiagnosticMessage += L"\nEnd of long note.";
                    }
                    if (scene == 26u || scene == 27u)
                    {
                        task.resultHr               = S_OK;
                        task.errorCount             = 0u;
                        task.noteCount              = 1u;
                        task.lastInfoNoteMessage    = std::exchange(task.lastDiagnosticMessage, std::wstring{});
                        task.showInformationalNotes = scene == 26u;
                    }
                    if (scene == 28u)
                    {
                        task.showInformationalNotes = true;
                        task.lastInfoNoteMessage    = L"This informational note must not replace the diagnostic.";
                        task.resultSummary          = LoadStringResource(nullptr, IDS_FILEOPS_DIAG_IDENTITY_UNAVAILABLE);
                    }
                    const auto& note = task.lastDiagnosticMessage.empty() ? task.lastInfoNoteMessage : task.lastDiagnosticMessage;
                    if (scene != 27u)
                        expectedNote = FormatStringResource(nullptr, IDS_FMT_FILEOPS_LAST_NOTE, note);
                }
                const bool collapsed = scene == 1u || scene == 3u || scene == 5u || scene == 6u || scene == 8u || scene == 9u;
                renderer->DebugSetVisualScenario({task}, collapsed, true);
                PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), width, 600);
                for (unsigned frame = 0u; frame < 12u; ++frame)
                {
                    Debug::Perf::Scope batch(L"FileOps.Popup.CardHierarchy.FrameUs");
                    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                }
                PopupLayoutDebugSnapshot layout{};
                layout.taskId = task.taskId;
                passed = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), layout) && layout.found && ! layout.taskHasVisibleButtonOverlap && passed;
                std::wstring subject;
                std::wstring help;
                const auto subjectId      = std::format(L"FileOperations.Task.{}.Subject", task.taskId);
                const auto expectedHeader = DebugBuildFileOperationsTaskHeaderText(task, GetTickCount64());
                const bool subjectCorrect =
                    renderer->DebugHostedRegionAccessibleTextForSelfTest(subjectId, subject, help) && subject == expectedHeader &&
                    layout.taskStatusHeaderPaintedText == expectedHeader && ! layout.taskSubjectPaintedText.empty() &&
                    (task.selection.singleName.empty() || task.selection.folderCount > 0u || layout.taskSubjectPaintedText.ends_with(L".mkv"));
                const bool paired =
                    collapsed || task.finished || task.operation == FILESYSTEM_DELETE ||
                    (((layout.taskPercentRect.right <= layout.taskPercentRect.left) ||
                      (layout.taskAmountRect.right <= layout.taskPercentRect.left && layout.taskAmountRect.top == layout.taskPercentRect.top)) &&
                     (layout.taskCountsText.empty() || layout.taskAmountRect.bottom <= layout.taskCountsRect.top) &&
                     (layout.taskCountsText.empty() || layout.taskCountsRect.right <= layout.taskEtaRect.left ||
                      layout.taskCountsRect.bottom <= layout.taskEtaRect.top) &&
                     (layout.taskThroughputLabelRect.right <= layout.taskRateRect.left || layout.taskThroughputLabelRect.bottom <= layout.taskRateRect.top) &&
                     layout.taskRateRect.right == layout.graphPlotRect.right);
                const bool headingFits = layout.taskSubjectRect.right <= layout.taskStatusChipPaintedRect.left;
                // Scene expectations are independent of whether the renderer published a region.
                // Unknown/open/skipped/empty Copy totals have no whole-task percentage.
                const bool expectsPercent  = ! collapsed && ! task.finished && task.operation != FILESYSTEM_DELETE && scene != 10u && scene != 14u &&
                                             scene != 15u && scene != 16u && scene != 21u;
                const auto expectedPercent = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPACT_PERCENT, scene == 11u ? 50u : 53u);
                std::wstring percentName;
                std::wstring percentHelp;
                const bool percentAccessible = renderer->DebugHostedRegionAccessibleTextForSelfTest(
                    std::format(L"FileOperations.Task.{}.Progress.Percent", task.taskId), percentName, percentHelp);
                const bool percentPresent =
                    layout.taskPercentRect.right > layout.taskPercentRect.left && layout.taskPercentRect.bottom > layout.taskPercentRect.top;
                const bool percentCorrect     = expectsPercent
                                                    ? percentPresent && percentAccessible && percentName == expectedPercent && percentHelp == expectedPercent
                                                    : ! percentPresent && ! percentAccessible;
                const bool expectsConcurrency = ! collapsed && ! task.finished && (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE) &&
                                                task.autoConcurrencyUsed && task.effectiveConcurrencyBudget > 0u;
                const bool detailsFit =
                    ! expectsConcurrency ||
                    (layout.taskConcurrencyText == FormatStringResource(nullptr, IDS_FMT_FILEOPS_PARALLEL_TRANSFERS, task.effectiveConcurrencyBudget) &&
                     layout.taskConcurrencyRect.top > layout.taskPaintedActionRowRect.bottom &&
                     layout.taskConcurrencyRect.bottom <= layout.taskPaintedCardRect.bottom);
                const bool graphFit =
                    collapsed || task.finished || task.operation == FILESYSTEM_DELETE || (layout.graphLabelsFit && layout.graphLabelsOutsidePlot);
                const bool amountCorrect = scene < 10u || scene > 11u ||
                                           layout.taskPaintedTransferText ==
                                               FormatStringResource(nullptr, IDS_FMT_FILEOPS_COPIED_FILES, task.copiedFileCount, task.discoveredFileCount);
                const bool countsCorrect = scene < 13u || (scene >= 24u && scene < 30u) || layout.taskCountsText == expectedCounts;
                std::wstring concurrencyName;
                const bool concurrencyAccessible = renderer->DebugHostedRegionAccessibleTextForSelfTest(
                    std::format(L"FileOperations.Task.{}.Concurrency", task.taskId), concurrencyName, help);
                const bool noUnexpectedConcurrency =
                    expectsConcurrency ? concurrencyAccessible && concurrencyName == layout.taskConcurrencyText && help == layout.taskConcurrencyText
                                       : layout.taskConcurrencyText.empty() && ! concurrencyAccessible;
                std::wstring accessibleCounts;
                const bool accessibleCorrect = scene < 13u || expectedCounts.empty() ||
                                               (renderer->DebugHostedRegionAccessibleTextForSelfTest(
                                                    std::format(L"FileOperations.Task.{}.Progress.Counts", task.taskId), accessibleCounts, help) &&
                                                accessibleCounts == expectedCounts);
                std::wstring accessibleNote;
                const auto noteId         = std::format(L"FileOperations.Task.{}.LastNote", task.taskId);
                const bool noteAccessible = renderer->DebugHostedRegionAccessibleTextForSelfTest(noteId, accessibleNote, help);
                const bool noteCorrect = scene < 24u || scene >= 30u ||
                                         (expectedNote.empty() ? ! noteAccessible && layout.taskLastNoteText.empty()
                                                               : noteAccessible && accessibleNote == expectedNote && layout.taskLastNoteText == expectedNote &&
                                                                     layout.taskLastNoteFits && layout.taskLastNotePrecedesProgress &&
                                                                     (width != 480 || layout.taskLastNoteLineCount > 1u));
                passed = subjectCorrect && paired && headingFits && detailsFit && graphFit && amountCorrect && countsCorrect && noUnexpectedConcurrency &&
                         accessibleCorrect && noteCorrect && percentCorrect && passed;
                const auto file        = std::format(L"scene{}-{}-{}dip.png", scene, language, width);
                passed                 = SUCCEEDED(RedSalamander::TestSupport::SaveWindowScreenshot(popup.get(), output / file)) && passed;
                const auto manifestRow = std::format("{}\t{}\t{}\t{}\t{}\t{}\n",
                                                     Common::Strings::Utf8FromUtf16ReplacingInvalid(file),
                                                     Common::Strings::Utf8FromUtf16ReplacingInvalid(language),
                                                     width == 480 ? "light" : "dark",
                                                     width,
                                                     GetDpiForWindow(popup.get()),
                                                     scene);
                manifest << manifestRow;
                AppendLog(std::format(
                    L"CardHierarchy: {} found={} buttonOverlap={} subject={} paired={} heading={} details={} graph={} counts={} accessible={} amount={}",
                    file,
                    layout.found,
                    layout.taskHasVisibleButtonOverlap,
                    subjectCorrect,
                    paired,
                    headingFits,
                    detailsFit,
                    graphFit,
                    countsCorrect,
                    accessibleCorrect,
                    amountCorrect));
                AppendLog(std::format(L"CardFacts: {} percentExpected={} percentCorrect={} concurrencyExpected={} concurrencyCorrect={}",
                                      file,
                                      expectsPercent,
                                      percentCorrect,
                                      expectsConcurrency,
                                      detailsFit && noUnexpectedConcurrency));
                if (scene >= 24u && scene < 30u)
                    AppendLog(std::format(L"LastNote: {} correct={} lines={} fits={} beforeProgress={} heightPx={}",
                                          file,
                                          noteCorrect,
                                          layout.taskLastNoteLineCount,
                                          layout.taskLastNoteFits,
                                          layout.taskLastNotePrecedesProgress,
                                          layout.taskLastNoteRect.bottom - layout.taskLastNoteRect.top));
                if (scene == 24u)
                {
                    const auto repaint = [&]()
                    {
                        for (unsigned paint = 0u; paint < 2u; ++paint)
                            RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    };
                    bool lifecycle = DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, task.taskId});
                    repaint();
                    lifecycle = ! renderer->DebugHostedRegionAccessibleTextForSelfTest(noteId, accessibleNote, help) && lifecycle;
                    lifecycle = DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, task.taskId}) && lifecycle;
                    repaint();
                    PopupLayoutDebugSnapshot changed{};
                    changed.taskId = task.taskId;
                    lifecycle      = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), changed) && changed.taskLastNoteText == expectedNote &&
                                     changed.taskLastNoteFits && changed.taskLastNotePrecedesProgress && lifecycle;
                    PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), width == 480 ? 960 : 480, 600);
                    repaint();
                    lifecycle                  = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), changed) && changed.taskLastNoteText == expectedNote &&
                                                 changed.taskLastNoteFits && changed.taskLastNotePrecedesProgress &&
                                                 (width == 480 ? changed.taskLastNoteLineCount <= layout.taskLastNoteLineCount
                                                               : changed.taskLastNoteLineCount >= layout.taskLastNoteLineCount) &&
                                                 lifecycle;
                    task.lastDiagnosticMessage = LoadStringResource(nullptr, IDS_FILEOPS_DIAG_SOURCE_MISSING);
                    renderer->DebugSetVisualScenario({task}, false, true);
                    repaint();
                    lifecycle = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), changed) && changed.taskLastNoteFits &&
                                changed.taskLastNoteText == FormatStringResource(nullptr, IDS_FMT_FILEOPS_LAST_NOTE, task.lastDiagnosticMessage) && lifecycle;
                    task.lastDiagnosticMessage.clear();
                    renderer->DebugSetVisualScenario({task}, false, true);
                    repaint();
                    lifecycle = ! renderer->DebugHostedRegionAccessibleTextForSelfTest(noteId, accessibleNote, help) &&
                                DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), changed) && changed.taskLastNoteText.empty() && lifecycle;
                    passed    = lifecycle && passed;
                    AppendLog(std::format(L"LastNote lifecycle: language={} widthDip={} passed={}", language, width, lifecycle));
                }
                if (scene == 29u)
                {
                    SendMessageW(popup.get(), WM_VSCROLL, SB_BOTTOM, 0);
                    for (unsigned paint = 0u; paint < 2u; ++paint)
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    PopupLayoutDebugSnapshot bottom{};
                    bottom.taskId = task.taskId;
                    RECT client{};
                    GetClientRect(popup.get(), &client);
                    const bool scrollFits = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), bottom) && bottom.taskLastNoteFits &&
                                            bottom.taskLastNotePrecedesProgress && bottom.taskLastNoteText == expectedNote &&
                                            bottom.taskLastNoteLineCount > 260u && bottom.taskLastNoteRect.bottom - bottom.taskLastNoteRect.top > 4096.0f &&
                                            bottom.taskLastNoteRect.top < 0.0f && bottom.taskLastNoteRect.bottom > 0.0f &&
                                            bottom.taskPaintedActionRowRect.bottom <= static_cast<float>(client.bottom);
                    passed                = scrollFits &&
                                            SUCCEEDED(RedSalamander::TestSupport::SaveWindowScreenshot(
                                                popup.get(), output / std::format(L"last-note-bottom-{}-{}dip.png", language, width))) &&
                                            passed;
                    AppendLog(std::format(L"LastNote scroll: language={} widthDip={} passed={} lines={} noteBottomPx={} actionsBottomPx={} clientBottomPx={}",
                                          language,
                                          width,
                                          scrollFits,
                                          bottom.taskLastNoteLineCount,
                                          bottom.taskLastNoteRect.bottom,
                                          bottom.taskPaintedActionRowRect.bottom,
                                          client.bottom));
                }
                if (scene == 0u)
                {
                    std::wstring technical;
                    passed                   = ! renderer->DebugHostedRegionAccessibleTextForSelfTest(
                                                   std::format(L"FileOperations.Task.{}.TechnicalDetails", task.taskId), technical, help) &&
                                               passed;
                    const auto concurrencyId = std::format(L"FileOperations.Task.{}.Concurrency", task.taskId);
                    const auto countsId      = std::format(L"FileOperations.Task.{}.Progress.Counts", task.taskId);
                    passed                   = DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, task.taskId}) && passed;
                    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    passed = ! renderer->DebugHostedRegionAccessibleTextForSelfTest(concurrencyId, technical, help) &&
                             ! renderer->DebugHostedRegionAccessibleTextForSelfTest(countsId, technical, help) && passed;
                    passed = DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, task.taskId}) && passed;
                    for (unsigned paint = 0u; paint < 2u; ++paint)
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    passed = renderer->DebugHostedRegionAccessibleTextForSelfTest(concurrencyId, technical, help) &&
                             technical == FormatStringResource(nullptr, IDS_FMT_FILEOPS_PARALLEL_TRANSFERS, 4u) &&
                             renderer->DebugHostedRegionAccessibleTextForSelfTest(countsId, technical, help) && passed;
                    AppendLog(std::format(L"CardHierarchy inline facts: accumulatedPass={}", passed));
                    passed = DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleThroughput, task.taskId}) && passed;
                    for (unsigned paint = 0u; paint < 2u; ++paint)
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    PopupLayoutDebugSnapshot hiddenPlot{};
                    hiddenPlot.taskId = task.taskId;
                    passed            = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), hiddenPlot) && ! hiddenPlot.taskThroughputExpanded &&
                                        hiddenPlot.hostedGraphControlCount == 0u && hiddenPlot.taskRateRect.right > hiddenPlot.taskRateRect.left &&
                                        hiddenPlot.taskEtaRect.right > hiddenPlot.taskEtaRect.left && ! operations->GetPopupThroughputExpanded() && passed;
                    passed            = SUCCEEDED(RedSalamander::TestSupport::SaveWindowScreenshot(
                                            popup.get(), output / std::format(L"throughput-collapsed-{}-{}dip.png", language, width))) &&
                                        passed;
                    passed            = DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, task.taskId}) && passed;
                    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    passed = DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, task.taskId}) && passed;
                    for (unsigned paint = 0u; paint < 2u; ++paint)
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    passed = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), hiddenPlot) && ! hiddenPlot.taskThroughputExpanded && passed;
                    passed = DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleThroughput, task.taskId}) && passed;
                    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    passed = operations->GetPopupThroughputExpanded() && passed;
                    AppendLog(std::format(
                        L"CardHierarchy throughput disclosure: retainedRateEta={} accumulatedPass={}", ! hiddenPlot.taskThroughputExpanded, passed));
                }
            }
        }
    }
    return passed && manifest.good();
}

[[nodiscard]] bool CaptureFileOpsVisualGallery(SelfTestState& state)
{
    using namespace FileOperationsPopupInternal;
    const bool interaction = state.step == SelfTestState::Step::FileOps_VisualGalleryInteraction;
    if (! interaction && ! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread())
    {
        Fail(L"Visual gallery requires the existing no-activation selftest guard.");
        return false;
    }
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = folder ? TryGetFileOps(folder) : nullptr;
    if (! folder || ! operations)
        return false;
    const auto output = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::FileOperations, L"visual-gallery");
    if (output.empty() || ! SelfTest::EnsureDirectory(output))
    {
        Fail(L"Visual gallery could not create its authorized artifact directory.");
        return false;
    }
    Localization::LanguagePreference previousLanguage{};
    if (g_settings.ui.has_value() && ! g_settings.ui.value().language.empty() && g_settings.ui.value().language != L"system")
        previousLanguage = {Localization::LanguagePreferenceKind::Culture, g_settings.ui.value().language};
    if (FAILED(Localization::ApplyLanguagePreference({Localization::LanguagePreferenceKind::Culture, L"fr-FR"})))
        return false;
    auto restoreLanguage = wil::scope_exit([&] { static_cast<void>(Localization::ApplyLanguagePreference(previousLanguage)); });
    if (LoadStringResource(nullptr, IDS_FILEOP_BTN_CANCEL).find(L"Annuler") == std::wstring::npos)
    {
        Fail(L"French resource satellite was not applied.");
        return false;
    }
    const auto monitors = GetFileOpsGalleryMonitors();
    if (monitors.empty())
        return false;
    const auto previousTheme      = folder->GetTheme();
    const bool previousFooter     = operations->GetPopupFooterOnly();
    const bool previousDensity    = operations->GetPopupCompactDensity();
    auto restore                  = wil::scope_exit([&]
    {
        folder->ApplyTheme(previousTheme);
        operations->SetPopupFooterOnly(previousFooter);
        operations->SetPopupCompactDensity(previousDensity);
    });
    auto lifetime                 = std::make_shared<int>(0);
    const auto scenarios          = BuildFileOpsVisualScenarios();
    const ULONGLONG scenarioClock = GetTickCount64();
    std::ofstream manifest(output / L"scenarios.tsv", std::ios::binary);
    manifest << "file\tscenario\ttheme\tclientWidthDip\tdpi\tfixture\n";
    if (interaction)
        return CaptureFileOpsGalleryInteraction(state, output, manifest, monitors);
    std::ofstream displayManifest(output / L"displays.tsv", std::ios::binary);
    displayManifest << "display\tdevice\tleft\ttop\tright\tbottom\n";
    size_t scopeProviderWitnessCount                     = 0u;
    size_t scanStatusPresentWitnessCount                 = 0u;
    size_t scanStatusAbsentWitnessCount                  = 0u;
    bool pathLabelsFit                                   = true;
    bool deletionPresentationFits                        = true;
    bool footerSummariesFit                              = true;
    size_t pathDetailCaptureSets                         = 0u;
    size_t pathDetailDarkSets                            = 0u;
    size_t pathDetailLightSets                           = 0u;
    size_t pathDetailCaptureCount                        = 0u;
    bool abbreviatedPathsFit                             = true;
    bool waitingCardsFit                                 = true;
    bool waitingTransitionsFit                           = true;
    bool speedLimitLabelsFit                             = true;
    bool terminalPresentationFits                        = true;
    bool terminalFactsFit                                = true;
    size_t terminalFactsWitnessCount                     = 0u;
    bool terminalFactsTransitionsFit                     = true;
    size_t terminalFactsTransitionCount                  = 0u;
    constexpr size_t kTerminalFactsTransitionsPerMonitor = 30u;
    bool directIssuesFit                                 = true;
    size_t directIssuesWitnessCount                      = 0u;
    bool notStartedFit                                   = true;
    size_t notStartedWitnessCount                        = 0u;
    bool verificationHalfFit                             = true;
    size_t verificationHalfWitnessCount                  = 0u;
    bool directIssuesTransitionsFit                      = true;
    size_t directIssuesTransitionCount                   = 0u;
    bool statusChipLabelsFit                             = true;
    size_t statusChipWitnessCount                        = 0u;
    bool hostedProgressValuesMatch                       = true;
    size_t hostedProgressWitnessCount                    = 0u;
    bool completedSemanticsFit                           = true;
    size_t doneHeadingWitnessCount                       = 0u;
    size_t groupBadgeWitnessCount                        = 0u;
    for (size_t display = 0u; display < monitors.size(); ++display)
    {
        const auto& monitor = monitors[display];
        displayManifest << display + 1u << '\t' << Common::Strings::Utf8FromUtf16ReplacingInvalid(monitor.device) << '\t' << monitor.work.left << '\t'
                        << monitor.work.top << '\t' << monitor.work.right << '\t' << monitor.work.bottom << '\n';
        const std::array themeModes{ThemeMode::Dark, ThemeMode::Light, ThemeMode::Rainbow, ThemeMode::HighContrast};
        const std::array<std::wstring_view, 4> themeNames{L"dark", L"light", L"rainbow", L"high-contrast"};
        for (size_t themeIndex = 0u; themeIndex < themeModes.size(); ++themeIndex)
        {
            auto theme                  = ResolveAppTheme(themeModes[themeIndex], L"FileOperationsGallery");
            theme.reducedMotionOverride = true;
            folder->ApplyTheme(theme);
            for (size_t sceneIndex = 0u; sceneIndex < scenarios.size(); ++sceneIndex)
            {
                // All semantics in the dark baseline; representative risk/progress/layouts in other themes.
                if (themeIndex != 0u && ! (themeIndex == 1u && scenarios[sceneIndex].name == L"27-long-unicode-paths") && sceneIndex != 9u &&
                    sceneIndex != 30u && sceneIndex != 32u && sceneIndex != 54u && scenarios[sceneIndex].name != L"152-completed-not-verified" &&
                    scenarios[sceneIndex].name != L"153-compact-verification-shares" &&
                    ! (themeIndex == 3u &&
                       (scenarios[sceneIndex].name == L"144-terminal-move-mixed-facts" || scenarios[sceneIndex].name == L"145-direct-issues-row")))
                    continue;
                const auto& scenario             = scenarios[sceneIndex];
                const bool waitingLayoutScenario = ! scenario.footerOnly && ! scenario.collapsed && ! scenario.compactDensity &&
                                                   std::any_of(scenario.tasks.begin(),
                                                               scenario.tasks.end(),
                                                               [](const TaskSnapshot& task)
                {
                    return ! task.started && ! task.finished && ! task.conflict.active && (task.waitingInQueue || task.waitingForOthers) &&
                           (task.operation == FILESYSTEM_COPY || task.operation == FILESYSTEM_MOVE);
                });
                const bool statusChipScenario    = scenario.name == L"06-discovery-before-transfer" || scenario.name == L"134-speed-limit-discovery" ||
                                                   scenario.name == L"139-synthetic-preparing-elapsed-chip";
                for (const int widthDip : {480, 760})
                {
                    const bool terminalTextScenario = scenario.tasks.size() == 1u && scenario.tasks.front().finished &&
                                                      ! scenario.tasks.front().resultSummary.empty() && ! scenario.collapsed;
                    if (widthDip == 760 && ! statusChipScenario && scenario.name.find(L"completed-group") == std::wstring::npos && ! terminalTextScenario &&
                        scenario.name.find(L"speed-limit") == std::wstring::npos && ! waitingLayoutScenario && scenario.name != L"09-copy-single-stream" &&
                        scenario.name != L"14-paused" && scenario.name.find(L"queue-paused") == std::wstring::npos && sceneIndex != 9u && sceneIndex != 12u &&
                        sceneIndex != 16u && sceneIndex != 26u && sceneIndex != 29u && sceneIndex != 32u &&
                        scenario.name != L"108-decision-long-french-paths" && scenario.name != L"116-footer-wrapped-summary" &&
                        scenario.name != L"117-expanded-wrapped-summary" && scenario.name != L"118-progress-long-filename" &&
                        scenario.name != L"119-progress-unicode-filename")
                        continue;
                    operations->SetPopupFooterOnly(scenario.footerOnly);
                    operations->SetPopupCompactDensity(scenario.compactDensity);
                    wil::unique_hwnd popup(FileOperationsPopup::Create(operations, folder, state.mainWindow, lifetime));
                    if (! popup)
                    {
                        Fail(L"Visual gallery could not create the production popup.");
                        return false;
                    }
                    auto* popupState = reinterpret_cast<FileOperationsPopupState*>(GetWindowLongPtrW(popup.get(), GWLP_USERDATA));
                    if (! popupState)
                    {
                        Fail(L"Visual gallery popup has no renderer state.");
                        return false;
                    }
                    auto snapshots             = scenario.tasks;
                    const ULONGLONG clockShift = GetTickCount64() - scenarioClock;
                    for (auto& snapshot : snapshots)
                    {
                        if (snapshot.lastProgressCallbackTick)
                            snapshot.lastProgressCallbackTick += clockShift;
                        if (snapshot.operationStartTick)
                            snapshot.operationStartTick += clockShift;
                        for (auto& stream : snapshot.inFlightFiles)
                        {
                            if (stream.lastUpdateTick)
                                stream.lastUpdateTick += clockShift;
                        }
                    }
                    popupState->DebugSetVisualScenario(std::move(snapshots), scenario.collapsed, scenario.groupExpanded);
                    PlaceFileOpsGalleryWindow(popup.get(), monitor, widthDip, scenario.footerOnly ? 88 : 720);
                    const UINT dpi = GetDpiForWindow(popup.get());
                    // A transient frame failure now retries (Q6); show the surface a host that
                    // cannot attach selects, pinned so reopening does not re-attach it.
                    if (scenario.graphicsFailure && ! popupState->DebugEnterFailureSurface(popup.get()))
                    {
                        Fail(L"Visual gallery could not select the popup failure surface.");
                        return false;
                    }
                    ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
                    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    // A second paint reflects the stable auto-size and retained-control rectangles.
                    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    if (themeIndex == 0u && widthDip == 480 && scenario.tasks.size() == 1u &&
                        scenario.tasks.front().kind == FileOperationsPopupInternal::TaskSnapshot::Kind::Informational)
                    {
                        // The fixture deliberately retains scan counters in terminal cards.
                        // Rendering/applying without a preceding scan must expose no stale row.
                        const auto& task         = scenario.tasks.front();
                        const auto& info         = task.informational;
                        const bool expectPresent = scenario.name != L"63-change-attributes-applying" && scenario.name != L"64-file-list-rendering" &&
                                                   ! scenario.name.starts_with(L"87-") && ! scenario.name.starts_with(L"88-") &&
                                                   ! scenario.name.starts_with(L"89-") && ! scenario.name.starts_with(L"90-") &&
                                                   ! scenario.name.starts_with(L"91-") && ! scenario.name.starts_with(L"92-");
                        std::wstring path;
                        uint64_t folders = 0u;
                        uint64_t entries = 0u;
                        using InfoKind   = FolderWindow::InformationalTaskUpdate::Kind;
                        switch (info.kind)
                        {
                            case InfoKind::CompareDirectories:
                                path    = info.scanCurrentRelative.native();
                                folders = info.scanFolderCount;
                                entries = info.scanEntryCount;
                                break;
                            case InfoKind::ChangeCase:
                                path    = info.changeCaseCurrentPath.native();
                                folders = info.changeCaseScannedFolders;
                                entries = info.changeCaseScannedEntries;
                                break;
                            case InfoKind::ChangeAttributes:
                                path    = info.changeAttributesCurrentPath.native();
                                folders = info.changeAttributesScannedFolders;
                                entries = info.changeAttributesScannedEntries;
                                break;
                            case InfoKind::MakeFileList:
                                path    = info.makeFileListCurrentPath.native();
                                folders = info.makeFileListScannedFolders;
                                entries = info.makeFileListScannedEntries;
                                break;
                        }
                        const std::wstring expectedCopyPath = path;
                        if (path.empty())
                            path = L".";
                        const std::wstring expectedText = FormatStringResource(nullptr, IDS_FMT_COMPARE_SCAN_STATUS, path, folders, entries);
                        PopupLayoutDebugSnapshot painted{};
                        painted.taskId  = task.taskId;
                        const bool read = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), painted);
                        bool valid      = read && painted.taskScanStatusVisible == expectPresent;
                        if (expectPresent)
                            valid = valid && painted.taskScanStatusFits && painted.taskScanStatusValidUnicode && painted.taskScanStatusTooltipComplete &&
                                    (! painted.taskScanStatusNeedsAbbreviation || painted.taskScanStatusAbbreviated) &&
                                    painted.taskScanStatusText == expectedText && painted.taskScanStatusCopyPath == expectedCopyPath;
                        // Identical paints reuse measurement and the retained text layout;
                        // every scenario/monitor still verifies the current complete value.
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                        PopupLayoutDebugSnapshot reused{};
                        reused.taskId          = task.taskId;
                        const bool cacheStable = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), reused) &&
                                                 reused.scanStatusMeasurementCount == painted.scanStatusMeasurementCount &&
                                                 reused.scanStatusRetainedLayouts <= reused.scanStatusLayoutSlots &&
                                                 reused.scanStatusRetainedLayouts == (expectPresent ? 1u : 0u) &&
                                                 reused.scanStatusLayoutSlots == scenario.tasks.size() && reused.taskScanStatusVisible == expectPresent &&
                                                 (! expectPresent || (reused.scanStatusRetainedLayouts > 0u && reused.taskScanStatusText == expectedText));
                        valid                  = valid && cacheStable;
                        std::wstring witnessFailure;
                        const std::wstring id    = std::format(L"FileOperations.Task.{}.ScanStatus", task.taskId);
                        const bool providerValid = VerifyGalleryHostedProvider(popup.get(),
                                                                               id,
                                                                               expectPresent,
                                                                               [&](IRawElementProviderSimple* provider)
                        {
                            for (const PROPERTYID property : {UIA_NamePropertyId, UIA_HelpTextPropertyId})
                            {
                                wil::unique_variant value;
                                if (FAILED(provider->GetPropertyValue(property, &value)) || value.vt != VT_BSTR || ! value.bstrVal ||
                                    std::wstring_view(value.bstrVal, SysStringLen(value.bstrVal)) != expectedText)
                                {
                                    witnessFailure = L"Scan status must expose its complete localized text in UIA Name and HelpText.";
                                    return false;
                                }
                            }
                            wil::unique_variant focusable;
                            if (FAILED(provider->GetPropertyValue(UIA_IsKeyboardFocusablePropertyId, &focusable)) || focusable.vt != VT_BOOL ||
                                (focusable.boolVal == VARIANT_TRUE) != ! expectedCopyPath.empty())
                            {
                                witnessFailure = L"Only a nonempty scan path may expose a focusable copy target; a placeholder is read-only.";
                                return false;
                            }
                            return true;
                        },
                                                                               witnessFailure);
                        valid                    = valid && providerValid;
                        AppendLog(std::format(L"ScanStatusWitness: {} dpi={} present={} fit={} unicode={} abbreviated={} needsAbbreviation={} tooltip={} "
                                              L"provider={} cacheStable={} valid={} reason={}",
                                              scenario.name,
                                              dpi,
                                              expectPresent,
                                              painted.taskScanStatusFits,
                                              painted.taskScanStatusValidUnicode,
                                              painted.taskScanStatusAbbreviated,
                                              painted.taskScanStatusNeedsAbbreviation,
                                              painted.taskScanStatusTooltipComplete,
                                              providerValid,
                                              cacheStable,
                                              valid,
                                              witnessFailure));
                        if (! valid)
                        {
                            Fail(std::format(L"Scan status paint/accessibility witness failed for {} at {} DPI: {}", scenario.name, dpi, witnessFailure));
                            return false;
                        }
                        if (expectPresent)
                            ++scanStatusPresentWitnessCount;
                        else
                            ++scanStatusAbsentWitnessCount;
                    }
                    if (themeIndex == 0u && widthDip == 480 &&
                        (scenario.name == L"57-decision-metadata-loading" || scenario.name == L"33-decision-file-exists" ||
                         scenario.name == L"55-decision-permanent-delete" || scenario.name == L"60-decision-apply-to-all-checked"))
                    {
                        const bool scopeEligible = scenario.name != L"57-decision-metadata-loading" && scenario.name != L"55-decision-permanent-delete";
                        const bool checked       = scenario.name == L"60-decision-apply-to-all-checked";
                        std::wstring witnessFailure;
                        const bool valid = VerifyApplyToAllScopeProvider(popup.get(), scenario.tasks.front().taskId, scopeEligible, checked, witnessFailure);
                        AppendLog(std::format(L"ScopeProviderWitness: {} dpi={} scopeEligible={} checked={} valid={} reason={}",
                                              scenario.name,
                                              dpi,
                                              scopeEligible,
                                              checked,
                                              valid,
                                              witnessFailure));
                        if (! valid)
                        {
                            Fail(std::format(L"Apply-to-all hosted UIA witness failed for {} at {} DPI: {}", scenario.name, dpi, witnessFailure));
                            return false;
                        }
                        ++scopeProviderWitnessCount;
                    }
                    if (statusChipScenario)
                    {
                        ++statusChipWitnessCount;
                        PopupLayoutDebugSnapshot chip{};
                        chip.taskId                      = scenario.tasks.front().taskId;
                        const bool read                  = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), chip);
                        const bool preparing             = scenario.name == L"139-synthetic-preparing-elapsed-chip";
                        const std::wstring expectedShort = LoadStringResource(nullptr, preparing ? IDS_FILEOPS_PREPARING : IDS_FILEOPS_STATUS_COUNTING_SHORT);
                        // fr-FR is forced above; this checks that the *painted header* still retains elapsed detail.
                        const bool elapsedHeader = chip.taskStatusHeaderPaintedText.find(L"écoulé") != std::wstring::npos;
                        const bool fits          = read && chip.taskStatusChipActuallyPainted && chip.taskStatusChipPaintedTextFits &&
                                                   chip.taskStatusChipPaintedText == expectedShort && elapsedHeader;
                        statusChipLabelsFit      = statusChipLabelsFit && fits;
                        AppendLog(std::format(
                            L"StatusChipPaintFit: {} dpi={} width={} painted={} text='{}' measuredPx={:.3f} availablePx={:.3f} elapsedHeader={} fit={}",
                            scenario.name,
                            dpi,
                            widthDip,
                            chip.taskStatusChipActuallyPainted,
                            chip.taskStatusChipPaintedText,
                            chip.taskStatusChipPaintedTextWidthPx,
                            chip.taskStatusChipPaintedTextAvailableWidthPx,
                            elapsedHeader,
                            fits));
                    }
                    if (themeIndex == 0u && widthDip == 480 &&
                        (scenario.name == L"18-completed" || scenario.name == L"19-completed-verified" || scenario.name == L"31-completed-group-expanded" ||
                         scenario.name == L"136-completed-group-all-results-expanded"))
                    {
                        ++doneHeadingWitnessCount;
                        PopupLayoutDebugSnapshot painted{};
                        painted.taskId               = 1u;
                        const bool read              = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), painted);
                        const std::wstring operation = LoadStringResource(nullptr, IDS_FILEOP_OPERATION_COPY);
                        const std::wstring done      = LoadStringResource(nullptr, IDS_FILEOPS_STATUS_COMPLETED);
                        const bool fits              = read && painted.taskStatusChipActuallyPainted && painted.taskStatusChipPaintedTextFits &&
                                                       painted.taskStatusChipPaintedText == done && painted.taskStatusHeaderPaintedText == operation &&
                                                       ! painted.taskTerminalSummaryVisible;
                        completedSemanticsFit        = completedSemanticsFit && fits;
                        AppendLog(std::format(L"DoneHeaderPaint: {} dpi={} width={} chip='{}' header='{}' fit={}",
                                              scenario.name,
                                              dpi,
                                              widthDip,
                                              painted.taskStatusChipPaintedText,
                                              painted.taskStatusHeaderPaintedText,
                                              fits));
                    }
                    if (themeIndex == 0u && scenario.name == L"137-completed-group-all-results-collapsed")
                    {
                        ++groupBadgeWitnessCount;
                        PopupLayoutDebugSnapshot painted{};
                        const bool read                = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), painted);
                        const std::wstring partialName = FormatStringResource(nullptr, IDS_FMT_FILEOPS_BADGE_PARTIAL_TOOLTIP, 1ul);
                        const std::wstring unknownName = FormatStringResource(nullptr, IDS_FMT_FILEOPS_BADGE_INDETERMINATE_TOOLTIP, 1ul);
                        const bool allGlyphsDrawn      = std::ranges::all_of(painted.completedGroupPaintedGlyphs, [](wchar_t glyph) { return glyph != L'\0'; });
                        const bool fits       = read && painted.completedGroupExpectedBadges == 5u && painted.completedGroupPaintedBadges == 5u &&
                                                painted.completedGroupPaintedBadgeMask == 31u && painted.completedGroupBadgeCountsFit &&
                                                painted.completedGroupBadgeTooltipsInformative && allGlyphsDrawn &&
                                                painted.completedGroupPaintedGlyphs[1] != painted.completedGroupPaintedGlyphs[2] &&
                                                painted.completedGroupHostedNames[1] == partialName && painted.completedGroupHostedNames[2] == unknownName &&
                                                partialName != unknownName;
                        completedSemanticsFit = completedSemanticsFit && fits;
                        AppendLog(std::format(L"CompletedGroupBadgePaint: {} dpi={} width={} partialGlyph=U+{:04X} unknownGlyph=U+{:04X} partialName='{}' "
                                              L"unknownName='{}' fit={}",
                                              scenario.name,
                                              dpi,
                                              widthDip,
                                              static_cast<unsigned int>(painted.completedGroupPaintedGlyphs[1]),
                                              static_cast<unsigned int>(painted.completedGroupPaintedGlyphs[2]),
                                              painted.completedGroupHostedNames[1],
                                              painted.completedGroupHostedNames[2],
                                              fits));
                    }
                    if (themeIndex == 0u && widthDip == 480 && scenario.tasks.size() == 1u)
                    {
                        // Literal oracles: these exercise the actual hosted ProgressBar after SyncHostedControls,
                        // not a second computation of the task model's expected percentage.
                        double expectedValue             = 0.0;
                        bool expectedIndeterminate       = false;
                        bool expectedSegmented           = false;
                        unsigned int expectedVerifiedPct = 50u;
                        bool checkProgress               = true;
                        if (scenario.name == L"07-discovery-and-transfer")
                            expectedIndeterminate = true;
                        else if (scenario.name == L"09-copy-single-stream" || scenario.name == L"68-zero-byte-items")
                            expectedValue = 37.5;
                        else if (scenario.name == L"17-verifying")
                        {
                            expectedValue     = 75.0; // 100% transfer and 50% verification in equal halves.
                            expectedSegmented = true;
                        }
                        else if (scenario.name == L"152-completed-not-verified")
                        {
                            expectedValue       = 50.0; // A full transfer half and nothing verified (I35).
                            expectedSegmented   = true;
                            expectedVerifiedPct = 0u;
                        }
                        else if (scenario.name == L"13-delete")
                            expectedValue = 37.5; // 48 of 128 known Delete items.
                        else if (scenario.name == L"141-delete-closed-known-items-zero" || scenario.name == L"142-delete-closed-known-bytes-zero")
                            expectedValue = 0.0;
                        else if (scenario.name == L"140-closed-zero-byte-unknown-items" || scenario.name == L"143-delete-closed-unknown")
                            expectedIndeterminate = true;
                        else
                            checkProgress = false;
                        if (checkProgress)
                        {
                            ++hostedProgressWitnessCount;
                            PopupLayoutDebugSnapshot observed{};
                            observed.taskId = scenario.tasks.front().taskId;
                            const bool read = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), observed);
                            const std::wstring expectedHelp =
                                expectedSegmented ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_PROGRESS_SEGMENTS_ACCESSIBLE, 100u, expectedVerifiedPct)
                                                  : std::wstring{};
                            const bool match = read && observed.found && observed.taskHostedProgressPresent &&
                                               std::abs(observed.taskHostedProgressValue - expectedValue) <= 0.01 &&
                                               observed.taskHostedProgressIndeterminate == expectedIndeterminate &&
                                               observed.taskHostedProgressSegmented == expectedSegmented && observed.taskHostedProgressHelpText == expectedHelp;
                            hostedProgressValuesMatch = hostedProgressValuesMatch && match;
                            AppendLog(std::format(
                                L"HostedProgressValue: {} dpi={} width={} actual={:.2f} expected={:.2f} indeterminate={} segmented={} help='{}' fit={}",
                                scenario.name,
                                dpi,
                                widthDip,
                                observed.taskHostedProgressValue,
                                expectedValue,
                                observed.taskHostedProgressIndeterminate,
                                observed.taskHostedProgressSegmented,
                                observed.taskHostedProgressHelpText,
                                match));
                        }
                    }
                    for (const auto& task : scenario.tasks)
                    {
                        PopupLayoutDebugSnapshot measured{};
                        measured.taskId = task.taskId;
                        const bool read = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), measured);
                        if (! read || ! measured.taskSpeedLimitVisible)
                        {
                            if (scenario.name.find(L"speed-limit") != std::wstring::npos)
                                speedLimitLabelsFit = false;
                            continue;
                        }
                        const bool fits     = measured.taskSpeedLimitUsesSelectorChrome && measured.taskSpeedLimitShowsCurrentValue &&
                                              measured.taskSpeedLimitLabelFits && ! measured.taskHasVisibleButtonOverlap;
                        speedLimitLabelsFit = speedLimitLabelsFit && fits;
                        AppendLog(std::format(L"SpeedLimitLabelFit: {} task={} dpi={} width={} labelPx={:.3f} availablePx={:.3f} fit={}",
                                              scenario.name,
                                              task.taskId,
                                              dpi,
                                              widthDip,
                                              measured.taskSpeedLimitLabelWidthPx,
                                              measured.taskSpeedLimitTextWidthPx,
                                              fits));
                    }
                    if (scenario.name.find(L"completed-group") != std::wstring::npos)
                    {
                        PopupLayoutDebugSnapshot group{};
                        const bool fits = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), group) && group.completedGroupVisible &&
                                          group.completedGroupClearVisible && group.completedGroupClearLabelFits &&
                                          group.completedGroupExpanded == scenario.groupExpanded &&
                                          (group.completedGroupExpanded ||
                                           (group.completedGroupPaintedBadges == group.completedGroupExpectedBadges && group.completedGroupBadgeCountsFit)) &&
                                          (scenario.name != L"137-completed-group-all-results-collapsed" ||
                                           (group.completedGroupExpectedBadges == 5u && group.completedGroupPaintedBadges == 5u &&
                                            group.completedGroupPaintedBadgeMask == 31u));
                        terminalPresentationFits = terminalPresentationFits && fits;
                        AppendLog(std::format(L"CompletedGroupLabelFit: {} dpi={} width={} expectedBadges={} paintedBadges={} clearFits={} fit={}",
                                              scenario.name,
                                              dpi,
                                              widthDip,
                                              group.completedGroupExpectedBadges,
                                              group.completedGroupPaintedBadges,
                                              group.completedGroupClearLabelFits,
                                              fits));
                    }
                    if (scenario.tasks.size() == 1u && scenario.tasks.front().finished && ! scenario.tasks.front().resultSummary.empty() &&
                        ! scenario.collapsed && ! scenario.footerOnly && ! scenario.compactDensity)
                    {
                        PopupLayoutDebugSnapshot summary{};
                        summary.taskId           = scenario.tasks.front().taskId;
                        const bool fits          = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), summary) && summary.taskTerminalSummaryVisible &&
                                                   summary.taskTerminalSummaryFits && summary.taskTerminalSummaryPrecedesActions;
                        terminalPresentationFits = terminalPresentationFits && fits;
                        AppendLog(std::format(L"TerminalSummaryFit: {} dpi={} width={} visible={} fits={} precedesActions={} fit={}",
                                              scenario.name,
                                              dpi,
                                              widthDip,
                                              summary.taskTerminalSummaryVisible,
                                              summary.taskTerminalSummaryFits,
                                              summary.taskTerminalSummaryPrecedesActions,
                                              fits));
                    }
                    if (scenario.name == L"144-terminal-move-mixed-facts")
                    {
                        ++terminalFactsWitnessCount;
                        PopupLayoutDebugSnapshot facts{};
                        facts.taskId    = scenario.tasks.front().taskId;
                        const bool read = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), facts);
                        // I35: the expected text is built from the same resources in the active language, so a wording
                        // change cannot strand this oracle. A Move that left sources in place states "completed".
                        const std::array<std::wstring, 9> requiredClauses{{FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_PUBLISHED, 9u),
                                                                           FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_PUBLICATION_UNKNOWN, 1u),
                                                                           FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_SOURCE_REMOVED, 7u),
                                                                           FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_SOURCE_RETAINED, 2u),
                                                                           FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_SOURCE_UNKNOWN, 1u),
                                                                           FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_VERIFIED, 6u),
                                                                           FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_VERIFY_FAILED, 1u),
                                                                           FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_VERIFY_UNAVAILABLE, 1u),
                                                                           FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_VERIFY_CANCELED, 1u)}};
                        const bool allClausesOnce = std::ranges::all_of(requiredClauses,
                                                                        [&](const std::wstring& clause)
                        {
                            const size_t first = facts.taskTerminalFactsText.find(clause);
                            return first != std::wstring::npos && facts.taskTerminalFactsText.find(clause, first + clause.size()) == std::wstring::npos;
                        });
                        const auto axisPrefix     = [](UINT labelId)
                        {
                            const std::wstring label = LoadStringResource(nullptr, labelId);
                            std::wstring prefix      = FormatEmbeddedStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_FACT_AXIS, label, std::wstring{});
                            while (! prefix.empty() && prefix.back() == L' ')
                                prefix.pop_back();
                            return prefix;
                        };
                        const std::array<std::wstring, 3> axes{{axisPrefix(IDS_FILEOPS_TERMINAL_AXIS_RESULT),
                                                                axisPrefix(IDS_ITEM_PROPERTIES_ARTIFACT_FIELD_SOURCE),
                                                                axisPrefix(IDS_ITEM_PROPERTIES_ARTIFACT_FIELD_VERIFICATION)}};
                        const size_t publication  = facts.taskTerminalFactsText.find(axes[0]);
                        const size_t source       = facts.taskTerminalFactsText.find(axes[1]);
                        const size_t verification = facts.taskTerminalFactsText.find(axes[2]);
                        const bool threeAxes = publication != std::wstring::npos && source != std::wstring::npos && verification != std::wstring::npos &&
                                               publication < source && source < verification && facts.taskTerminalFactsText.find(L'\n', publication) < source &&
                                               facts.taskTerminalFactsText.find(L'\n', source) < verification;
                        bool clausesOnTheirAxes = threeAxes;
                        for (size_t index = 0u; index < requiredClauses.size(); ++index)
                        {
                            const size_t position = facts.taskTerminalFactsText.find(requiredClauses[index]);
                            const size_t lower    = index < 2u ? publication : (index < 5u ? source : verification);
                            const size_t upper    = index < 2u ? source : (index < 5u ? verification : facts.taskTerminalFactsText.size());
                            clausesOnTheirAxes    = clausesOnTheirAxes && position > lower && position < upper;
                        }
                        const bool fits  = read && facts.taskTerminalFactsVisible && facts.taskTerminalFactsFits && facts.taskTerminalFactsFollowsSummary &&
                                           facts.taskTerminalFactsPrecedesActions && facts.taskTerminalFactsText == facts.taskTerminalFactsHostedName &&
                                           allClausesOnce && clausesOnTheirAxes;
                        terminalFactsFit = terminalFactsFit && fits;
                        AppendLog(std::format(L"TerminalFactsFit: {} dpi={} width={} theme={} visible={} fits={} followsSummary={} precedesActions={} "
                                              L"clauses={} axes={} fit={} text='{}'",
                                              scenario.name,
                                              dpi,
                                              widthDip,
                                              themeNames[themeIndex],
                                              facts.taskTerminalFactsVisible,
                                              facts.taskTerminalFactsFits,
                                              facts.taskTerminalFactsFollowsSummary,
                                              facts.taskTerminalFactsPrecedesActions,
                                              allClausesOnce,
                                              clausesOnTheirAxes,
                                              fits,
                                              facts.taskTerminalFactsText));
                    }
                    if (themeIndex == 0u && widthDip == 480 &&
                        (scenario.name == L"152-completed-not-verified" || scenario.name == L"153-compact-verification-shares"))
                    {
                        // I35 (owner report 2026-10-13): a finished Verify On bar keeps its verification half visible (the
                        // divider and the outlined remainder), and a compact row reads 100% only when every byte was verified.
                        ++verificationHalfWitnessCount;
                        struct ExpectedCard final
                        {
                            uint64_t taskId = 0u;
                            std::wstring label;
                            double marksPercent = 0.0;
                            double value        = 0.0;
                        };
                        std::vector<ExpectedCard> expectedCards;
                        if (scenario.collapsed)
                        {
                            expectedCards.push_back({1u, FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPACT_PERCENT, 100u), 100.0, 100.0});
                            expectedCards.push_back({2u, FormatStringResource(nullptr, IDS_FMT_FILEOPS_COMPACT_VERIFIED_PERCENT, 25u), 25.0, 62.5});
                            expectedCards.push_back({3u, LoadStringResource(nullptr, IDS_FILEOPS_COMPACT_NOT_VERIFIED), 0.0, 50.0});
                        }
                        else
                        {
                            expectedCards.push_back({1u, std::wstring{}, 0.0, 50.0});
                        }
                        for (const ExpectedCard& expected : expectedCards)
                        {
                            PopupLayoutDebugSnapshot card{};
                            card.taskId         = expected.taskId;
                            const bool read     = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), card);
                            const bool fits     = read && card.found && card.taskHostedProgressPresent && card.taskVerificationMarksVisible &&
                                                  card.taskVerificationMarksOutlined &&
                                                  std::abs(card.taskVerificationMarksPercent - expected.marksPercent) <= 0.5 &&
                                                  std::abs(card.taskHostedProgressValue - expected.value) <= 0.5 &&
                                                  (! scenario.collapsed || card.taskCompactMeterLabel == expected.label);
                            verificationHalfFit = verificationHalfFit && fits;
                            AppendLog(
                                std::format(L"VerificationHalfFit: {} task={} dpi={} label='{}' marks={} outlined={} marksPercent={:.1f} value={:.1f} fit={}",
                                            scenario.name,
                                            expected.taskId,
                                            dpi,
                                            card.taskCompactMeterLabel,
                                            card.taskVerificationMarksVisible,
                                            card.taskVerificationMarksOutlined,
                                            card.taskVerificationMarksPercent,
                                            card.taskHostedProgressValue,
                                            fits));
                        }
                    }
                    if (scenario.name == L"149-delete-not-started")
                    {
                        // I35 R9.4 (P3): the result line is the card's one fact; untouched sources are not listed.
                        ++notStartedWitnessCount;
                        PopupLayoutDebugSnapshot card{};
                        card.taskId     = scenario.tasks.front().taskId;
                        const bool read = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), card);
                        const bool fits = read && card.taskTerminalSummaryVisible && card.taskTerminalSummaryFits && ! card.taskTerminalFactsVisible &&
                                          card.taskTerminalFactsText.empty();
                        notStartedFit   = notStartedFit && fits;
                        AppendLog(std::format(L"NotStartedFit: {} dpi={} width={} theme={} summary={} summaryFits={} facts={} fit={}",
                                              scenario.name,
                                              dpi,
                                              widthDip,
                                              themeNames[themeIndex],
                                              card.taskTerminalSummaryVisible,
                                              card.taskTerminalSummaryFits,
                                              card.taskTerminalFactsVisible,
                                              fits));
                    }
                    if (scenario.name == L"145-direct-issues-row" || scenario.name == L"146-diagnostic-counters-no-issues-row" ||
                        scenario.name == L"147-clean-no-issues")
                    {
                        ++directIssuesWitnessCount;
                        PopupLayoutDebugSnapshot actions{};
                        actions.taskId             = scenario.tasks.front().taskId;
                        const bool read            = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), actions);
                        const bool issuesExpected  = scenario.name == L"145-direct-issues-row";
                        const bool moreExpected    = scenario.name != L"147-clean-no-issues";
                        const size_t countExpected = 1u + static_cast<size_t>(issuesExpected) + static_cast<size_t>(moreExpected);
                        const std::wstring label   = LoadStringResource(nullptr, IDS_FILEOPS_BTN_ISSUES);
                        const bool issuesMatch =
                            actions.completedIssuesVisible == issuesExpected &&
                            (issuesExpected ? actions.completedIssuesHitTargetActive && actions.completedIssuesLabelFits &&
                                                  actions.completedIssuesHostedText == label && actions.completedIssuesHostedName == label &&
                                                  actions.completedIssuesButtonRect.right > actions.completedIssuesButtonRect.left &&
                                                  actions.completedIssuesButtonRect.bottom > actions.completedIssuesButtonRect.top
                                            : actions.completedIssuesHostedText.empty() && actions.completedIssuesHostedName.empty());
                        const bool fits = read && ! label.empty() && issuesMatch && actions.completedDismissVisible &&
                                          actions.completedDiagnosticsMoreVisible == moreExpected && actions.completedVisibleActionCount == countExpected &&
                                          actions.completedActionRowFits && ! actions.taskHasVisibleButtonOverlap &&
                                          ! actions.completedFailedItemsActionVisible;
                        directIssuesFit = directIssuesFit && fits;
                        AppendLog(std::format(L"DirectIssuesActionFit: {} dpi={} width={} theme={} issues={} more={} count={} rowFits={} fit={}",
                                              scenario.name,
                                              dpi,
                                              widthDip,
                                              themeNames[themeIndex],
                                              actions.completedIssuesVisible,
                                              actions.completedDiagnosticsMoreVisible,
                                              actions.completedVisibleActionCount,
                                              actions.completedActionRowFits,
                                              fits));
                    }
                    if (scenario.name == L"31-completed-group-expanded" || scenario.name == L"32-completed-group-collapsed" ||
                        scenario.name == L"69-all-items-skipped")
                    {
                        // Identical warmed fixture before/after layout changes. These are in-process
                        // resource samples and synchronous paints, not presented FPS or GPU bytes.
                        for (unsigned int frame = 0u; frame < 8u; ++frame)
                            RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                        for (unsigned int round = 0u; round < 5u; ++round)
                        {
                            const auto started = std::chrono::steady_clock::now();
                            for (unsigned int frame = 0u; frame < 8u; ++frame)
                                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                            const auto elapsed = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count();
                            const auto memory  = FolderWindowFileOperationsStateInternal::CaptureProcessMemorySnapshot();
                            if (memory.privateBytes == 0u || memory.workingSetBytes == 0u)
                                terminalPresentationFits = false;
                            AppendLog(std::format(
                                L"TerminalLayoutRedraw: {} dpi={} width={} round={} frames=8 elapsedUs={} privateBytes={} workingSetBytes={} gdi={} user={}",
                                scenario.name,
                                dpi,
                                widthDip,
                                round,
                                elapsed,
                                memory.privateBytes,
                                memory.workingSetBytes,
                                GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS),
                                GetGuiResources(GetCurrentProcess(), GR_USEROBJECTS)));
                        }
                    }
                    if (scenario.name == L"69-all-items-skipped")
                    {
                        // Keep the HWND and task identity while changing the result. The cache
                        // must neither reuse an old caption nor its shorter measured height.
                        const std::wstring original = popupState->debugVisualScenario.value().front().resultSummary;
                        const std::array<std::wstring, 3> values{
                            {L"Aucun élément n’a été copié.",
                             L"La copie est terminée, mais plusieurs éléments ont été ignorés à la suite de votre décision. Les fichiers déjà présents dans le "
                             L"dossier de destination ont été conservés. Consultez les détails des opérations pour examiner chaque conflit et le résultat "
                             L"effectivement confirmé.",
                             original}};
                        for (size_t index = 0u; index < values.size(); ++index)
                        {
                            popupState->debugVisualScenario.value().front().resultSummary = values[index];
                            for (unsigned int frame = 0u; frame < 2u; ++frame)
                                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                            PopupLayoutDebugSnapshot changed{};
                            changed.taskId           = scenario.tasks.front().taskId;
                            const bool fits          = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), changed) && changed.taskTerminalSummaryVisible &&
                                                       changed.taskTerminalSummaryFits && changed.taskTerminalSummaryPrecedesActions;
                            terminalPresentationFits = terminalPresentationFits && fits;
                            AppendLog(std::format(L"TerminalSummaryChange: dpi={} width={} stage={} fit={}", dpi, widthDip, index, fits));
                        }
                    }
                    if (scenario.name == L"11-speed-limited")
                    {
                        // Change the value in the same HWND without clearing renderer caches.
                        // Unlimited -> custom -> original must each size from the current caption.
                        const uint64_t originalLimit = popupState->debugVisualScenario.value().front().desiredSpeedLimitBytesPerSecond;
                        for (const uint64_t limit : {uint64_t{0}, uint64_t{999ull * 1024ull * 1024ull * 1024ull}, originalLimit})
                        {
                            popupState->debugVisualScenario.value().front().desiredSpeedLimitBytesPerSecond = limit;
                            for (unsigned int frame = 0u; frame < 2u; ++frame)
                                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                            PopupLayoutDebugSnapshot changed{};
                            changed.taskId      = scenario.tasks.front().taskId;
                            const bool fits     = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), changed) && changed.taskSpeedLimitVisible &&
                                                  changed.taskSpeedLimitShowsCurrentValue && changed.taskSpeedLimitLabelFits &&
                                                  ! changed.taskHasVisibleButtonOverlap;
                            speedLimitLabelsFit = speedLimitLabelsFit && fits;
                            AppendLog(std::format(L"SpeedLimitValueChange: dpi={} width={} limit={} labelPx={:.3f} availablePx={:.3f} fit={}",
                                                  dpi,
                                                  widthDip,
                                                  limit,
                                                  changed.taskSpeedLimitLabelWidthPx,
                                                  changed.taskSpeedLimitTextWidthPx,
                                                  fits));
                        }
                        for (unsigned int round = 0u; round < 5u; ++round)
                        {
                            const auto started = std::chrono::steady_clock::now();
                            for (unsigned int frame = 0u; frame < 8u; ++frame)
                                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                            const auto elapsed = std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count();
                            AppendLog(std::format(L"SpeedLimitRedraw: dpi={} width={} round={} frames=8 elapsedUs={}", dpi, widthDip, round, elapsed));
                        }
                    }
                    if (waitingLayoutScenario)
                    {
                        for (const auto& task : scenario.tasks)
                        {
                            if (task.started || task.finished || task.conflict.active || (! task.waitingInQueue && ! task.waitingForOthers))
                                continue;
                            PopupLayoutDebugSnapshot measured{};
                            measured.taskId    = task.taskId;
                            const bool read    = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), measured);
                            const float gapDip = (measured.taskPaintedActionRowRect.top - measured.taskPaintedBodyBottomPx) * 96.0f / static_cast<float>(dpi);
                            const bool fits    = read && measured.taskPaintedBodyBoundsValid && gapDip >= 4.0f && gapDip <= 16.0f &&
                                                 measured.taskPaintedActionRowRect.bottom <= measured.taskPaintedCardRect.bottom &&
                                                 measured.taskPaintedBodyBottomPx > measured.taskPaintedCardRect.top && measured.taskPaintedPathCount >= 2u &&
                                                 measured.taskPaintedPathsFit && ! measured.taskHasVisibleButtonOverlap && measured.taskCancelVisible &&
                                                 measured.taskCancelHitTargetActive && ! measured.taskTransferProgressVisible &&
                                                 ! measured.taskDiscoveryIndicatorVisible && ! measured.taskThroughputDisclosureVisible &&
                                                 ! measured.taskThroughputSummaryVisible && measured.graphPlotRect.bottom <= measured.graphPlotRect.top &&
                                                 (task.overlapQueued ? ! measured.taskStartNowVisible : measured.taskStartNowVisible) &&
                                                 measured.taskQueueMoveUpVisible == (! task.overlapQueued && task.canMoveQueueUp) &&
                                                 measured.taskQueueMoveDownVisible == (! task.overlapQueued && task.canMoveQueueDown);
                            waitingCardsFit    = waitingCardsFit && fits;
                            AppendLog(
                                std::format(L"WaitingCardFit: {} task={} dpi={} width={} heightDip={:.2f} gapDip={:.2f} paths={} startNow={} cancel={} fit={}",
                                            scenario.name,
                                            task.taskId,
                                            dpi,
                                            widthDip,
                                            (measured.taskPaintedCardRect.bottom - measured.taskPaintedCardRect.top) * 96.0f / static_cast<float>(dpi),
                                            gapDip,
                                            measured.taskPaintedPathCount,
                                            measured.taskStartNowVisible,
                                            measured.taskCancelVisible,
                                            fits));
                        }
                    }
                    if (scenario.name == L"116-footer-wrapped-summary" || scenario.name == L"117-expanded-wrapped-summary")
                    {
                        PopupLayoutDebugSnapshot footer{};
                        const bool fit     = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), footer) && footer.footerSummaryFits &&
                                             (widthDip != 480 || footer.footerSummaryLineCount > 1u) && footer.footerQueueModeSelectorHitTargetActive;
                        footerSummariesFit = footerSummariesFit && fit;
                        AppendLog(std::format(L"FooterSummaryFit: {} display={} dpi={} width={} lines={} requiredPx={:.3f} availablePx={:.3f} fit={}",
                                              scenario.name,
                                              display + 1u,
                                              dpi,
                                              widthDip,
                                              footer.footerSummaryLineCount,
                                              footer.footerSummaryRequiredHeightPx,
                                              footer.footerSummaryRect.bottom - footer.footerSummaryRect.top,
                                              fit));
                    }
                    if (scenario.name == L"118-progress-long-filename" || scenario.name == L"119-progress-unicode-filename" || scenario.name == L"17-verifying")
                    {
                        PopupLayoutDebugSnapshot paths{};
                        paths.taskId        = 1u;
                        const bool fit      = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), paths) && paths.taskPaintedPathCount >= 2u &&
                                              paths.taskPaintedPathsFit && paths.taskPaintedPathsValidUnicode && paths.taskPaintedPathsRetainRawExtension &&
                                              paths.completedGroupBadgeTooltipsInformative;
                        abbreviatedPathsFit = abbreviatedPathsFit && fit;
                        AppendLog(std::format(L"AbbreviatedPathsFit: {} display={} dpi={} width={} paths={} fit={} unicode={} extension={} accessible={}",
                                              scenario.name,
                                              display + 1u,
                                              dpi,
                                              widthDip,
                                              paths.taskPaintedPathCount,
                                              paths.taskPaintedPathsFit,
                                              paths.taskPaintedPathsValidUnicode,
                                              paths.taskPaintedPathsRetainRawExtension,
                                              paths.completedGroupBadgeTooltipsInformative));
                    }
                    if (scenario.name == L"17-verifying" || scenario.name == L"13-delete")
                    {
                        for (size_t warmup = 0u; warmup < 3u; ++warmup)
                            RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                        for (size_t batch = 0u; batch < 5u; ++batch)
                        {
                            const auto started = std::chrono::steady_clock::now();
                            for (size_t redraw = 0u; redraw < 8u; ++redraw)
                                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                            Debug::Perf::Emit(scenario.name == L"13-delete" ? L"FileOps.Popup.Delete.Redraw8Us" : L"FileOps.Popup.PathLabels.Redraw8Us",
                                              L"French progress; eight warm redraws",
                                              static_cast<uint64_t>(
                                                  std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count()),
                                              dpi,
                                              static_cast<uint64_t>(widthDip),
                                              S_OK);
                        }
                    }
                    if (scenario.name == L"17-verifying" || scenario.name == L"13-delete" || scenario.name == L"18-completed")
                    {
                        PopupLayoutDebugSnapshot labelSnapshot{};
                        labelSnapshot.taskId        = 1u;
                        const size_t expectedLabels = scenario.name == L"13-delete" ? 1u : 2u;
                        const bool fit              = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), labelSnapshot) &&
                                                      labelSnapshot.taskPaintedPathLabelCount == expectedLabels && labelSnapshot.taskPaintedPathLabelsFit;
                        pathLabelsFit               = pathLabelsFit && fit;
                        AppendLog(std::format(L"PathLabelFit: {} display={} dpi={} width={} labels={} slackPx={:.3f} fit={}",
                                              scenario.name,
                                              display + 1u,
                                              dpi,
                                              widthDip,
                                              labelSnapshot.taskPaintedPathLabelCount,
                                              labelSnapshot.taskPaintedPathLabelMinimumSlackPx,
                                              fit));
                        if (scenario.name == L"13-delete")
                        {
                            const bool noGraph = labelSnapshot.graphPlotRect.right <= labelSnapshot.graphPlotRect.left;
                            const bool deletionFits =
                                fit && noGraph && ! labelSnapshot.taskThroughputDisclosureVisible && labelSnapshot.taskInlineSpeedRowVisible;
                            deletionPresentationFits = deletionPresentationFits && deletionFits;
                            AppendLog(std::format(L"DeletePresentation: display={} dpi={} width={} graph={} disclosure={} accessibleRate={} fit={}",
                                                  display + 1u,
                                                  dpi,
                                                  widthDip,
                                                  ! noGraph,
                                                  labelSnapshot.taskThroughputDisclosureVisible,
                                                  labelSnapshot.taskInlineSpeedRowVisible,
                                                  deletionFits));
                        }
                    }
                    const std::wstring filename =
                        std::format(L"{}__{}__{}dip__fr-FR__display{}__{}dpi.png", scenario.name, themeNames[themeIndex], widthDip, display + 1u, dpi);
                    const HRESULT captured = RedSalamander::TestSupport::SaveWindowScreenshot(popup.get(), output / filename);
                    if (FAILED(captured))
                    {
                        Fail(std::format(L"Visual gallery capture {} failed: 0x{:08X}.", filename, static_cast<unsigned long>(captured)));
                        return false;
                    }
                    const auto narrow = [](std::wstring_view text) { return Common::Strings::Utf8FromUtf16ReplacingInvalid(text); };
                    manifest << narrow(filename) << '\t' << narrow(scenario.name) << '\t' << narrow(themeNames[themeIndex]) << '\t' << widthDip << '\t' << dpi
                             << "\tproduction-renderer-synthetic-presentation\n";
                    manifest.flush();
                    AppendLog(std::format(L"VisualGallery: {}", filename));
                    if (themeIndex <= 1u && widthDip == 480 && scenario.name == L"27-long-unicode-paths")
                    {
                        // Capture root and scrolled states on each actual display in dark and light.
                        // This test-only opener uses the production panel and never activates the popup.
                        const auto& pathTask            = scenario.tasks.front();
                        const std::wstring expectedPath = pathTask.currentSourcePath;
                        const std::wstring detailRowId  = std::format(L"FileOperations.Task.{}.Path.Source.0", pathTask.taskId);
                        const HWND nativeFocusBefore    = GetFocus();
                        const HWND foregroundBefore     = GetForegroundWindow();
                        ++pathDetailCaptureSets;
                        if (themeIndex == 0u)
                            ++pathDetailDarkSets;
                        else
                            ++pathDetailLightSets;

                        popupState->DebugOpenPathDetail(detailRowId);
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                        auto* const field = popupState->DebugPathDetailField();
                        DxUi::TextFieldDebugMultilineState top{};
                        if (expectedPath.size() <= 4096u || ! expectedPath.starts_with(L"\\\\") || expectedPath.find(L"école") == std::wstring::npos ||
                            expectedPath.find(L"שלום") == std::wstring::npos || expectedPath.find(L"\U0001F4F7") == std::wstring::npos || ! field ||
                            ! field->IsReadOnly() || ! field->IsMultiline() || field->GetText() != expectedPath ||
                            ! popupState->DebugPathDetailMultilineState(top) || ! top.canScrollVertically || top.firstVisibleLine != 0u ||
                            GetFocus() != nativeFocusBefore || GetForegroundWindow() != foregroundBefore)
                        {
                            Fail(L"Long UNC detail must expose the exact read-only Unicode path at its root without taking native focus.");
                            return false;
                        }

                        std::wstring providerFailure;
                        const bool readonlyTextProvider = VerifyGalleryHostedProvider(popup.get(),
                                                                                      L"FileOperations.PathDetail.Text",
                                                                                      true,
                                                                                      [&](IRawElementProviderSimple* simple)
                        {
                            wil::unique_variant controlType;
                            if (FAILED(simple->GetPropertyValue(UIA_ControlTypePropertyId, &controlType)) || controlType.vt != VT_I4 ||
                                controlType.lVal != UIA_EditControlTypeId)
                            {
                                providerFailure = L"Path detail provider must be the expected Edit control.";
                                return false;
                            }
                            wil::com_ptr_nothrow<IUnknown> patternUnknown;
                            wil::com_ptr_nothrow<ITextProvider> textPattern;
                            wil::com_ptr_nothrow<ITextRangeProvider> documentRange;
                            wil::unique_bstr documentText;
                            if (FAILED(simple->GetPatternProvider(UIA_TextPatternId, patternUnknown.put())) || ! patternUnknown ||
                                FAILED(patternUnknown.query_to(textPattern.put())) || ! textPattern ||
                                FAILED(textPattern->get_DocumentRange(documentRange.put())) || ! documentRange ||
                                FAILED(documentRange->GetText(-1, documentText.put())) || ! documentText ||
                                std::wstring_view(documentText.get(), SysStringLen(documentText.get())) != expectedPath)
                            {
                                providerFailure = L"TextPattern must expose the complete path.";
                                return false;
                            }
                            return true;
                        },
                                                                                      providerFailure);
                        if (! readonlyTextProvider)
                        {
                            Fail(providerFailure.empty() ? L"Read-only path detail UIA was unavailable." : providerFailure);
                            return false;
                        }

                        const auto saveDetail = [&](std::wstring_view suffix, std::wstring_view fixtureTag)
                        {
                            const std::wstring detailName = std::format(
                                L"27-long-unicode-paths-detail-{}__{}__480dip__fr-FR__display{}__{}dpi.png", suffix, themeNames[themeIndex], display + 1u, dpi);
                            if (GetFocus() != nativeFocusBefore || GetForegroundWindow() != foregroundBefore ||
                                FAILED(RedSalamander::TestSupport::SaveWindowScreenshot(popup.get(), output / detailName)))
                                return false;
                            manifest << narrow(detailName) << '\t' << narrow(scenario.name) << '\t' << narrow(themeNames[themeIndex]) << "\t480\t" << dpi
                                     << '\t' << narrow(fixtureTag) << '\n';
                            manifest.flush();
                            AppendLog(std::format(L"VisualGallery: {}", detailName));
                            ++pathDetailCaptureCount;
                            return true;
                        };
                        if (! saveDetail(L"root", L"production-renderer-synthetic-read-only-path-detail-root"))
                        {
                            Fail(L"Could not capture the non-activating long-path detail root view.");
                            return false;
                        }
                        // The non-activating gallery deliberately has no native
                        // editor focus. Drive its real control handler; protected
                        // Commands separately exercises routed PageDown input.
                        const bool pageDownHandled = popupState->DebugScrollPathDetailPageDownForSelfTest();
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                        DxUi::TextFieldDebugMultilineState scrolled{};
                        const bool scrolledStateRead = popupState->DebugPathDetailMultilineState(scrolled);
                        AppendLog(std::format(L"PathDetailPageDown: display={} dpi={} handled={} read={} first={}->{} visible={} total={} textSame={} "
                                              L"nativeFocus={}->{} foreground={}->{} logicalFocus='{}'",
                                              display + 1u,
                                              dpi,
                                              pageDownHandled,
                                              scrolledStateRead,
                                              top.firstVisibleLine,
                                              scrolled.firstVisibleLine,
                                              scrolled.visibleLineCount,
                                              scrolled.totalLineCount,
                                              field->GetText() == expectedPath,
                                              reinterpret_cast<UINT_PTR>(nativeFocusBefore),
                                              reinterpret_cast<UINT_PTR>(GetFocus()),
                                              reinterpret_cast<UINT_PTR>(foregroundBefore),
                                              reinterpret_cast<UINT_PTR>(GetForegroundWindow()),
                                              popupState->DebugPathDetailLogicalFocusForSelfTest()));
                        if (! pageDownHandled || ! scrolledStateRead || scrolled.firstVisibleLine <= top.firstVisibleLine ||
                            scrolled.firstVisibleLine + scrolled.visibleLineCount > scrolled.totalLineCount || field->GetText() != expectedPath ||
                            GetFocus() != nativeFocusBefore || GetForegroundWindow() != foregroundBefore)
                        {
                            Fail(L"The PageDown control handler must scroll the same read-only document without editing it or activating the window.");
                            return false;
                        }
                        if (! saveDetail(L"scrolled", L"production-renderer-synthetic-read-only-path-detail-scrolled-control-handler"))
                        {
                            Fail(L"Could not capture the non-activating long-path detail scrolled view.");
                            return false;
                        }
                        if (display == 0u && themeIndex == 0u && monitors.size() > 1u)
                        {
                            const HWND detailHwnd = popup.get();
                            const size_t caret    = field->GetCaretIndex();
                            const auto selection  = field->GetSelectionRange();
                            for (const size_t targetDisplay : std::array<size_t, 2>{{1u, 0u}})
                            {
                                PlaceFileOpsGalleryWindow(detailHwnd, monitors[targetDisplay], 480, 720);
                                RedrawWindow(detailHwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                                const auto* moved = popupState->DebugPathDetailField();
                                PopupLayoutDebugSnapshot movedState{};
                                movedState.taskId    = pathTask.taskId;
                                const bool preserved = moved && moved->GetText() == expectedPath && moved->GetCaretIndex() == caret &&
                                                       moved->GetSelectionRange() == selection &&
                                                       DebugGetFileOperationsPopupLayoutSnapshot(detailHwnd, movedState) &&
                                                       popupState->DebugPathDetailLogicalFocusForSelfTest() == L"FileOperations.PathDetail.Text" &&
                                                       GetFocus() == nativeFocusBefore && GetForegroundWindow() == foregroundBefore;
                                AppendLog(std::format(L"PathDetailDpiHop: display={} dpi={} sameHwnd={} textSelectionLogicalFocusPreserved={}",
                                                      targetDisplay + 1u,
                                                      GetDpiForWindow(detailHwnd),
                                                      detailHwnd == popup.get(),
                                                      preserved));
                                if (! preserved)
                                {
                                    Fail(L"Physical monitor moves must preserve path text, selection and logical detail focus without taking native focus.");
                                    return false;
                                }
                            }
                        }
                        SendMessageW(popup.get(), WM_KEYDOWN, VK_ESCAPE, 0);
                        SendMessageW(popup.get(), WM_KEYUP, VK_ESCAPE, 0);
                        RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW);
                        PopupLayoutDebugSnapshot restored{};
                        restored.taskId = pathTask.taskId;
                        if (popupState->DebugPathDetailField() || ! DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), restored) ||
                            restored.hostedFocusedAutomationId != detailRowId || GetFocus() != nativeFocusBefore || GetForegroundWindow() != foregroundBefore)
                        {
                            Fail(L"Closing path detail must restore its exact source row and preserve native focus/foreground.");
                            return false;
                        }
                    }
                    if (themeIndex == 0u && widthDip == 480 && scenario.name == L"144-terminal-move-mixed-facts")
                    {
                        // The application-original PNG above must retain the initial nine-item state.
                        // These probes reuse its HWND/task and mutate the retained scenario directly,
                        // preserving terminal layout caches between paints.
                        auto& current = popupState->debugVisualScenario.value().front();
                        PopupLayoutDebugSnapshot initial{};
                        initial.taskId                  = current.taskId;
                        const bool initialRead          = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), initial);
                        const std::wstring originalText = initial.taskTerminalFactsText;
                        std::wstring changedText        = originalText;
                        const std::wstring oldClause    = FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_PUBLISHED, 9u);
                        const std::wstring newClause    = FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_PUBLISHED, 18u);
                        const size_t oldPosition        = changedText.find(oldClause);
                        const bool originalKnown        = initialRead && initial.taskTerminalFactsVisible && oldPosition != std::wstring::npos;
                        if (oldPosition != std::wstring::npos)
                            changedText.replace(oldPosition, oldClause.size(), newClause.data(), newClause.size());
                        terminalFactsTransitionsFit = terminalFactsTransitionsFit && originalKnown;
                        const auto inspect = [&](std::wstring_view stage, bool visible, std::wstring_view expectedText, bool collapsed, bool footerOnly)
                            -> PopupLayoutDebugSnapshot
                        {
                            for (unsigned frame = 0u; frame < 2u; ++frame)
                                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                            PopupLayoutDebugSnapshot observed{};
                            observed.taskId             = current.taskId;
                            const bool read             = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), observed);
                            const bool content          = visible ? observed.taskTerminalFactsVisible && observed.taskTerminalFactsFits &&
                                                                        observed.taskTerminalFactsFollowsSummary && observed.taskTerminalFactsPrecedesActions &&
                                                                        observed.taskTerminalFactsText == expectedText &&
                                                                        observed.taskTerminalFactsHostedName == expectedText
                                                                  : ! observed.taskTerminalFactsVisible && observed.taskTerminalFactsText.empty() &&
                                                                        observed.taskTerminalFactsHostedName.empty();
                            const bool fits             = read && observed.taskCollapsed == collapsed && observed.footerOnly == footerOnly && content;
                            terminalFactsTransitionsFit = terminalFactsTransitionsFit && fits;
                            ++terminalFactsTransitionCount;
                            AppendLog(std::format(L"TerminalFactsTransition: dpi={} width={} stage={} visible={} collapsed={} footerOnly={} fit={}",
                                                  dpi,
                                                  widthDip,
                                                  stage,
                                                  observed.taskTerminalFactsVisible,
                                                  observed.taskCollapsed,
                                                  observed.footerOnly,
                                                  fits));
                            if (! fits || stage == L"footer-restored")
                            {
                                RECT client{};
                                GetClientRect(popup.get(), &client);
                                const auto& facts = observed.taskTerminalFactsPaintedRect;
                                AppendLog(
                                    std::format(L"TerminalFactsTransitionDetail: stage={} read={} nativeDpi={} reducedMotion={} resizeMotion={} client={}x{} "
                                                L"body={} summary={} textUnits={} hostedUnits={} factsRect=({},{},{},{}) fits={} follows={} precedes={}",
                                                stage,
                                                read,
                                                GetDpiForWindow(popup.get()),
                                                observed.reducedMotionEnabled,
                                                observed.autoResizeMotionActive,
                                                client.right,
                                                client.bottom,
                                                observed.taskPaintedBodyBoundsValid,
                                                observed.taskTerminalSummaryVisible,
                                                observed.taskTerminalFactsText.size(),
                                                observed.taskTerminalFactsHostedName.size(),
                                                facts.left,
                                                facts.top,
                                                facts.right,
                                                facts.bottom,
                                                observed.taskTerminalFactsFits,
                                                observed.taskTerminalFactsFollowsSummary,
                                                observed.taskTerminalFactsPrecedesActions));
                                const auto& viewport = observed.listViewportRect;
                                const auto& card     = observed.taskPaintedCardRect;
                                const auto& actions  = observed.taskPaintedActionRowRect;
                                AppendLog(std::format(L"TerminalFactsVisibility: stage={} groupVisible={} groupExpanded={} hiddenByGroup={} scroll={} "
                                                      L"viewport=({},{},{},{}) paintedCard=({},{},{},{}) actions=({},{},{},{})",
                                                      stage,
                                                      observed.completedGroupVisible,
                                                      observed.completedGroupExpanded,
                                                      observed.taskHiddenByCompletedGroup,
                                                      observed.scrollPosition,
                                                      viewport.left,
                                                      viewport.top,
                                                      viewport.right,
                                                      viewport.bottom,
                                                      card.left,
                                                      card.top,
                                                      card.right,
                                                      card.bottom,
                                                      actions.left,
                                                      actions.top,
                                                      actions.right,
                                                      actions.bottom));
                            }
                            return observed;
                        };
                        const auto perfPath   = SelfTest::GetPerfArtifactPath(L"perf_metrics.jsonl");
                        const auto perfOffset = [&]() -> std::optional<std::uintmax_t>
                        {
                            std::error_code error;
                            const auto bytes = std::filesystem::file_size(perfPath, error);
                            return ! error && Debug::Perf::HasJsonlOutput() ? std::optional<std::uintmax_t>{bytes} : std::nullopt;
                        };
                        const auto newMeasurements = [&](std::optional<std::uintmax_t> offset) -> std::optional<size_t>
                        {
                            if (! offset)
                                return std::nullopt;
                            std::ifstream perf(perfPath, std::ios::binary);
                            if (! perf || ! perf.seekg(static_cast<std::streamoff>(offset.value())))
                                return std::nullopt;
                            size_t count = 0u;
                            for (std::string line; std::getline(perf, line);)
                                count += line.find("\"metric\":\"FileOps.Popup.TerminalFacts.MeasureUs\"") != std::string::npos;
                            return perf.bad() ? std::nullopt : std::optional<size_t>{count};
                        };
                        const auto beforeChange = perfOffset();
                        // Synthetic count-only mutation probes the cache key; this is not an engine outcome fixture.
                        current.publishedItemCount   = 18u;
                        const auto changed           = inspect(L"count-18", true, changedText, false, false);
                        const auto changedMeasures   = newMeasurements(beforeChange);
                        const auto beforeUnchanged   = perfOffset();
                        const auto unchanged         = inspect(L"unchanged-repaint", true, changedText, false, false);
                        const auto unchangedMeasures = newMeasurements(beforeUnchanged);
                        const auto& a                = changed.taskTerminalFactsPaintedRect;
                        const auto& b                = unchanged.taskTerminalFactsPaintedRect;
                        const bool stable            = a.left == b.left && a.top == b.top && a.right == b.right && a.bottom == b.bottom &&
                                                       changedMeasures == std::optional<size_t>{1u} && unchangedMeasures == std::optional<size_t>{0u};
                        terminalFactsTransitionsFit  = terminalFactsTransitionsFit && stable;
                        AppendLog(std::format(L"TerminalFactsStableRepaint: dpi={} width={} changedMeasures={} unchangedMeasures={} fit={}",
                                              dpi,
                                              widthDip,
                                              changedMeasures.value_or(SIZE_MAX),
                                              unchangedMeasures.value_or(SIZE_MAX),
                                              stable));
                        static_cast<void>(DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, current.taskId, 0u}));
                        static_cast<void>(inspect(L"collapsed", false, {}, true, false));
                        static_cast<void>(DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, current.taskId, 0u}));
                        static_cast<void>(inspect(L"expanded", true, changedText, false, false));
                        operations->SetPopupFooterOnly(true);
                        static_cast<void>(inspect(L"footer-hidden", false, {}, false, true));
                        operations->SetPopupFooterOnly(false);
                        static_cast<void>(inspect(L"footer-restored", true, changedText, false, false));
                        current.publishedItemCount = 9u;
                        static_cast<void>(inspect(L"count-restored", true, originalText, false, false));

                        // A format-locale setting notification must invalidate the retained
                        // count layout even when the selected language and count stay fixed.
                        // Use a grouped value without changing the machine's regional settings.
                        constexpr unsigned long groupedCount = 1009u;
                        const std::wstring groupedClause     = FormatStringResource(nullptr, IDS_FMT_FILEOPS_TERMINAL_PUBLISHED, groupedCount);
                        std::wstring groupedText             = originalText;
                        const size_t groupedPosition         = groupedText.find(oldClause);
                        const bool groupedOracle             = groupedPosition != std::wstring::npos && ! groupedClause.empty();
                        if (groupedOracle)
                            groupedText.replace(groupedPosition, oldClause.size(), groupedClause);
                        terminalFactsTransitionsFit   = terminalFactsTransitionsFit && groupedOracle;
                        current.publishedItemCount    = groupedCount;
                        const auto beforeSettingsText = inspect(L"count-1009", true, groupedText, false, false);
                        const auto beforeSettings     = perfOffset();
                        static_cast<void>(SendMessageW(popup.get(), WM_SETTINGCHANGE, 0, 0));
                        const auto afterSettings    = inspect(L"settings-refreshed", true, groupedText, false, false);
                        const auto settingsMeasures = newMeasurements(beforeSettings);
                        const auto beforeRepeat     = perfOffset();
                        const auto repeatedSettings = inspect(L"settings-unchanged", true, groupedText, false, false);
                        const auto repeatMeasures   = newMeasurements(beforeRepeat);
                        const auto& beforeRc        = beforeSettingsText.taskTerminalFactsPaintedRect;
                        const auto& afterRc         = afterSettings.taskTerminalFactsPaintedRect;
                        const auto& repeatedRc      = repeatedSettings.taskTerminalFactsPaintedRect;
                        const bool settingsStable   = beforeRc.left == afterRc.left && beforeRc.top == afterRc.top && beforeRc.right == afterRc.right &&
                                                      beforeRc.bottom == afterRc.bottom && afterRc.left == repeatedRc.left && afterRc.top == repeatedRc.top &&
                                                      afterRc.right == repeatedRc.right && afterRc.bottom == repeatedRc.bottom &&
                                                      afterSettings.taskTerminalFactsHostedName == groupedText &&
                                                      repeatedSettings.taskTerminalFactsHostedName == groupedText &&
                                                      settingsMeasures == std::optional<size_t>{1u} && repeatMeasures == std::optional<size_t>{0u};
                        terminalFactsTransitionsFit = terminalFactsTransitionsFit && settingsStable;
                        AppendLog(std::format(L"TerminalFactsSettingsRefresh: dpi={} width={} rebuilds={} repeats={} stable={}",
                                              dpi,
                                              widthDip,
                                              settingsMeasures.value_or(SIZE_MAX),
                                              repeatMeasures.value_or(SIZE_MAX),
                                              settingsStable));

                        current.resultSummary.clear();
                        const auto factsOnly        = inspect(L"summary-empty", true, groupedText, false, false);
                        terminalFactsTransitionsFit = terminalFactsTransitionsFit && ! factsOnly.taskTerminalSummaryVisible;
                        current.resultSummary       = scenario.tasks.front().resultSummary;
                        current.publishedItemCount  = 9u;
                        static_cast<void>(inspect(L"original-restored", true, originalText, false, false));
                        // Repeat the live setting transition on the same retained HWND: a
                        // one-off successful repaint must not hide intermittent retirement.
                        for (unsigned cycle = 0u; cycle < 8u; ++cycle)
                        {
                            operations->SetPopupFooterOnly(true);
                            static_cast<void>(inspect(std::format(L"footer-cycle-{}-hidden", cycle), false, {}, false, true));
                            operations->SetPopupFooterOnly(false);
                            static_cast<void>(inspect(std::format(L"footer-cycle-{}-restored", cycle), true, originalText, false, false));
                        }
                        // Also exercise the actual disclosure handler, which retains and
                        // restores expanded placement in addition to changing the setting.
                        terminalFactsTransitionsFit =
                            DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::FooterToggleDetails, 0u, 0u}) && terminalFactsTransitionsFit;
                        static_cast<void>(inspect(L"footer-action-hidden", false, {}, false, true));
                        terminalFactsTransitionsFit =
                            DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::FooterToggleDetails, 0u, 0u}) && terminalFactsTransitionsFit;
                        static_cast<void>(inspect(L"footer-action-restored", true, originalText, false, false));
                    }
                    if (themeIndex == 0u && widthDip == 480 && scenario.name == L"145-direct-issues-row")
                    {
                        // Mutate one retained presentation snapshot on one HWND. The
                        // counter-only stage must retain recovery More without an
                        // Issues button; a zero-counter row must still show Issues.
                        auto& current                  = popupState->debugVisualScenario.value().front();
                        const std::wstring issuesLabel = LoadStringResource(nullptr, IDS_FILEOPS_BTN_ISSUES);
                        const auto inspectActions      = [&](std::wstring_view stage, bool issues, bool more, size_t count, bool collapsed, bool footerOnly)
                        {
                            for (unsigned frame = 0u; frame < 2u; ++frame)
                                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                            PopupLayoutDebugSnapshot observed{};
                            observed.taskId = current.taskId;
                            const bool read = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), observed);
                            const bool issueContent =
                                observed.completedIssuesVisible == issues &&
                                (issues ? observed.completedIssuesHitTargetActive && observed.completedIssuesLabelFits &&
                                              observed.completedIssuesHostedText == issuesLabel && observed.completedIssuesHostedName == issuesLabel
                                        : observed.completedIssuesHostedText.empty() && observed.completedIssuesHostedName.empty());
                            const bool fits = read && observed.taskCollapsed == collapsed && observed.footerOnly == footerOnly &&
                                              observed.completedDismissVisible == (count != 0u) && observed.completedDiagnosticsMoreVisible == more &&
                                              observed.completedVisibleActionCount == count && observed.completedActionRowFits &&
                                              ! observed.taskHasVisibleButtonOverlap && ! observed.completedFailedItemsActionVisible && issueContent;
                            directIssuesTransitionsFit = directIssuesTransitionsFit && fits;
                            ++directIssuesTransitionCount;
                            AppendLog(
                                std::format(L"DirectIssuesTransition: dpi={} width={} stage={} issues={} more={} count={} collapsed={} footerOnly={} fit={}",
                                            dpi,
                                            widthDip,
                                            stage,
                                            observed.completedIssuesVisible,
                                            observed.completedDiagnosticsMoreVisible,
                                            observed.completedVisibleActionCount,
                                            observed.taskCollapsed,
                                            observed.footerOnly,
                                            fits));
                        };
                        current.hasProjectableIssues = false;
                        inspectActions(L"row-removed-counters-retained", false, true, 2u, false, false);
                        current.hasProjectableIssues = true;
                        inspectActions(L"row-arrived", true, true, 3u, false, false);
                        current.warningCount = 0u;
                        current.errorCount   = 0u;
                        inspectActions(L"row-without-counters", true, true, 3u, false, false);
                        current.operation = FILESYSTEM_DELETE;
                        inspectActions(L"issues-without-more", true, false, 2u, false, false);
                        current.operation = scenario.tasks.front().operation;
                        static_cast<void>(DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, current.taskId, 0u}));
                        inspectActions(L"collapsed", false, false, 0u, true, false);
                        static_cast<void>(DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, current.taskId, 0u}));
                        inspectActions(L"expanded", true, true, 3u, false, false);
                        operations->SetPopupFooterOnly(true);
                        inspectActions(L"footer-hidden", false, false, 0u, false, true);
                        operations->SetPopupFooterOnly(false);
                        inspectActions(L"footer-restored", true, true, 3u, false, false);
                        current.hasProjectableIssues = false;
                        inspectActions(L"row-removed-zero-counters", false, true, 2u, false, false);
                        current.hasProjectableIssues = true;
                        current.warningCount         = scenario.tasks.front().warningCount;
                        current.errorCount           = scenario.tasks.front().errorCount;
                        inspectActions(L"original-restored", true, true, 3u, false, false);
                    }
                    if (scenario.menu != PopupHitTest::Kind::None)
                    {
                        // Menu dispatch follows the window-procedure convention of returning zero;
                        // the newly created owned flyout, rather than that value, proves readiness.
                        static_cast<void>(DebugInvokeFileOperationsPopup(popup.get(), {scenario.menu, 1u, 0u}));
                        struct MenuLookup
                        {
                            HWND owner;
                            HWND menu = nullptr;
                        } lookup{popup.get()};
                        EnumThreadWindows(GetCurrentThreadId(),
                                          [](HWND candidate, LPARAM context) noexcept -> BOOL
                        {
                            auto& found = *reinterpret_cast<MenuLookup*>(context);
                            if (GetWindow(candidate, GW_OWNER) == found.owner && IsWindowVisible(candidate))
                                found.menu = candidate;
                            return TRUE;
                        },
                                          reinterpret_cast<LPARAM>(&lookup));
                        if (! lookup.menu)
                        {
                            Fail(std::format(L"Visual gallery menu missing: {}", scenario.name));
                            return false;
                        }
                        RedrawWindow(lookup.menu, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                        const std::wstring menuFilename = std::format(
                            L"{}__{}__{}dip__fr-FR__display{}__{}dpi__flyout.png", scenario.name, themeNames[themeIndex], widthDip, display + 1u, dpi);
                        const HRESULT menuCaptured = RedSalamander::TestSupport::SaveWindowScreenshot(lookup.menu, output / menuFilename);
                        // Dismiss only this owned test menu. No desktop input is synthesized.
                        SendMessageW(lookup.menu, WM_KEYDOWN, VK_ESCAPE, 0);
                        if (FAILED(menuCaptured))
                        {
                            Fail(L"Visual gallery menu capture failed.");
                            return false;
                        }
                        manifest << narrow(menuFilename) << '\t' << narrow(scenario.name) << '\t' << narrow(themeNames[themeIndex]) << '\t' << widthDip << '\t'
                                 << dpi << "\tproduction-menu-synthetic-owner\n";
                    }
                    if (scenario.name == L"04-queued-start-now")
                    {
                        // Keep this exact HWND/task and its height history; the gallery setter clears animation state.
                        auto motionTheme                  = theme;
                        motionTheme.reducedMotionOverride = false;
                        folder->ApplyTheme(motionTheme);
                        auto restoreMotion  = wil::scope_exit([&] { folder->ApplyTheme(theme); });
                        auto& task          = popupState->debugVisualScenario.value().front();
                        const float scale   = static_cast<float>(dpi) / 96.0f;
                        float runningHeight = 0.0f;
                        for (unsigned phase = 0u; phase < 12u; ++phase)
                        {
                            if (phase == 7u || phase == 9u)
                            {
                                static_cast<void>(DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, task.taskId, 0u}));
                                RedSalamander::TestSupport::MessagePumpWaitOptions collapseOptions{};
                                collapseOptions.timeout            = std::chrono::milliseconds(1500);
                                collapseOptions.maxMessagesPerPump = 16u;
                                const ULONGLONG collapseStart      = GetTickCount64();
                                const auto collapsed               = RedSalamander::TestSupport::PumpMessagesUntil(
                                    [&]() noexcept
                                {
                                    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                                    PopupLayoutDebugSnapshot measured{};
                                    measured.taskId = task.taskId;
                                    return DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), measured) && measured.taskCollapsed &&
                                           GetTickCount64() - collapseStart >= 320u && ! measured.selectedTaskHeightMotionActive &&
                                           ! measured.autoResizeMotionActive;
                                },
                                    collapseOptions);
                                waitingTransitionsFit = waitingTransitionsFit && collapsed.conditionMet;
                                // Expand the same task and inspect its first paint, for both waiting and running content.
                                static_cast<void>(DebugInvokeFileOperationsPopup(popup.get(), {PopupHitTest::Kind::TaskToggleCollapse, task.taskId, 0u}));
                            }
                            const bool admissionWait  = phase >= 4u && phase < 8u;
                            task.started              = ! admissionWait;
                            task.waitingInQueue       = admissionWait;
                            task.waitingForOthers     = false;
                            task.queuePaused          = phase == 3u;
                            task.paused               = phase == 1u;
                            task.operationStartTick   = task.started ? GetTickCount64() - 25'000u : 0u;
                            task.statusKind           = TaskSnapshot::StatusKind::None;
                            task.totalBytes           = phase == 5u ? 0u : 8ull * 1024ull * 1024ull * 1024ull;
                            task.autoConcurrencyUsed  = phase == 6u;
                            task.autoTunedConcurrency = phase == 6u ? 4u : 0u;
                            // Q7: four streams still publish one Source/Destination region pair.
                            task.inFlightFileCount = phase == 10u ? 4u : phase == 11u ? 1u : 0u;
                            for (size_t stream = 0u; stream < task.inFlightFileCount; ++stream)
                            {
                                auto& file            = task.inFlightFiles[stream];
                                file.progressStreamId = stream + 1u;
                                file.sourcePath       = std::format(L"C:\\Photographies\\Photographie familiale numéro {}.raw", stream + 1u);
                                file.totalBytes       = 512ull * 1024ull * 1024ull;
                                file.completedBytes   = 64ull * 1024ull * 1024ull;
                                file.lastUpdateTick   = GetTickCount64();
                            }
                            const ULONGLONG start = GetTickCount64();
                            size_t frames         = 0u;
                            bool phaseFits        = true;
                            RedSalamander::TestSupport::MessagePumpWaitOptions options{};
                            options.timeout = std::chrono::milliseconds(1500);
                            // Animation paints may replenish the queue. Observe geometry between bounded batches.
                            options.maxMessagesPerPump = 16u;
                            const auto settled         = RedSalamander::TestSupport::PumpMessagesUntil(
                                [&]() noexcept
                            {
                                RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                                PopupLayoutDebugSnapshot measured{};
                                measured.taskId    = task.taskId;
                                const bool read    = DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), measured);
                                const float height = (measured.taskPaintedCardRect.bottom - measured.taskPaintedCardRect.top) / scale;
                                const bool graph   = measured.graphPlotRect.bottom > measured.graphPlotRect.top;
                                const bool fits =
                                    read && measured.taskPaintedBodyBoundsValid && ! measured.taskHasVisibleButtonOverlap &&
                                    measured.taskPaintedBodyBottomPx + 4.0f * scale <= measured.taskPaintedActionRowRect.top &&
                                    (! admissionWait || measured.taskPaintedActionRowRect.top - measured.taskPaintedBodyBottomPx <= 16.0f * scale) &&
                                    measured.taskCancelVisible && measured.taskCancelHitTargetActive &&
                                    ((! admissionWait && phase != 3u) ? graph && measured.graphPlotRect.bottom <= measured.taskPaintedActionRowRect.top
                                                                      : ! graph) &&
                                    measured.taskPaintedPathCount == 2u && measured.taskPaintedPathsFit &&
                                    measured.taskPaintedActionRowRect.bottom <= measured.taskPaintedCardRect.bottom &&
                                    (admissionWait ? height < runningHeight - 80.0f : height >= 244.0f);
                                if (phase == 0u)
                                    runningHeight = height;
                                phaseFits = phaseFits && fits;
                                ++frames;
                                return frames >= 2u && GetTickCount64() - start >= 320u && ! measured.selectedTaskHeightMotionActive &&
                                       ! measured.autoResizeMotionActive;
                            },
                                options);
                            waitingTransitionsFit = waitingTransitionsFit && phaseFits && settled.conditionMet && frames >= 2u;
                            AppendLog(std::format(L"WaitingCardTransition: dpi={} width={} phase={} frames={} fit={} settled={} elapsedMs={} messages={}",
                                                  dpi,
                                                  widthDip,
                                                  phase,
                                                  frames,
                                                  phaseFits,
                                                  settled.conditionMet,
                                                  settled.elapsed.count(),
                                                  settled.dispatchedMessageCount));
                        }
                    }
                }
            }
        }
        auto promptTheme                  = ResolveAppTheme(ThemeMode::Dark, L"FileOperationsGallery");
        promptTheme.reducedMotionOverride = true;
        folder->ApplyTheme(promptTheme);
        wil::unique_hwnd promptOwner(FileOperationsPopup::Create(operations, folder, state.mainWindow, lifetime));
        if (! promptOwner)
            return false;
        PlaceFileOpsGalleryWindow(promptOwner.get(), monitor, 600, 400);
        FileOpsGalleryPromptCapture promptCapture{output, &manifest, false, display, &monitors};
        g_fileOpsGalleryPromptCapture = &promptCapture;
        const UINT_PTR promptTimer    = SetTimer(nullptr, 0u, 100u, CaptureFileOpsGalleryPrompt);
        auto resetPrompt              = wil::scope_exit([&]
        {
            if (promptTimer)
                KillTimer(nullptr, promptTimer);
            g_fileOpsGalleryPromptCapture = nullptr;
        });
        if (! promptTimer)
            return false;
        DebugShowFileOperationsSpeedLimitPromptForGallery(promptOwner.get(), promptTheme);
        if (! promptCapture.succeeded)
            Fail(L"Visual gallery custom speed prompt capture failed.");
        if (! promptCapture.succeeded || ! manifest.good())
            return false;
    }
    if (pathDetailCaptureSets != monitors.size() * 2u || pathDetailDarkSets != monitors.size() || pathDetailLightSets != monitors.size() ||
        pathDetailCaptureCount != pathDetailCaptureSets * 2u)
    {
        Fail(L"Expected exactly one root/scrolled detail pair per actual display in dark and light.");
        return false;
    }
    const bool supplemental = CaptureFileOpsGallerySupplemental(state, output, manifest, monitors);
    if (scanStatusPresentWitnessCount != 15u * monitors.size() || scanStatusAbsentWitnessCount != 8u * monitors.size())
    {
        Fail(L"The gallery must verify all informational scan stages and terminal/absent rows on every native display.");
        return false;
    }
    if (scopeProviderWitnessCount != 4u * monitors.size())
    {
        Fail(L"The gallery must verify absent, unchecked and checked scope providers on every native display.");
        return false;
    }
    if (! pathLabelsFit)
        Fail(L"French path labels must fit their actual painted rectangles without overlapping the accessible path; see PathLabelFit diagnostics.");
    if (! deletionPresentationFits)
        Fail(L"Deletion must expose its item rate without an empty byte-throughput graph or graph disclosure; see DeletePresentation diagnostics.");
    if (! footerSummariesFit)
        Fail(L"French mixed-state footer summaries must wrap without clipping or covering commands in full and footer-only modes; see FooterSummaryFit "
             L"diagnostics.");
    if (! abbreviatedPathsFit)
        Fail(L"Long French/Unicode filenames must fit their painted path rows, retain the extension and expose the complete accessible path.");
    if (! statusChipLabelsFit || statusChipWitnessCount != 6u * monitors.size())
        Fail(L"Elapsed Discovering/Preparing status chips must paint their short French labels entirely inside the actual chrome while the header retains "
             L"elapsed detail.");
    if (! terminalFactsFit || terminalFactsWitnessCount != 4u * monitors.size())
        Fail(L"Mixed terminal Move must paint all nine localized publication/source/verification facts in one hosted region after the result summary and "
             L"before actions at both widths and both tested themes.");
    if (! terminalFactsTransitionsFit || terminalFactsTransitionCount != kTerminalFactsTransitionsPerMonitor * monitors.size())
        Fail(L"Terminal facts must update and retire on one HWND across count, collapse, footer, settings-change, and empty-summary transitions; see "
             L"TerminalFactsTransition diagnostics.");
    if (! directIssuesFit || directIssuesWitnessCount != 8u * monitors.size())
        Fail(L"Direct Issues must follow actual projectable rows, while counter-only diagnostics retain More and clean Delete retains Dismiss; all hosted "
             L"terminal action labels, hit targets, and bounds must fit at both widths and tested themes.");
    if (! directIssuesTransitionsFit || directIssuesTransitionCount != 10u * monitors.size())
        Fail(L"Direct Issues must arrive and retire on one popup HWND independently of warning counters and hide for collapsed/footer-only views; see "
             L"DirectIssuesTransition diagnostics.");
    if (! notStartedFit || notStartedWitnessCount != 2u * monitors.size())
        Fail(L"A Delete canceled before it started must paint its \"Not started; nothing was changed\" result line without Source facts at both widths; "
             L"see NotStartedFit diagnostics.");
    if (! verificationHalfFit || verificationHalfWitnessCount != 2u * monitors.size())
        Fail(L"A finished Verify On card must keep its verification half visible (divider and outlined remainder), and its compact row must read 100% "
             L"only when every byte was verified; see VerificationHalfFit diagnostics.");
    if (! terminalPresentationFits)
        Fail(L"Completed-group actions and badges, and full expanded terminal results, must remain readable without clipping.");
    if (! completedSemanticsFit || doneHeadingWitnessCount != 4u * monitors.size() || groupBadgeWitnessCount != 2u * monitors.size())
        Fail(L"Done headings must not repeat the painted status chip; all five completed-group badges must draw and retain distinct Partial/Unknown hosted "
             L"names and glyphs.");
    if (! hostedProgressValuesMatch || hostedProgressWitnessCount != 10u * monitors.size())
        Fail(L"The retained hosted task ProgressBar must expose the whole-task value, indeterminate state, and transfer/verification segments; see "
             L"HostedProgressValue diagnostics.");
    if (! speedLimitLabelsFit)
        Fail(L"The full French speed-limit label must fit the actual Selector text lane, including both chevron-balancing margins.");
    if (! waitingCardsFit || ! waitingTransitionsFit)
        Fail(L"Waiting cards must use only visible body/action space and retain safe same-task phase transitions; see WaitingCardFit/Transition diagnostics.");
    return statusChipLabelsFit && statusChipWitnessCount == 6u * monitors.size() && hostedProgressValuesMatch &&
           hostedProgressWitnessCount == 10u * monitors.size() && completedSemanticsFit && doneHeadingWitnessCount == 4u * monitors.size() &&
           groupBadgeWitnessCount == 2u * monitors.size() && terminalPresentationFits && terminalFactsFit &&
           terminalFactsWitnessCount == 4u * monitors.size() && terminalFactsTransitionsFit &&
           terminalFactsTransitionCount == kTerminalFactsTransitionsPerMonitor * monitors.size() && directIssuesFit &&
           directIssuesWitnessCount == 8u * monitors.size() && directIssuesTransitionsFit && directIssuesTransitionCount == 10u * monitors.size() &&
           notStartedFit && notStartedWitnessCount == 2u * monitors.size() && verificationHalfFit && verificationHalfWitnessCount == 2u * monitors.size() &&
           speedLimitLabelsFit && waitingCardsFit && waitingTransitionsFit && supplemental && manifest.good() && pathLabelsFit && deletionPresentationFits &&
           footerSummariesFit && abbreviatedPathsFit;
}
