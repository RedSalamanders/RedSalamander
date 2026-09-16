#include "FolderWindow.FileOperations.SelfTest.VisualGallerySurfaces.cpp"

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
    task                = base;
    task.lifecyclePhase = static_cast<uint8_t>(Phase::Stopping);
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
    task                        = base;
    task.currentSourcePath      = L"C:\\Projects\\International archive\\Very long collection title for the September exhibition\\Photographs\\東京 — été — "
                                  L"München\\IMG_2026_0916_final_edited_version.raw";
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
    constexpr std::array<std::wstring_view, 24> names{{L"file-exists",           L"read-only-file",
                                                       L"read-only-exists",      L"type-mismatch",
                                                       L"destination-link",      L"name-not-representable",
                                                       L"target-conflict",       L"sharing-violation",
                                                       L"access-denied",         L"disk-full",
                                                       L"path-too-long",         L"recycle-failed",
                                                       L"insufficient-space",    L"space-unknown",
                                                       L"efs-plaintext",         L"sparse-inflation",
                                                       L"placeholder-hydration", L"metadata-loss",
                                                       L"network-offline",       L"unsupported-reparse",
                                                       L"same-host-overlap",     L"same-host-live-output",
                                                       L"permanent-delete",      L"unknown-error"}};
    static_assert(names.size() == static_cast<size_t>(FileOperations::ConflictClass::Count));
    for (size_t i = 0u; i < names.size(); ++i)
    {
        task                 = base;
        task.conflict.active = true;
        task.conflict.bucket = static_cast<uint8_t>(i);
        constexpr std::array<DWORD, 24> errors{{ERROR_FILE_EXISTS,
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
                                                ERROR_GEN_FAILURE}};
        task.conflict.status                        = HRESULT_FROM_WIN32(errors[i]);
        task.conflict.sourcePath                    = base.currentSourcePath;
        task.conflict.destinationPath               = base.currentDestinationPath;
        task.conflict.sourceMetadata.available      = true;
        task.conflict.sourceMetadata.sizeKnown      = true;
        task.conflict.sourceMetadata.sizeBytes      = 512ull * 1024ull * 1024ull;
        task.conflict.sourceMetadata.lastWriteTime  = 134'338'680'000'000'000ll;
        task.conflict.destinationMetadata           = task.conflict.sourceMetadata;
        task.conflict.destinationMetadata.sizeBytes = 480ull * 1024ull * 1024ull;
        task.conflict.destinationMetadata.lastWriteTime -= 864'000'000'000ll;
        task.conflict.overlapTaskId      = 2u;
        task.conflict.overlapProblem     = static_cast<uint8_t>(FileOperations::SameHostOverlapProblem::RemovesRead);
        task.conflict.consentDetail      = L"128 fichiers du dossier C:\\Photographies\\Vacances en famille. Cette opération est irréversible.";
        task.conflict.factItemCountKnown = i >= 11u;
        task.conflict.factItemCount      = 128u;
        task.conflict.factBytesKnown     = i >= 11u;
        task.conflict.factBytes          = base.totalBytes;
        task.conflict.deferredConsent    = i >= 11u && i != 18u && i != 19u && i != 23u;
        if (i == 11u || i == 22u)
            task.operation = FILESYSTEM_DELETE;
        if (i == 17u)
            task.operation = FILESYSTEM_MOVE;
        add(std::format(L"{:02}-decision-{}", 33u + i, names[i]), task);
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
        detail.changeAttributesApplying       = true;
        detail.changeAttributesPlannedItems   = 128u;
        detail.changeAttributesCompletedItems = 48u;
        detail.changeAttributesCurrentPath    = base.currentSourcePath;
        detail.makeFileListRendering          = true;
        detail.makeFileListTotalEntries       = 128u;
        detail.makeFileListRenderedEntries    = 48u;
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
    task             = waiting;
    task.taskId      = 1u;
    task.queuePaused = true;
    addVariant(L"queue-paused", task);
    task                  = base;
    task.waitingForOthers = true;
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
    task.informational.changeCaseCurrentPath      = base.currentSourcePath;
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
    auto destination = result[3];
    destination.name = L"107-destination-history";
    destination.menu = FileOperationsPopupInternal::PopupHitTest::Kind::TaskDestination;
    result.push_back(std::move(destination));
    auto longDecision = result[32];
    longDecision.name = L"108-decision-long-french-paths";
    const std::wstring longName =
        L"Photographie panoramique de la réunion familiale — version définitive retouchée pour impression grand format et archivage patrimonial.raw";
    longDecision.tasks.front().conflict.sourcePath =
        L"C:\\Photographies\\Vacances en famille — septembre 2026\\Sélection définitive pour impression et archivage\\" + longName;
    longDecision.tasks.front().conflict.destinationPath =
        L"D:\\Sauvegardes\\Archives photographiques personnelles\\Exposition annuelle de la médiathèque\\" + longName;
    result.push_back(std::move(longDecision));
    auto& longTask                  = result.back().tasks.front();
    longTask.currentSourcePath      = longTask.conflict.sourcePath;
    longTask.currentDestinationPath = longTask.conflict.destinationPath;
    longTask.destinationFolder      = std::filesystem::path(longTask.conflict.destinationPath).parent_path();
    for (const size_t index : {21u, 22u})
    {
        auto recovery = result[index];
        recovery.name = index == 21u ? L"114-recovery-known-source-menu" : L"115-recovery-unknown-source-menu";
        recovery.menu = FileOperationsPopupInternal::PopupHitTest::Kind::TaskCompletedMore;
        result.push_back(std::move(recovery));
    }
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
    for (size_t display = 0u; display < monitors.size(); ++display)
    {
        renderer->DebugSetVisualScenario(scenarios[32].tasks, false, true);
        PlaceFileOpsGalleryWindow(popup.get(), monitors[display], 760, 560);
        if (! RedSalamander::TestSupport::DirectedSelfTestInputWarning::TryActivateOwnedWindow(popup.get()))
        {
            Fail(L"Interactive gallery could not acquire its own foreground window; no input was sent.");
            return false;
        }
        RedSalamander::TestSupport::MessagePumpWaitOptions ready{};
        ready.timeout  = std::chrono::milliseconds(1500);
        const auto key = [&](WORD virtualKey, bool shift)
        {
            if (GetForegroundWindow() != popup.get())
                return false;
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
            return SendInput(count, input.data(), sizeof(INPUT)) == count;
        };
        PopupLayoutDebugSnapshot focused{};
        focused.taskId = 1u;
        if (! key(VK_TAB, false))
            return false;
        const auto focusReady = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept { return DebugGetFileOperationsPopupLayoutSnapshot(popup.get(), focused) && ! focused.hostedFocusedAutomationId.empty(); }, ready);
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
                   focused.hostedFocusedAutomationId != firstFocus;
        },
            ready);
        if (! reverseReady.conditionMet ||
            ! SaveFileOpsGallerySurface(popup.get(), output, manifest, L"110-keyboard-shift-tab-focus", display, L"real-foreground-keyboard-input"))
            return false;
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
        ClientToScreen(popup.get(), &point);
        const HWND pointerTarget = WindowFromPoint(point);
        if (pointerTarget != popup.get() && ! IsChild(popup.get(), pointerTarget))
        {
            Fail(L"Interactive gallery pointer target is not owned by the popup; no pointer input was sent.");
            return false;
        }
        SetCursorPos(point.x, point.y);
        const auto hovered = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept { return renderer->DebugVisualButtonHovered(PopupHitTest::Kind::FooterOptions); }, ready);
        if (! hovered.conditionMet)
        {
            Fail(L"Interactive gallery did not reach the target button's real hover state.");
            return false;
        }
        if (! SaveFileOpsGallerySurface(popup.get(), output, manifest, L"111-pointer-hover", display, L"real-pointer-hover"))
            return false;
        if (GetForegroundWindow() != popup.get())
            return false;
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
    PlaceFileOpsGalleryWindow(popup.get(), monitors.front(), 760, 560);
    return SaveFileOpsGallerySurface(popup.get(), output, manifest, L"113-dpi-return-to-first-display", 0u, L"same-HWND-round-trip-DPI");
}

struct FileOpsGalleryPromptCapture final
{
    std::filesystem::path output;
    std::ofstream* manifest = nullptr;
    bool succeeded          = false;
    size_t display          = 0u;
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
    }
    static_cast<void>(DebugCancelFileOperationsSpeedLimitPrompt());
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
                if (themeIndex != 0u && sceneIndex != 9u && sceneIndex != 30u && sceneIndex != 32u && sceneIndex != 54u)
                    continue;
                const auto& scenario = scenarios[sceneIndex];
                for (const int widthDip : {480, 760})
                {
                    if (widthDip == 760 && sceneIndex != 9u && sceneIndex != 26u && sceneIndex != 29u && sceneIndex != 32u &&
                        scenario.name != L"108-decision-long-french-paths")
                        continue;
                    operations->SetPopupFooterOnly(scenario.footerOnly);
                    operations->SetPopupCompactDensity(scenario.compactDensity);
                    if (scenario.graphicsFailure)
                        DebugFailNextFileOperationsD2DTargetAttempts(100u);
                    auto resetFailureInjection = wil::scope_exit([] { DebugFailNextFileOperationsD2DTargetAttempts(0u); });
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
                    popupState->debugVisualDestinationHistory = {
                        L"D:\\Sauvegardes\\Archives photographiques personnelles\\Sélection définitive pour impression et archivage",
                        L"D:\\Sauvegardes\\Exposition annuelle de la médiathèque\\Photographies retouchées en très haute résolution"};
                    PlaceFileOpsGalleryWindow(popup.get(), monitor, widthDip, scenario.footerOnly ? 88 : 720);
                    const UINT dpi = GetDpiForWindow(popup.get());
                    ShowWindow(popup.get(), SW_SHOWNOACTIVATE);
                    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
                    // A second paint reflects the stable auto-size and retained-control rectangles.
                    RedrawWindow(popup.get(), nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
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
        FileOpsGalleryPromptCapture promptCapture{output, &manifest, false, display};
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
    return CaptureFileOpsGallerySupplemental(state, output, manifest, monitors) && manifest.good();
}
