// Included by the FileOps runner. Owned storage only; no activation, UIA or input.
[[nodiscard]] bool RunFileOpsSameHostOverlapClassifierAssertions(const std::vector<FileOperations::MutationInterlockScope>& retainedPublishScopes,
                                                                 const std::vector<FileOperations::MutationInterlockScope>& disjointPublishScopes) noexcept
{
    constexpr size_t kSelectionCount              = 256u;
    constexpr uint64_t kCartesianScopeComparisons = kSelectionCount * kSelectionCount;
    const auto reject                             = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Same-host overlap classifier: {}", reason));
        return false;
    };
    if (retainedPublishScopes.size() != kSelectionCount || disjointPublishScopes.size() != kSelectionCount)
        return reject(std::format(L"expected two 256-scope sets, received {} and {}", retainedPublishScopes.size(), disjointPublishScopes.size()));

    bool byNameDelete            = false;
    uint64_t classifierPairs     = 0u;
    const auto classifierStarted = std::chrono::steady_clock::now();
    const FileOperations::SameHostOverlapProblem disjointProblem =
        FileOperations::DebugClassifySameHostOverlapForSelfTest(retainedPublishScopes, disjointPublishScopes, byNameDelete, classifierPairs);
    const uint64_t classifierUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - classifierStarted).count());
    Debug::Perf::Emit(L"FileOps.Interlock.ClassifierUs", L"selftest-local-256x256-disjoint-publish-leaves", classifierUs, classifierPairs, 1u, S_OK);
    if (disjointProblem != FileOperations::SameHostOverlapProblem::None || byNameDelete || classifierPairs == 0u ||
        classifierPairs > kCartesianScopeComparisons)
        return reject(std::format(L"disjoint case returned problem={} byNameDelete={} comparisons={} maximum={}",
                                  static_cast<unsigned int>(disjointProblem),
                                  byNameDelete,
                                  classifierPairs,
                                  kCartesianScopeComparisons));

    std::vector<FileOperations::MutationInterlockScope> lastPairOverlapScopes = disjointPublishScopes;
    lastPairOverlapScopes.back()                                              = retainedPublishScopes.back();
    bool lastPairByNameDelete                                                 = false;
    uint64_t lastPairComparisons                                              = 0u;
    const auto lastPairStarted                                                = std::chrono::steady_clock::now();
    const FileOperations::SameHostOverlapProblem lastPairProblem =
        FileOperations::DebugClassifySameHostOverlapForSelfTest(retainedPublishScopes, lastPairOverlapScopes, lastPairByNameDelete, lastPairComparisons);
    const uint64_t lastPairUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - lastPairStarted).count());
    Debug::Perf::Emit(L"FileOps.Interlock.ClassifierUs", L"selftest-local-256x256-last-pair-publish-overlap", lastPairUs, lastPairComparisons, 1u, S_OK);
    if (lastPairProblem != FileOperations::SameHostOverlapProblem::SameNames || lastPairByNameDelete || lastPairComparisons == 0u ||
        lastPairComparisons > kCartesianScopeComparisons)
        return reject(std::format(L"last-pair case returned problem={} byNameDelete={} comparisons={} maximum={}",
                                  static_cast<unsigned int>(lastPairProblem),
                                  lastPairByNameDelete,
                                  lastPairComparisons,
                                  kCartesianScopeComparisons));

    return true;
}

#if defined(NDEBUG)
struct FileOpsOverlapClassifierObservation
{
    FileOperations::SameHostOverlapProblem problem = FileOperations::SameHostOverlapProblem::None;
    bool byNameDelete                              = false;
    uint64_t pairComparisons                       = 0u;
};

[[nodiscard]] FileOpsOverlapClassifierObservation ObserveFileOpsOverlapClassifier(const std::vector<FileOperations::MutationInterlockScope>& mine,
                                                                                  const std::vector<FileOperations::MutationInterlockScope>& theirs) noexcept
{
    FileOpsOverlapClassifierObservation observation{};
    observation.problem = FileOperations::DebugClassifySameHostOverlapForSelfTest(mine, theirs, observation.byNameDelete, observation.pairComparisons);
    return observation;
}

[[nodiscard]] bool SameFileOpsOverlapClassifierObservation(const FileOpsOverlapClassifierObservation& left,
                                                           const FileOpsOverlapClassifierObservation& right) noexcept
{
    return left.problem == right.problem && left.byNameDelete == right.byNameDelete && left.pairComparisons == right.pairComparisons;
}

[[nodiscard]] bool RunFileOpsSameHostOverlapClassifierReleaseComparison(const std::vector<FileOperations::MutationInterlockScope>& candidateRetainedScopes,
                                                                        const std::vector<FileOperations::MutationInterlockScope>& candidateDisjointScopes,
                                                                        std::wstring& failure) noexcept
{
    using Observation = FileOpsOverlapClassifierObservation;
    failure.clear();
    const auto reject = [&](std::wstring_view message) noexcept
    {
        failure.assign(message);
        return false;
    };
    const auto makeKeyless = [](const std::vector<FileOperations::MutationInterlockScope>& keyed)
    {
        std::vector<FileOperations::MutationInterlockScope> keyless = keyed;
        for (FileOperations::MutationInterlockScope& scope : keyless)
        {
            scope.providerPathKey.reset();
            for (FileOperations::MutationInterlockAnchor& anchor : scope.anchors)
            {
                anchor.relativePathKey.reset();
            }
        }
        return keyless;
    };

    std::vector<FileOperations::MutationInterlockScope> candidateLastPair = candidateDisjointScopes;
    std::vector<FileOperations::MutationInterlockScope> keylessRetained   = makeKeyless(candidateRetainedScopes);
    std::vector<FileOperations::MutationInterlockScope> keylessDisjoint   = makeKeyless(candidateDisjointScopes);
    std::vector<FileOperations::MutationInterlockScope> keylessLastPair   = keylessDisjoint;
    if (candidateLastPair.empty() || keylessLastPair.empty())
    {
        return reject(L"Release comparison needs the retained 256-scope pair vectors");
    }
    candidateLastPair.back() = candidateRetainedScopes.back();
    keylessLastPair.back()   = keylessRetained.back();

    const std::array<std::wstring_view, 2> scenarioNames{L"disjoint", L"last-pair"};
    const std::array<const std::vector<FileOperations::MutationInterlockScope>*, 2> candidateTheirs{std::addressof(candidateDisjointScopes),
                                                                                                    std::addressof(candidateLastPair)};
    const std::array<const std::vector<FileOperations::MutationInterlockScope>*, 2> keylessTheirs{std::addressof(keylessDisjoint),
                                                                                                  std::addressof(keylessLastPair)};
    const std::array<Observation, 2> expected{{
        {FileOperations::SameHostOverlapProblem::None, false, 65'536u},
        {FileOperations::SameHostOverlapProblem::SameNames, false, 65'536u},
    }};

    // One unmeasured invocation of each implementation/scenario warms code and data before paired samples.
    for (size_t scenario = 0u; scenario < scenarioNames.size(); ++scenario)
    {
        Observation candidateWarmup{};
        Observation keylessWarmup{};
        if ((scenario % 2u) == 0u)
        {
            candidateWarmup = ObserveFileOpsOverlapClassifier(candidateRetainedScopes, *candidateTheirs[scenario]);
            keylessWarmup   = ObserveFileOpsOverlapClassifier(keylessRetained, *keylessTheirs[scenario]);
        }
        else
        {
            keylessWarmup   = ObserveFileOpsOverlapClassifier(keylessRetained, *keylessTheirs[scenario]);
            candidateWarmup = ObserveFileOpsOverlapClassifier(candidateRetainedScopes, *candidateTheirs[scenario]);
        }
        if (! SameFileOpsOverlapClassifierObservation(candidateWarmup, expected[scenario]) ||
            ! SameFileOpsOverlapClassifierObservation(keylessWarmup, expected[scenario]))
        {
            return reject(L"Release comparison warmup changed the established classifier result or pair count");
        }
    }

    const auto measure = [&](const std::vector<FileOperations::MutationInterlockScope>& mine,
                             const std::vector<FileOperations::MutationInterlockScope>& theirs,
                             const std::wstring_view scenario,
                             const size_t sample,
                             const bool candidateFirst,
                             const bool keyed) noexcept
    {
        const auto startedAt          = std::chrono::steady_clock::now();
        const Observation observation = ObserveFileOpsOverlapClassifier(mine, theirs);
        const uint64_t elapsedUs =
            static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - startedAt).count());
        const std::wstring detail     = std::format(L"release-{}-sample-{:02}-{}-{}",
                                                    scenario,
                                                    sample,
                                                    candidateFirst ? L"candidate-first" : L"keyless-first",
                                                    keyed ? L"candidate" : L"keyless-fallback");
        const uint64_t encodedOutcome = (static_cast<uint64_t>(observation.problem) << 1u) | (observation.byNameDelete ? 1u : 0u);
        Debug::Perf::Emit(L"FileOps.Interlock.ClassifierKeySampleUs", detail, elapsedUs, observation.pairComparisons, encodedOutcome, S_OK);
        return observation;
    };

    constexpr size_t kPairedSamples = 5u;
    for (size_t scenario = 0u; scenario < scenarioNames.size(); ++scenario)
    {
        for (size_t sample = 0u; sample < kPairedSamples; ++sample)
        {
            const bool candidateFirst = (sample % 2u) == 0u;
            const Observation candidate =
                candidateFirst ? measure(candidateRetainedScopes, *candidateTheirs[scenario], scenarioNames[scenario], sample, candidateFirst, true)
                               : Observation{};
            const Observation keyless = measure(keylessRetained, *keylessTheirs[scenario], scenarioNames[scenario], sample, candidateFirst, false);
            const Observation orderedCandidate =
                candidateFirst ? candidate
                               : measure(candidateRetainedScopes, *candidateTheirs[scenario], scenarioNames[scenario], sample, candidateFirst, true);
            if (! SameFileOpsOverlapClassifierObservation(orderedCandidate, expected[scenario]) ||
                ! SameFileOpsOverlapClassifierObservation(keyless, expected[scenario]) || ! SameFileOpsOverlapClassifierObservation(orderedCandidate, keyless))
            {
                return reject(std::format(L"Release {} sample {} changed result/by-name/pair-count parity", scenarioNames[scenario], sample));
            }
        }
    }
    return true;
}
#endif

[[nodiscard]] bool RunFileOpsSameHostOverlapClassifierCostChecks(SelfTestState& state) noexcept
{
    using Task                       = FolderWindow::FileOperationState::Task;
    constexpr size_t kSelectionCount = 256u;
    const auto reject                = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Same-host overlap classifier fixture: {}", reason));
        return false;
    };
    if (! state.fileOps || ! state.fsLocal || state.fileOps->HasActiveOperations())
        return reject(L"requires an idle File Operations owner and Local provider");

    const auto root            = state.tempRoot / L"review-same-host-overlap-classifier";
    const auto sourceRoot      = root / L"source";
    const auto destinationRoot = root / L"destination";
    const auto cleanup         = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    if (! RecreateEmptyDirectory(sourceRoot) || ! RecreateEmptyDirectory(destinationRoot))
        return reject(L"owned source and destination roots could not be created");

    FileOperations::QualifiedEndpoint localEndpoint{};
    if (! FileOperations::TryQualifyEndpointForSelfTest(state.fsLocal, sourceRoot.native(), FILESYSTEM_COPY, kPluginIdLocal, L"host/default", localEndpoint))
        return reject(L"the actual Local source path contract could not be qualified");

    FileOperations::TransferPlan plan{};
    plan.intent                         = FileOperations::TransferIntent::Copy;
    plan.strategy                       = FileOperations::OperationStrategy::Copy;
    plan.sourceEndpoint                 = localEndpoint;
    plan.destinationEndpoint            = localEndpoint;
    plan.destination.providerFolderPath = destinationRoot.native();
    plan.selectedItems.reserve(kSelectionCount);
    std::vector<std::filesystem::path> sources;
    sources.reserve(kSelectionCount);
    for (size_t index = 0u; index < kSelectionCount; ++index)
    {
        const std::filesystem::path source = sourceRoot / std::format(L"item-{:03}.bin", index);
        if (! SelfTest::WriteTextFile(source, "x"))
            return reject(std::format(L"could not create selected source {}", source.native()));
        plan.selectedItems.emplace_back(FileOperations::QualifiedSourceItem{.providerPath = source.native()});
        sources.emplace_back(source);
    }

    const auto plans = std::make_shared<const FileOperations::FileOperationPlanGroup>(
        FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(plan)}});
    auto taskStorage = std::make_unique<Task>(*state.fileOps);
    Task& task       = *taskStorage;
    task.StorePlans(plans);
    task._operation     = FILESYSTEM_COPY;
    task._executionMode = FolderWindow::FileOperationState::ExecutionMode::PerItem;
    task._fileSystem    = state.fsLocal;
    task._sourcePaths   = std::move(sources);
    task.InitializeAdmissionInputs(plans, {.fileSystem = state.fsLocal, .folder = destinationRoot, .pluginId = localEndpoint.pluginId});

    const auto prepareStarted = std::chrono::steady_clock::now();
    const HRESULT prepareHr   = task.PrepareMutationInterlockScopes();
    const auto scopeOwner     = task.LoadMutationInterlockScopes();
    if (FAILED(prepareHr) || ! scopeOwner)
        return reject(std::format(L"actual Local PrepareMutationInterlockScopes failed (hr=0x{:08X})", static_cast<unsigned long>(prepareHr)));

    std::vector<FileOperations::MutationInterlockScope> retainedPublishScopes;
    std::vector<FileOperations::MutationInterlockScope> disjointPublishScopes;
    retainedPublishScopes.reserve(kSelectionCount);
    disjointPublishScopes.reserve(kSelectionCount);
    for (const FileOperations::MutationInterlockScope& scope : *scopeOwner)
    {
        if (scope.access != FileOperations::MutationInterlockAccess::PublishDestination)
            continue;

        retainedPublishScopes.emplace_back(scope);
        FileOperations::MutationInterlockScope siblingScope = scope;
        siblingScope.providerPath                           = (destinationRoot / std::format(L"disjoint-{:03}.bin", disjointPublishScopes.size())).native();
        siblingScope.providerPathKey                        = TryMakePathKey(localEndpoint.pathIdentity.value(), siblingScope.providerPath);
        bool rebuiltAnchors                                 = true;
        for (FileOperations::MutationInterlockAnchor& anchor : siblingScope.anchors)
        {
            rebuiltAnchors         = rebuiltAnchors && anchor.authority &&
                                     TryGetFileSystemRelativePath(
                                         localEndpoint.pathIdentity.value(), anchor.authority->retained.providerPath, siblingScope.providerPath, anchor.relativePath);
            anchor.relativePathKey = TryMakePathKey(localEndpoint.pathIdentity.value(), anchor.relativePath);
            rebuiltAnchors         = rebuiltAnchors && anchor.relativePathKey.has_value();
        }
        if (! rebuiltAnchors)
            return reject(L"could not rebuild the Local retained-anchor suffix for a disjoint publication leaf");
        disjointPublishScopes.emplace_back(std::move(siblingScope));
    }
    const uint64_t fixturePrepareUs =
        static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - prepareStarted).count());
    Debug::Perf::Emit(L"FileOps.Interlock.ClassifierFixturePrepareUs",
                      L"selftest-local-256x256-retained-local-anchors",
                      fixturePrepareUs,
                      scopeOwner->size(),
                      kSelectionCount * 2u,
                      prepareHr);
    if (scopeOwner->size() != kSelectionCount * 2u || retainedPublishScopes.size() != kSelectionCount || disjointPublishScopes.size() != kSelectionCount)
        return reject(std::format(L"prepared scopes={} retained={} disjoint={} expected each set={}",
                                  scopeOwner->size(),
                                  retainedPublishScopes.size(),
                                  disjointPublishScopes.size(),
                                  kSelectionCount));

    AppendLog(L"Same-host overlap classifier fixture: Local retained anchors prepared outside classifier timing; 256 disjoint publication leaves and last-pair "
              L"overlap are ready.");
#if defined(NDEBUG)
    std::wstring releaseComparisonFailure;
    if (! RunFileOpsSameHostOverlapClassifierReleaseComparison(retainedPublishScopes, disjointPublishScopes, releaseComparisonFailure))
        return reject(std::format(L"Release keyed/keyless comparison failed: {}", releaseComparisonFailure));
#endif
    if (! RunFileOpsSameHostOverlapClassifierAssertions(retainedPublishScopes, disjointPublishScopes))
        return false;

    const auto classify = [](const std::vector<FileOperations::MutationInterlockScope>& left,
                             const std::vector<FileOperations::MutationInterlockScope>& right,
                             bool& byNameDelete) noexcept
    {
        uint64_t pairs = 0u;
        byNameDelete   = false;
        return FileOperations::DebugClassifySameHostOverlapForSelfTest(left, right, byNameDelete, pairs);
    };
    // The absent publication leaves share real retained Local ancestors. Rebuild every
    // suffix when retargeting; copied authority must never retain a stale leaf key.
    const auto supplementaryScope = [&](const std::wstring_view leaf, FileOperations::MutationInterlockScope& scope)
    {
        scope                                  = retainedPublishScopes.back();
        const FileSystemPathIdentity& identity = scope.endpoint.pathIdentity.value();
        scope.providerPath                     = JoinFileSystemPath(identity, destinationRoot.native(), leaf);
        scope.providerPathKey                  = TryMakePathKey(identity, scope.providerPath);
        if (! scope.providerPathKey.has_value())
            return false;
        for (FileOperations::MutationInterlockAnchor& anchor : scope.anchors)
        {
            if (! anchor.authority ||
                ! TryGetFileSystemRelativePath(identity, anchor.authority->retained.providerPath, scope.providerPath, anchor.relativePath))
                return false;
            anchor.relativePathKey = TryMakePathKey(identity, anchor.relativePath);
            if (! anchor.relativePathKey.has_value())
                return false;
        }
        return true;
    };
    FileOperations::MutationInterlockScope supplementaryLower;
    FileOperations::MutationInterlockScope supplementaryUpper;
    if (! supplementaryScope(L"x\U00010428z.bin", supplementaryLower) || ! supplementaryScope(L"X\U00010400Z.bin", supplementaryUpper))
        return reject(L"valid supplementary publication names must retain usable provider and anchor keys");
    bool supplementaryByName = false;
    if (classify({supplementaryLower}, {supplementaryUpper}, supplementaryByName) != FileOperations::SameHostOverlapProblem::None || supplementaryByName ||
        classify({supplementaryLower}, {supplementaryLower}, supplementaryByName) != FileOperations::SameHostOverlapProblem::SameNames || supplementaryByName)
        return reject(L"keyed Local supplementary siblings must stay distinct while identical leaves conflict");
    supplementaryLower.providerPathKey.reset();
    supplementaryUpper.providerPathKey.reset();
    for (FileOperations::MutationInterlockAnchor& anchor : supplementaryLower.anchors)
        anchor.relativePathKey.reset();
    for (FileOperations::MutationInterlockAnchor& anchor : supplementaryUpper.anchors)
        anchor.relativePathKey.reset();
    if (classify({supplementaryLower}, {supplementaryUpper}, supplementaryByName) != FileOperations::SameHostOverlapProblem::None || supplementaryByName)
        return reject(L"keyless Local supplementary siblings disagree with the keyed classifier");

    FileOperations::MutationInterlockScope write = retainedPublishScopes.back();
    write.access                                 = FileOperations::MutationInterlockAccess::WriteSource;
    FileOperations::MutationInterlockScope read  = retainedPublishScopes.back();
    read.access                                  = FileOperations::MutationInterlockAccess::ReadSource;
    bool byNameDelete                            = false;
    if (classify({read}, {read}, byNameDelete) != FileOperations::SameHostOverlapProblem::None || byNameDelete ||
        classify({write}, {read}, byNameDelete) != FileOperations::SameHostOverlapProblem::RemovesRead || byNameDelete)
        return reject(L"cached classifier changed mixed-access rules: read/read must pass and write/read must name its first conflict");

    FileOperations::MutationInterlockScope firstPublish = retainedPublishScopes[0];
    FileOperations::MutationInterlockScope firstRead    = firstPublish;
    firstRead.access                                    = FileOperations::MutationInterlockAccess::ReadSource;
    FileOperations::MutationInterlockScope secondWrite  = retainedPublishScopes[1];
    secondWrite.access                                  = FileOperations::MutationInterlockAccess::WriteSource;
    if (classify({firstPublish, secondWrite}, {firstRead, secondWrite}, byNameDelete) != FileOperations::SameHostOverlapProblem::ReplacesRead || byNameDelete)
        return reject(L"classifier no longer reports the first conflicting scope pair in immutable snapshot order");

    FileOperations::MutationInterlockScope byNameWriter = secondWrite;
    byNameWriter.byNameCorroborated                     = true;
    if (classify({firstPublish, secondWrite}, {firstRead, byNameWriter}, byNameDelete) != FileOperations::SameHostOverlapProblem::DeletedByNameWritten ||
        ! byNameDelete)
        return reject(L"delete-by-name no longer retains its documented precedence over an earlier ordinary overlap");

    FileOperations::MutationInterlockScope alias = retainedPublishScopes.back();
    alias.endpoint.rootId                        = L"distinct-root-same-bound-objects";
    if (classify({retainedPublishScopes.back()}, {alias}, byNameDelete) != FileOperations::SameHostOverlapProblem::SameNames || byNameDelete)
        return reject(L"cached classifier failed to preserve exact retained-anchor alias overlap across distinct root IDs");

    FileOperations::MutationInterlockScope legacyAlias = alias;
    legacyAlias.providerPathKey                        = std::nullopt;
    for (FileOperations::MutationInterlockAnchor& anchor : legacyAlias.anchors)
        anchor.relativePathKey = std::nullopt;
    if (classify({retainedPublishScopes.back()}, {legacyAlias}, byNameDelete) != FileOperations::SameHostOverlapProblem::SameNames || byNameDelete)
        return reject(L"keyless hand-authored alias scopes no longer use the exact path fallback");

    std::wstring parentPath;
    FileOperations::MutationInterlockScope byNameChild  = retainedPublishScopes.back();
    FileOperations::MutationInterlockScope byNameParent = byNameChild;
    if (! byNameChild.endpoint.pathIdentity.has_value() ||
        ! TryGetFileSystemParentPath(byNameChild.endpoint.pathIdentity.value(), byNameChild.providerPath, parentPath))
        return reject(L"could not derive the by-name parent control path");
    byNameChild.byNameCorroborated  = true;
    byNameParent.byNameCorroborated = true;
    byNameParent.providerPath       = std::move(parentPath);
    byNameParent.providerPathKey    = TryMakePathKey(byNameParent.endpoint.pathIdentity.value(), byNameParent.providerPath);
    // This synthetic by-name parent no longer has the copied leaf's retained-anchor chain. Its
    // classifier route proves overlap from the provider path, so discard stale authority suffixes.
    byNameParent.root.reset();
    byNameParent.anchors.clear();
    const auto byNameFast        = classify({byNameChild}, {byNameParent}, byNameDelete);
    const bool byNameFastDelete  = byNameDelete;
    byNameChild.providerPathKey  = std::nullopt;
    byNameParent.providerPathKey = std::nullopt;
    const auto byNameFallback    = classify({byNameChild}, {byNameParent}, byNameDelete);
    if (byNameFast != FileOperations::SameHostOverlapProblem::DeletedByNameWritten || ! byNameFastDelete || byNameFallback != byNameFast || ! byNameDelete)
        return reject(L"cached by-name ancestor classification changed from its keyless directional fallback");

    AppendLog(L"Same-host overlap classifier controls: supplementary siblings, mixed read/write, cross-root retained alias, by-name ancestor, and keyless "
              L"fallback passed.");
    return true;
}

[[nodiscard]] bool WriteLegacyMoveBreadcrumbFixture(const std::filesystem::path& path, const uint64_t taskId, const std::string_view phase) noexcept
{
    const std::string json = std::format(
        R"json({{"kind":"move-task-breadcrumb","schemaVersion":1,"taskId":{},"createdFileTime":1,"phase":"{}","sourcePane":"left","destinationPane":"right","nativeItems":0,"managedItems":2,"copyOnlyItems":0,"sourceRootCount":1,"sourceRootSamplesTruncated":false,"sourceRootSamples":[{{"pluginId":"builtin/file-system","pluginShortId":"file","instanceId":"host/default","profileId":"local","rootId":"fixture-root","representativePath":"C:\\fixture\\source.txt"}}],"destination":{{"pluginId":"builtin/file-system","pluginShortId":"file","instanceId":"host/default","profileId":"local","rootId":"fixture-root","representativePath":"C:\\fixture\\destination"}}}})json",
        taskId,
        phase);
    return SelfTest::WriteTextFile(path, json);
}

// This child deliberately remains suspended: the test needs an owned process-table
// entry with the same image basename, but must never run app startup or create a window.
// The canonical captured-child helper resumes its child, so it cannot supply this fixture.
[[nodiscard]] bool StartSuspendedLegacyBreadcrumbPeerForSelfTest(wil::unique_handle& jobOut, wil::unique_handle& processOut) noexcept
{
    std::array<wchar_t, 32768u> imagePath{};
    const DWORD imageLength = GetModuleFileNameW(nullptr, imagePath.data(), static_cast<DWORD>(imagePath.size()));
    if (imageLength == 0u || imageLength >= imagePath.size())
    {
        return false;
    }

    wil::unique_handle job(CreateJobObjectW(nullptr, nullptr));
    if (! job)
    {
        return false;
    }
    JOBOBJECT_EXTENDED_LIMIT_INFORMATION jobLimits{};
    jobLimits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
    if (SetInformationJobObject(job.get(), JobObjectExtendedLimitInformation, &jobLimits, sizeof(jobLimits)) == FALSE)
    {
        return false;
    }

    const std::wstring executable(imagePath.data(), imageLength);
    std::wstring commandLine = Common::Process::QuoteWindowsCommandLineArgument(executable);
    commandLine.append(L" --legacy-breadcrumb-suspended-peer");
    STARTUPINFOW startupInfo{};
    startupInfo.cb = sizeof(startupInfo);
    PROCESS_INFORMATION processInfo{};
    if (CreateProcessW(executable.c_str(),
                       commandLine.data(),
                       nullptr,
                       nullptr,
                       FALSE,
                       CREATE_SUSPENDED | CREATE_NO_WINDOW,
                       nullptr,
                       nullptr,
                       &startupInfo,
                       &processInfo) == FALSE)
    {
        return false;
    }

    wil::unique_handle process(processInfo.hProcess);
    wil::unique_handle thread(processInfo.hThread);
    if (AssignProcessToJobObject(job.get(), process.get()) == FALSE)
    {
        static_cast<void>(TerminateProcess(process.get(), 1u));
        static_cast<void>(WaitForSingleObject(process.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(5000u))));
        return false;
    }

    jobOut     = std::move(job);
    processOut = std::move(process);
    return true;
}

[[nodiscard]] bool RunFileOpsLegacyBreadcrumbPolicyChecks(const std::filesystem::path& store, FolderWindow& owner) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Legacy Move breadcrumb policy: {}", reason));
        return false;
    };
    using LegacyObservation = FileOperationMoveBreadcrumb::LegacyPeerObservationForSelfTest;
    FileOperationMoveBreadcrumb::SetLegacyPeerObservationForSelfTest(LegacyObservation::UseSystem);
    FileOperationMoveBreadcrumb::SetWriterObservationForSelfTest(FileOperationMoveBreadcrumb::WriterObservationForSelfTest::UseSystem);
    const auto resetProbes = wil::scope_exit([]() noexcept
    {
        FileOperationMoveBreadcrumb::SetLegacyPeerObservationForSelfTest(LegacyObservation::UseSystem);
        FileOperationMoveBreadcrumb::SetWriterObservationForSelfTest(FileOperationMoveBreadcrumb::WriterObservationForSelfTest::UseSystem);
    });

    const auto writeLegacy = [&](std::wstring_view name, uint64_t taskId, std::string_view phase) noexcept
    { return WriteLegacyMoveBreadcrumbFixture(store / name, taskId, phase); };
    std::vector<FileOperationMoveBreadcrumb::LoadedRecord> loaded;
    FileOperationMoveBreadcrumb::LoadStats stats{};
    const auto nonterminal = store / L"legacy-projected.move.json";
    FileOperationMoveBreadcrumb::SetLegacyPeerObservationForSelfTest(LegacyObservation::NoOtherProcess);
    if (! writeLegacy(L"legacy-projected.move.json", 0xB201u, "executing") ||
        FAILED(FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats)) || loaded.size() != 1u || loaded.front().record.schemaVersion != 1u ||
        loaded.front().record.writerProcessId != 0u || stats.legacyRecordsWithheld != 0u || stats.unverifiedWriterRecords != 0u)
    {
        return reject(L"a literal nonterminal schema-1 record was not projected under the no-peer test observation");
    }
    loaded.clear();
    {
        FolderWindow::FileOperationState legacyState(owner);
        std::vector<FolderWindow::FileOperationState::CompletedTaskSummary> summaries;
        legacyState.CollectCompletedTasks(summaries);
        if (summaries.size() != 1u || ! summaries.front().interruptedMoveNotice || summaries.front().selection.itemCount != 2u ||
            ! summaries.front().selection.singleName.empty())
        {
            return reject(L"the schema-1 record did not project as a read-only two-item notice with no fabricated subject");
        }
        legacyState.DismissCompletedTask(summaries.front().taskId);
        FileOperationMoveBreadcrumb::Record removed{};
        if (SUCCEEDED(FileOperationMoveBreadcrumb::Breadcrumb::Load(nonterminal, removed)))
        {
            return reject(L"explicit notice dismissal did not remove the exact safe schema-1 child");
        }
    }

    const auto terminal = store / L"legacy-terminal.move.json";
    std::error_code fixtureError;
    if (! writeLegacy(L"legacy-terminal.move.json", 0xB202u, "terminal") || FAILED(FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats)) ||
        ! loaded.empty() || stats.terminalRecordsRemoved != 1u || std::filesystem::exists(terminal, fixtureError) || fixtureError)
    {
        return reject(L"a terminal schema-1 record was not removed under the no-peer test observation");
    }

    const auto livePeerRecord = store / L"legacy-live-peer.move.json";
    wil::unique_handle peerJob;
    wil::unique_handle peerProcess;
    if (! writeLegacy(L"legacy-live-peer.move.json", 0xB203u, "admitted") || ! StartSuspendedLegacyBreadcrumbPeerForSelfTest(peerJob, peerProcess))
    {
        return reject(L"could not create the owned suspended same-name peer process");
    }
    const auto retirePeer = wil::scope_exit([&]() noexcept
    {
        peerJob.reset();
        if (peerProcess)
        {
            static_cast<void>(WaitForSingleObject(peerProcess.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(5000u))));
        }
    });
    FileOperationMoveBreadcrumb::SetLegacyPeerObservationForSelfTest(LegacyObservation::UseSystem);
    if (FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats) != S_FALSE || ! loaded.empty() || stats.legacyRecordsWithheld != 1u ||
        stats.unverifiedWriterRecords != 0u || FileOperationMoveBreadcrumb::Acknowledge(livePeerRecord) != HRESULT_FROM_WIN32(ERROR_BUSY) ||
        ! std::filesystem::exists(livePeerRecord, fixtureError) || fixtureError)
    {
        return reject(L"a live owned peer was not fail-closed through load and acknowledgement");
    }
    peerJob.reset();
    if (WaitForSingleObject(peerProcess.get(), static_cast<DWORD>(SelfTest::ScaleTimeout(5000u))) != WAIT_OBJECT_0)
    {
        return reject(L"the owned suspended same-name peer did not terminate when its job closed");
    }
    FileOperationMoveBreadcrumb::SetLegacyPeerObservationForSelfTest(LegacyObservation::Uncertain);
    if (FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats) != S_FALSE || ! loaded.empty() || stats.legacyRecordsWithheld != 1u ||
        stats.unverifiedWriterRecords != 0u || FileOperationMoveBreadcrumb::Acknowledge(livePeerRecord) != HRESULT_FROM_WIN32(ERROR_BUSY))
    {
        return reject(L"an uncertain peer enumeration was not retained separately from schema-2 writer uncertainty");
    }
    FileOperationMoveBreadcrumb::SetLegacyPeerObservationForSelfTest(LegacyObservation::NoOtherProcess);
    if (FileOperationMoveBreadcrumb::Acknowledge(livePeerRecord) != S_OK || std::filesystem::exists(livePeerRecord, fixtureError) || fixtureError)
    {
        return reject(L"the exact schema-1 fixture was not acknowledgeable after the no-peer control");
    }

    constexpr size_t kLegacyBatchSize = 16u;
    std::array<std::filesystem::path, kLegacyBatchSize> batchPaths{};
    const auto removeBatch = wil::scope_exit([&]() noexcept
    {
        FileOperationMoveBreadcrumb::SetLegacyPeerObservationForSelfTest(LegacyObservation::NoOtherProcess);
        for (const auto& path : batchPaths)
        {
            if (! path.empty())
                static_cast<void>(FileOperationMoveBreadcrumb::Acknowledge(path));
        }
    });
    for (size_t index = 0u; index < kLegacyBatchSize; ++index)
    {
        batchPaths[index] = store / std::format(L"legacy-batch-{:02}.move.json", index);
        if (! WriteLegacyMoveBreadcrumbFixture(batchPaths[index], 0xB300u + index, "executing"))
            return reject(L"the bounded legacy batch fixture could not be written");
    }
    if (FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats) != S_OK || loaded.size() != kLegacyBatchSize ||
        stats.legacyPeerObservations != 1u || stats.legacyRecordsWithheld != 0u)
        return reject(L"legacy batch projection did not share one post-enumeration eligibility observation");
    FileOperationMoveBreadcrumb::SetLegacyPeerObservationForSelfTest(LegacyObservation::Uncertain);
    if (FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats) != S_FALSE || ! loaded.empty() || stats.legacyPeerObservations != 1u ||
        stats.legacyRecordsWithheld != kLegacyBatchSize || stats.unverifiedWriterRecords != 0u)
        return reject(L"legacy batch withholding did not share one conservative eligibility observation");
    FileOperationMoveBreadcrumb::SetLegacyPeerObservationForSelfTest(LegacyObservation::NoOtherProcess);
    for (size_t index = 0u; index < kLegacyBatchSize; ++index)
    {
        if (! WriteLegacyMoveBreadcrumbFixture(batchPaths[index], 0xB300u + index, "terminal"))
            return reject(L"the terminal legacy batch fixture could not be written");
    }
    if (FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats) != S_OK || ! loaded.empty() || stats.legacyPeerObservations != 1u ||
        stats.terminalRecordsRemoved != kLegacyBatchSize)
        return reject(L"terminal legacy batch cleanup repeated the process observation or retained completed children");
    return true;
}

[[nodiscard]] bool RunFileOpsInterruptedMoveWriterChecks(SelfTestState& state) noexcept
{
    using State       = FolderWindow::FileOperationState;
    using Task        = State::Task;
    using Summary     = State::CompletedTaskSummary;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Interrupted Move writer: {}", reason));
        return false;
    };
    const auto root  = state.tempRoot / std::format(L"move-writer-{}", GetTickCount64());
    const auto store = root / L"breadcrumbs";
    if (! SelfTest::EnsureDirectory(store) || ! SelfTest::EnsureDirectory(root / L"logs"))
        return reject(L"the owned fixture store could not be created");
    FileOperationMoveBreadcrumb::SetRootForSelfTest(store);
    const auto restore                = wil::scope_exit([]() noexcept { FileOperationMoveBreadcrumb::ClearRootForSelfTest(); });
    const auto resetWriterObservation = wil::scope_exit([]() noexcept
    {
        FileOperationMoveBreadcrumb::SetWriterObservationForSelfTest(FileOperationMoveBreadcrumb::WriterObservationForSelfTest::UseSystem);
        FileOperationMoveBreadcrumb::SetLegacyPeerObservationForSelfTest(FileOperationMoveBreadcrumb::LegacyPeerObservationForSelfTest::UseSystem);
    });
    FileOperationMoveBreadcrumb::Record record{};
    record.taskId                          = 0xB001u;
    record.admittedStrategies.managedItems = 3u;
    record.sourceRootCount                 = 1u;
    record.sourceRootSamples               = {{.pluginId           = L"builtin/file-system",
                                               .pluginShortId      = L"file",
                                               .instanceId         = L"host/default",
                                               .profileId          = L"local",
                                               .rootId             = L"fixture-root",
                                               .representativePath = root / L"source.txt"}};
    record.destination                     = {.pluginId           = L"builtin/file-system",
                                              .pluginShortId      = L"file",
                                              .instanceId         = L"host/default",
                                              .profileId          = L"local",
                                              .rootId             = L"fixture-root",
                                              .representativePath = root / L"destination"};
    FileOperationMoveBreadcrumb::Breadcrumb breadcrumb;
    if (FAILED(FileOperationMoveBreadcrumb::Breadcrumb::Create(record, breadcrumb)))
        return reject(L"the current writer could not persist its admitted record");
    std::vector<FileOperationMoveBreadcrumb::LoadedRecord> loaded;
    FileOperationMoveBreadcrumb::LoadStats stats{};
    if (FAILED(FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats)) || ! loaded.empty() || stats.activeWriterRecords != 1u ||
        FileOperationMoveBreadcrumb::Acknowledge(breadcrumb.GetPath()) != HRESULT_FROM_WIN32(ERROR_BUSY))
        return reject(L"a live writer was misclassified or its record was acknowledged");
    FileOperationMoveBreadcrumb::SetWriterObservationForSelfTest(FileOperationMoveBreadcrumb::WriterObservationForSelfTest::Unverified);
    if (FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats) != S_FALSE || ! loaded.empty() || stats.unverifiedWriterRecords != 1u ||
        FileOperationMoveBreadcrumb::Acknowledge(breadcrumb.GetPath()) != HRESULT_FROM_WIN32(ERROR_BUSY))
        return reject(L"an unverifiable schema-2 writer did not remain withheld and unacknowledgeable");
    FileOperationMoveBreadcrumb::SetWriterObservationForSelfTest(FileOperationMoveBreadcrumb::WriterObservationForSelfTest::UseSystem);
    if (FAILED(breadcrumb.SimulateExitedWriterForSelfTest()) || FAILED(FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats)) ||
        loaded.size() != 1u || loaded.front().record.writerProcessId != GetCurrentProcessId())
        return reject(L"a recycled PID with a different creation time did not become an interrupted record");
    Common::Settings::Settings settings{};
    settings.fileOperations.emplace();
    auto firstOwner = std::make_unique<FolderWindow>();
    auto otherOwner = std::make_unique<FolderWindow>();
    firstOwner->SetSettings(&settings);
    otherOwner->SetSettings(&settings);
    {
        State operations(*firstOwner);
        operations.DebugConfigureDiagnosticsLogRouteForSelfTest(root / L"logs");
        State other(*otherOwner);
        std::vector<Summary> summaries;
        std::vector<Summary> duplicate;
        operations.CollectCompletedTasks(summaries);
        other.CollectCompletedTasks(duplicate);
        if (summaries.size() != 1u || ! duplicate.empty() || ! summaries.front().interruptedMoveNotice || summaries.front().selection.itemCount != 3u ||
            ! summaries.front().selection.singleName.empty())
            return reject(L"a notice was duplicated or counted provider roots instead of admitted items");
        const uint64_t noticeId = summaries.front().taskId;
        for (uint64_t index = 0u; index < 32u; ++index)
        {
            FileOperations::QualifiedEndpoint endpoint{.pluginId               = L"builtin/file-system",
                                                       .instanceId             = L"host/default",
                                                       .profileId              = L"local",
                                                       .rootId                 = L"fixture-root",
                                                       .pathIdentity           = FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(),
                                                       .cancellationRouteClass = FileOperations::CancellationRouteClass::Bounded};
            FileOperations::TransferPlan plan{};
            plan.intent                         = FileOperations::TransferIntent::Copy;
            plan.strategy                       = FileOperations::OperationStrategy::Copy;
            plan.sourceEndpoint                 = endpoint;
            plan.destinationEndpoint            = endpoint;
            plan.selectedItems                  = {{.providerPath = (root / L"item.bin").native()}};
            plan.destination.providerFolderPath = root.native();
            auto plans                          = std::make_shared<const FileOperations::FileOperationPlanGroup>(
                FileOperations::FileOperationPlanGroup{FileOperations::FileOperationPlan{std::move(plan)}});
            Task task(operations);
            task._taskId        = 0xB100u + index;
            task._operation     = FILESYSTEM_COPY;
            task._fileSystem    = state.fsLocal;
            task._executionMode = State::ExecutionMode::PerItem;
            task._sourcePaths   = {root / L"item.bin"};
            task.StorePlans(plans);
            task.InitializeAdmissionInputs(plans, {.fileSystem = state.fsLocal, .folder = root, .pluginId = endpoint.pluginId});
            task.InitializeSourceItemResultBuilders();
            static_cast<void>(operations.DebugRecordCompletedTaskForSelfTest(task));
        }
        summaries.clear();
        operations.CollectCompletedTasks(summaries);
        if (summaries.size() != 25u ||
            std::ranges::count_if(summaries, [noticeId](const auto& item) noexcept { return item.taskId == noticeId && item.interruptedMoveNotice; }) != 1)
            return reject(L"ordinary history eviction removed an undismissed notice or exceeded its own cap");
    }
    State successor(*firstOwner);
    std::vector<Summary> transferred;
    successor.CollectCompletedTasks(transferred);
    if (transferred.size() != 1u || ! transferred.front().interruptedMoveNotice)
        return reject(L"the notice could not transfer after its previous host and snapshots retired");
    successor.DismissCompletedTask(transferred.front().taskId);
    FileOperationMoveBreadcrumb::Record afterDismiss;
    if (SUCCEEDED(FileOperationMoveBreadcrumb::Breadcrumb::Load(breadcrumb.GetPath(), afterDismiss)))
        return reject(L"explicit dismissal did not acknowledge the exact inactive writer record");
    record.taskId                          = 0xB002u;
    record.admittedStrategies.managedItems = 1u;
    record.singleItemName                  = L"source.txt";
    FileOperationMoveBreadcrumb::Breadcrumb single;
    if (FAILED(FileOperationMoveBreadcrumb::Breadcrumb::Create(record, single)) || FAILED(single.SimulateExitedWriterForSelfTest()))
        return reject(L"the single-selection fixture could not persist");
    successor.DebugLoadInterruptedMoveBreadcrumbsForSelfTest();
    transferred.clear();
    successor.CollectCompletedTasks(transferred);
    if (transferred.size() != 1u || transferred.front().selection.itemCount != 1u || transferred.front().selection.singleName != L"source.txt")
        return reject(L"the single-selection subject was replaced by a full provider path");
    successor.DismissCompletedTask(transferred.front().taskId);
    if (! RunFileOpsLegacyBreadcrumbPolicyChecks(store, *otherOwner))
        return reject(L"the legacy schema-1 compatibility boundary failed");

    record.taskId         = 0xB003u;
    record.singleItemName = L"concurrent.txt";
    FileOperationMoveBreadcrumb::Breadcrumb concurrent;
    if (FAILED(FileOperationMoveBreadcrumb::Breadcrumb::Create(record, concurrent)) || FAILED(concurrent.SimulateExitedWriterForSelfTest()))
        return reject(L"the concurrent ID-allocation fixture could not persist");
    const auto removeConcurrent = wil::scope_exit([&]() noexcept { static_cast<void>(FileOperationMoveBreadcrumb::Acknowledge(concurrent.GetPath())); });
    const uint64_t before       = successor.AllocateLateQueueOrderKey();
    constexpr size_t kQueueIds  = 2'048u;
    std::array<uint64_t, kQueueIds> queueIds{};
    std::atomic<bool> start{false};
    // This public-path contention smoke complements the source lock-ownership guard;
    // it is not a deterministic reproduction of an unsynchronized lost update.
    std::jthread queueAllocator([&]() noexcept
    {
        start.wait(false, std::memory_order_acquire);
        for (uint64_t& id : queueIds)
        {
            id = successor.AllocateLateQueueOrderKey();
            std::this_thread::yield();
        }
    });
    start.store(true, std::memory_order_release);
    start.notify_all();
    successor.DebugLoadInterruptedMoveBreadcrumbsForSelfTest();
    queueAllocator.join();
    transferred.clear();
    successor.CollectCompletedTasks(transferred);
    const uint64_t after = successor.AllocateLateQueueOrderKey();
    if (transferred.size() != 1u || ! transferred.front().interruptedMoveNotice || transferred.front().interruptedMoveBreadcrumbPath != concurrent.GetPath() ||
        after != before + kQueueIds + 2u)
        return reject(L"concurrent notice and queue-key allocation did not preserve exact counter accounting");
    std::vector<uint64_t> allocated(queueIds.begin(), queueIds.end());
    allocated.push_back(transferred.front().taskId);
    std::ranges::sort(allocated);
    for (size_t index = 0u; index < allocated.size(); ++index)
    {
        if (allocated[index] != before + index + 1u)
            return reject(L"concurrent notice and queue-key allocation produced duplicate, missing or out-of-range IDs");
    }
    successor.DismissCompletedTask(transferred.front().taskId);
    FileOperationMoveBreadcrumb::Record dismissedConcurrent{};
    if (SUCCEEDED(FileOperationMoveBreadcrumb::Breadcrumb::Load(concurrent.GetPath(), dismissedConcurrent)))
        return reject(L"the concurrently allocated notice did not acknowledge its exact breadcrumb");
    return true;
}

[[nodiscard]] bool RunFileOpsDeleteParentPinChecks(SelfTestState& state) noexcept
{
    using Env         = RedSalamander::TestSupport::ScopedEnvironmentVariable;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Delete parent pin: {}", reason));
        return false;
    };
    if (! state.fsLocal || ! state.infoLocal)
        return reject(L"Local provider unavailable");
    std::string configuration;
    if (! BackupPluginConfiguration(state.infoLocal.get(), configuration))
        return reject(L"Local configuration backup failed");
    const auto restore = wil::scope_exit([&]() noexcept { static_cast<void>(SetPluginConfiguration(state.infoLocal.get(), configuration)); });
    const auto fixture = state.tempRoot / L"delete-parent-pins";
    if (! RecreateEmptyDirectory(fixture))
        return reject(L"owned fixture unavailable");
    const auto cleanup = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(fixture)); });
    wil::com_ptr<IFileSystemObjectBinding> binding;
    if (! state.fsLocal.try_query_to(binding.put()))
        return reject(L"Local binding unavailable");
    for (const unsigned int concurrency : {1u, 4u})
        for (const bool bound : {false, true})
            for (const bool cancel : {false, true})
            {
                if (! SetPluginConfiguration(state.infoLocal.get(), std::format("{{\"deleteMaxConcurrency\":{}}}", concurrency)))
                    return reject(L"delete concurrency configuration failed");
                const auto tree                 = fixture / L"tree";
                constexpr uint32_t nestedLevels = 32u;
                auto parent                     = tree / L"parent";
                for (uint32_t depth = 0u; depth < nestedLevels; ++depth)
                    parent /= L"d";
                const auto renamed  = fixture / L"renamed-parent";
                const auto sentinel = fixture / L"outside.bin";
                if (! WriteFilledTestFile(parent / L"a.bin", 17u, 0x31u) || ! WriteFilledTestFile(parent / L"b.bin", 29u, 0x92u) ||
                    ! WriteFilledTestFile(tree / L"sibling" / L"c.bin", 11u, 0x51u) || ! WriteFilledTestFile(sentinel, 31u, 0x77u))
                    return reject(L"owned tree staging failed");
                const std::wstring eventPrefix = std::format(L"Local\\RedSalamander.ParentPin.{}.{}", GetCurrentProcessId(), GetTickCount64());
                const std::wstring enteredName = eventPrefix + L".entered";
                const std::wstring releaseName = eventPrefix + L".release";
                wil::unique_handle entered(CreateEventW(nullptr, TRUE, FALSE, enteredName.c_str()));
                wil::unique_handle release(CreateEventW(nullptr, TRUE, FALSE, releaseName.c_str()));
                wil::unique_handle done(CreateEventW(nullptr, TRUE, FALSE, nullptr));
                if (! entered || ! release || ! done)
                    return reject(L"owned events unavailable");
                const std::wstring extendedParent  = Common::Paths::ToExtendedWin32Path(parent.native());
                const std::wstring extendedRenamed = Common::Paths::ToExtendedWin32Path(renamed.native());
                Env parentEnv(L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_PARENT", extendedParent);
                Env enteredEnv(L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_ENTERED_EVENT", enteredName);
                Env releaseEnv(L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_RELEASE_EVENT", releaseName);
                Env removedEnv(L"REDSALAMANDER_FILEOPS_DELETE_CHILD_PAUSE_REMOVED_PATH", std::nullopt);
                auto task          = std::make_unique<FolderWindow::FileOperationState::Task>(*state.fileOps);
                task->_operation   = FILESYSTEM_DELETE;
                task->_sourcePaths = {tree};
                task->BeginDiscovery();
                task->SkipDiscovery();
                FileSystemOptions options{};
                task->InitializeFileSystemOptions(options, nullptr);
                wil::com_ptr<IFileSystemBoundObject> retained;
                if (bound && FAILED(binding->BindObject(
                                 tree.c_str(),
                                 static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA | FILESYSTEM_BIND_DELETE),
                                 retained.put())))
                    return reject(L"selected-root binding failed");
                std::atomic<HRESULT> result{E_PENDING};
                const auto started = std::chrono::steady_clock::now();
                std::jthread worker([&]() noexcept
                {
                    const auto initialized = wil::CoInitializeEx(COINIT_MULTITHREADED);
                    static_cast<void>(initialized);
                    task->_dbgCallbackActiveScopeCount.fetch_add(1u, std::memory_order_relaxed);
                    const auto callback = wil::scope_exit([&]() noexcept { task->_dbgCallbackActiveScopeCount.fetch_sub(1u, std::memory_order_relaxed); });
                    FileSystemConditionalMutationResult receipt{};
                    receipt.sizeBytes = sizeof(receipt);
                    const HRESULT hr  = bound ? retained->DeleteIfUnchanged(FILESYSTEM_FLAG_RECURSIVE, &options, &receipt)
                                              : state.fsLocal->DeleteItem(tree.c_str(), FILESYSTEM_FLAG_RECURSIVE, &options, nullptr, nullptr);
                    result.store(hr, std::memory_order_release);
                    static_cast<void>(SetEvent(done.get()));
                });
                const auto ensureRelease = wil::scope_exit([&]() noexcept { static_cast<void>(SetEvent(release.get())); });
                if (WaitForSingleObject(entered.get(), 10'000u) != WAIT_OBJECT_0)
                    return reject(L"walk did not reach the post-listing committed-child pause");
                const bool renamedDuringWalk = MoveFileExW(extendedParent.c_str(), extendedRenamed.c_str(), 0u) != FALSE;
                const DWORD renameError      = renamedDuringWalk ? ERROR_SUCCESS : GetLastError();
                wil::unique_handle writer(CreateFileW(parent.c_str(),
                                                      GENERIC_WRITE,
                                                      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                                      nullptr,
                                                      OPEN_EXISTING,
                                                      FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS,
                                                      nullptr));
                if (renamedDuringWalk || (renameError != ERROR_SHARING_VIOLATION && renameError != ERROR_ACCESS_DENIED) || writer)
                    return reject(L"a traversed parent could be renamed or reopened for in-place write");
                if (cancel)
                    task->RequestCancel();
                static_cast<void>(SetEvent(release.get()));
                if (WaitForSingleObject(done.get(), 10'000u) != WAIT_OBJECT_0)
                    return reject(L"delete did not reach callback quiescence after release");
                worker.join();
                const HRESULT hr = result.load(std::memory_order_acquire);
                retained.reset();
                if ((! cancel && hr != S_OK) || (cancel && hr != HRESULT_FROM_WIN32(ERROR_CANCELLED)))
                    return reject(std::format(L"unexpected delete result 0x{:08X}", static_cast<unsigned long>(hr)));
                if (cancel)
                {
                    // Scanner/notification cleanup can lag callback quiescence briefly. Retry only
                    // sharing/access contention, bounded by the canonical pump, and retain the
                    // actual failure instead of diagnosing every rename failure as a leaked pin.
                    const auto renameAfterJoin = [&](const std::wstring& from, const std::wstring& to) noexcept
                    {
                        DWORD error           = ERROR_SUCCESS;
                        bool renamedAfterJoin = false;
                        const auto quiet      = RedSalamander::TestSupport::PumpMessagesUntil(
                            [&]() noexcept
                        {
                            renamedAfterJoin = MoveFileExW(from.c_str(), to.c_str(), 0u) != FALSE;
                            error            = renamedAfterJoin ? ERROR_SUCCESS : GetLastError();
                            return renamedAfterJoin || (error != ERROR_SHARING_VIOLATION && error != ERROR_ACCESS_DENIED);
                        },
                            {.timeout = SelfTest::Scale(std::chrono::seconds(2)), .operationName = L"released delete-parent pin"});
                        if (! quiet.conditionMet || ! renamedAfterJoin)
                            return reject(std::format(
                                L"post-join rename failed: bound={} configuredConcurrency={} error={} from=[{}] to=[{}]", bound, concurrency, error, from, to));
                        return true;
                    };
                    // A parallel child may have committed its directory removal before the
                    // root observes cancellation. Probe the fixture ancestor, which survives
                    // either partial outcome and cannot move while a descendant pin is live.
                    const std::wstring fixturePath  = Common::Paths::ToExtendedWin32Path(fixture.native());
                    const std::wstring releasedPath = Common::Paths::ToExtendedWin32Path((state.tempRoot / L"delete-parent-pins-released").native());
                    if (! renameAfterJoin(fixturePath, releasedPath) || ! renameAfterJoin(releasedPath, fixturePath))
                        return false;
                }
                else if (GetFileAttributesW(tree.c_str()) != INVALID_FILE_ATTRIBUTES)
                    return reject(L"ordinary recursive deletion retained its root");
                std::string sentinelBytes;
                if (! ReadFileTextFsIo(state.fsLocal.query<IFileSystemIO>(), sentinel, sentinelBytes) ||
                    sentinelBytes != std::string(31u, static_cast<char>(0x77)))
                    return reject(L"the outside-tree sentinel changed");
                Debug::Perf::Emit(
                    L"FileOps.SelfTest.DeleteParentPin",
                    std::format(L"bound={} configuredConcurrency={} effectiveConcurrency={} cancel={} nestedLevels={}",
                                bound,
                                concurrency,
                                bound ? 1u : concurrency,
                                cancel,
                                nestedLevels),
                    static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - started).count()),
                    bound ? 1u : concurrency,
                    nestedLevels,
                    hr);
                if (! SelfTest::RemoveAll(tree))
                    return reject(L"case storage cleanup failed");
            }
    return true;
}

[[nodiscard]] bool RunFileOpsCrossRootSkipContractChecks(SelfTestState& state) noexcept
{
    const auto unreachable = RunFileOpsCrossRootLocalCopyChecks(state, LoopbackShareStatus::Unreachable);
    const auto refused     = RunFileOpsCrossRootLocalCopyChecks(state, LoopbackShareStatus::Refused);
    return unreachable.status == SelfTest::SelfTestCaseResult::Status::skipped && ! unreachable.reason.empty() &&
           refused.status == SelfTest::SelfTestCaseResult::Status::failed && ! refused.reason.empty();
}
