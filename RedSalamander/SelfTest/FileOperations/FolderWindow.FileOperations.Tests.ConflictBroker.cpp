[[nodiscard]] bool RunFileOpsConflictBrokerChecks() noexcept
{
    using namespace FileOperations::ConflictBroker;
    using Action = Task::ConflictAction;
    Task::ConflictArbiter arbiter;
    arbiter.decisionEvent.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    if (! arbiter.decisionEvent)
        return false;
    const auto publish = [&](uint64_t generation)
    {
        std::scoped_lock lock(arbiter.mutex);
        ResetLocked(arbiter);
        auto& prompt               = arbiter.prompt;
        prompt.active              = true;
        prompt.generation          = generation;
        prompt.buttonsPublishable  = true;
        prompt.applyToAllEligible  = true;
        prompt.skipAllEligible     = true;
        prompt.actions[0]          = Action::Overwrite;
        prompt.actions[1]          = Action::Skip;
        prompt.actions[2]          = Action::Cancel;
        prompt.actions[3]          = Action::SkipAll;
        prompt.actions[4]          = Action::Retry;
        prompt.actionCount         = 5u;
        prompt.primaryActions[0]   = Action::Overwrite;
        prompt.primaryActions[1]   = Action::Cancel;
        prompt.primaryActionCount  = 2u;
        prompt.overflowActions[0]  = Action::Skip;
        prompt.overflowActions[1]  = Action::SkipAll;
        prompt.overflowActionCount = 2u;
        prompt.escapeAction        = Action::Cancel;
    };
    const auto accept = [&](uint64_t generation, Action action, Placement placement, bool repeat = false)
    {
        std::scoped_lock lock(arbiter.mutex);
        return TryAcceptLocked(arbiter, generation, action, repeat, placement);
    };
    publish(1u);
    if (! accept(1u, Action::Overwrite, Placement::Primary, true) || accept(1u, Action::Cancel, Placement::Escape) ||
        accept(1u, Action::Overwrite, Placement::Primary))
        return false;
    {
        std::scoped_lock lock(arbiter.mutex);
        const auto decision = PendingLocked(arbiter, 1u);
        if (! decision || decision.value().action != Action::Overwrite || ! decision.value().applyToAll)
            return false;
    }
    publish(2u);
    // Every route retained from A is harmless on B even when its action is offered.
    if (accept(1u, Action::Overwrite, Placement::Primary) || accept(1u, Action::Skip, Placement::Overflow) || accept(1u, Action::Cancel, Placement::Escape) ||
        accept(1u, Action::Overwrite, Placement::Published))
        return false;
    if (accept(2u, Action::Skip, Placement::Primary) || accept(2u, Action::Overwrite, Placement::Overflow) ||
        accept(2u, Action::Overwrite, Placement::Escape) || accept(0u, Action::Cancel, Placement::Published))
        return false;
    {
        std::scoped_lock lock(arbiter.mutex);
        // A stale/spurious wakeup has no decision authority.
        static_cast<void>(SetEvent(arbiter.decisionEvent.get()));
        if (PendingLocked(arbiter, 2u))
            return false;
        static_cast<void>(ResetEvent(arbiter.decisionEvent.get()));
        arbiter.prompt.metadataLoading = true;
    }
    if (accept(2u, Action::Cancel, Placement::Escape))
        return false;
    publish(3u);
    if (! accept(3u, Action::SkipAll, Placement::Overflow))
        return false;
    {
        std::scoped_lock lock(arbiter.mutex);
        const auto decision = PendingLocked(arbiter, 3u);
        if (! decision || decision.value().action != Action::Skip || ! decision.value().applyToAll)
            return false;
        ResetLocked(arbiter);
    }
    if (WaitForSingleObject(arbiter.decisionEvent.get(), 0u) != WAIT_TIMEOUT)
        return false;
    publish(4u);
    {
        std::scoped_lock lock(arbiter.mutex);
        arbiter.prompt.buttonsPublishable = false;
    }
    if (accept(4u, Action::Overwrite, Placement::Primary))
        return false;
    publish(4u);
    {
        std::scoped_lock lock(arbiter.mutex);
        arbiter.prompt.applyToAllChecked = true;
        if (! TryAcceptLocked(arbiter, 4u, Action::Overwrite, std::nullopt, Placement::Primary))
            return false;
        const auto decision = PendingLocked(arbiter, 4u);
        if (! decision || ! decision.value().applyToAll)
            return false;
    }
    publish(4u);
    {
        std::scoped_lock lock(arbiter.mutex);
        arbiter.prompt.applyToAllEligible = false;
        arbiter.prompt.skipAllEligible    = false;
        if (TryAcceptLocked(arbiter, 4u, Action::SkipAll, true, Placement::Overflow))
            return false;
        const auto decision = PendingLocked(arbiter, 4u);
        if (decision)
            return false;
    }
    publish(5u);
    {
        std::scoped_lock lock(arbiter.mutex);
        // DiskFull publishes an explicit scoped Skip All without a redundant checkbox.
        arbiter.prompt.applyToAllEligible  = false;
        arbiter.prompt.actions             = {{Action::Retry, Action::SkipAll, Action::Cancel}};
        arbiter.prompt.actionCount         = 3u;
        arbiter.prompt.primaryActions      = arbiter.prompt.actions;
        arbiter.prompt.primaryActionCount  = 3u;
        arbiter.prompt.overflowActionCount = 0u;
        if (TryAcceptLocked(arbiter, 4u, Action::SkipAll, std::nullopt, Placement::Primary) ||
            TryAcceptLocked(arbiter, 5u, Action::Skip, std::nullopt, Placement::Primary) ||
            ! TryAcceptLocked(arbiter, 5u, Action::SkipAll, std::nullopt, Placement::Primary))
            return false;
        const auto decision = PendingLocked(arbiter, 5u);
        if (! decision || decision.value().action != Action::Skip || ! decision.value().applyToAll)
            return false;
    }
    struct RepeatScenario
    {
        Action action;
        std::optional<bool> applyToAll;
        bool checked;
        bool eligible;
        bool expected;
    };
    constexpr std::array<RepeatScenario, 9> repeatScenarios{{{Action::Skip, true, false, true, true},
                                                             {Action::Skip, std::nullopt, true, true, true},
                                                             {Action::Skip, std::nullopt, false, true, false},
                                                             {Action::Skip, false, true, true, false},
                                                             {Action::Skip, true, true, false, false},
                                                             {Action::Retry, std::nullopt, true, true, false},
                                                             {Action::Retry, true, true, true, false},
                                                             {Action::Cancel, std::nullopt, true, true, false},
                                                             {Action::Cancel, true, true, true, false}}};
    uint64_t generation = 6u;
    for (const auto& scenario : repeatScenarios)
    {
        publish(generation);
        std::scoped_lock lock(arbiter.mutex);
        arbiter.prompt.applyToAllChecked  = scenario.checked;
        arbiter.prompt.applyToAllEligible = scenario.eligible;
        if (! TryAcceptLocked(arbiter, generation, scenario.action, scenario.applyToAll, Placement::Published))
            return false;
        const auto decision = PendingLocked(arbiter, generation);
        if (! decision || decision.value().action != scenario.action || decision.value().applyToAll != scenario.expected)
            return false;
        ++generation;
    }
    publish(generation);
    {
        std::scoped_lock lock(arbiter.mutex);
        if (TryAcceptLocked(arbiter, generation, Action::DeleteReadOnly, true, Placement::Published))
            return false;
        arbiter.prompt.actions[0]        = Action::DeleteReadOnly;
        arbiter.prompt.primaryActions[0] = Action::DeleteReadOnly;
        arbiter.prompt.applyToAllChecked = true;
        if (! TryAcceptLocked(arbiter, generation, Action::DeleteReadOnly, std::nullopt, Placement::Primary))
            return false;
        const auto decision = PendingLocked(arbiter, generation);
        if (! decision || decision.value().action != Action::DeleteReadOnly || decision.value().applyToAll)
            return false;
    }
    ++generation;
    // Serialize competing submissions at the production mutex; only one wins.
    publish(generation);
    std::atomic<unsigned int> accepted{0u};
    std::jthread first([&]
    {
        if (accept(generation, Action::Overwrite, Placement::Primary))
            accepted.fetch_add(1u);
    });
    std::jthread second([&]
    {
        if (accept(generation, Action::Cancel, Placement::Escape))
            accepted.fetch_add(1u);
    });
    first.join();
    second.join();
    return accepted.load() == 1u;
}

// Two real selected roots exercise the bridge constructor's namespace-root wiring. A
// second metadata risk remains independent, while a fresh unchecked task never reuses consent.
[[nodiscard]] bool RunFileOpsMetadataConsentRootScopeChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps metadata consent roots: FAILED: {}", reason));
        return false;
    };
    auto* operations = TryGetFileOps(TryGetFolderWindow(state.mainWindow));
    if (! operations || ! state.fsLocal || state.tempRoot.empty() || operations->HasActiveOperations())
        return reject(L"requires an idle owner, Local provider and owned test root");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const auto previousSettings                               = g_settings.fileOperations;
    const auto restoreSettings                                = wil::scope_exit([&]() noexcept { g_settings.fileOperations = previousSettings; });
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    constexpr std::wstring_view forceName                     = L"REDSALAMANDER_FILEOPS_METADATA_FORCE_PRESENT_MASK";
    constexpr std::wstring_view failName                      = L"REDSALAMANDER_FILEOPS_METADATA_FAIL_MASK";
    const auto previousForce                                  = RedSalamander::TestSupport::ReadEnvironmentValue(forceName);
    const auto previousFail                                   = RedSalamander::TestSupport::ReadEnvironmentValue(failName);
    if (previousForce.error != ERROR_SUCCESS || previousFail.error != ERROR_SUCCESS)
        return reject(L"metadata injection state could not be preserved");
    const auto restoreEnvironment = wil::scope_exit([&]() noexcept
    {
        static_cast<void>(SetEnvironmentVariableW(forceName.data(), previousForce.value.has_value() ? previousForce.value.value().c_str() : nullptr));
        static_cast<void>(SetEnvironmentVariableW(failName.data(), previousFail.value.has_value() ? previousFail.value.value().c_str() : nullptr));
    });
    const std::wstring forceMask  = std::format(L"{}", static_cast<uint32_t>(FILESYSTEM_METADATA_PLACEHOLDER | FILESYSTEM_METADATA_SPARSE));
    const std::wstring failMask   = std::format(L"{}", static_cast<uint32_t>(FILESYSTEM_METADATA_SPARSE));
    if (! SetEnvironmentVariableW(forceName.data(), forceMask.c_str()) || ! SetEnvironmentVariableW(failName.data(), failMask.c_str()))
        return reject(L"metadata risk injection could not be enabled");
    const auto root    = state.tempRoot / L"metadata-consent-selected-roots";
    bool workerQuiet   = true;
    const auto cleanup = wil::scope_exit([&]() noexcept
    {
        if (workerQuiet)
            static_cast<void>(SelfTest::RemoveAll(root));
        else
            state.retainTempRootOnFailure = true;
    });
    if (! RecreateEmptyDirectory(root))
        return reject(L"fixture root could not be created");
    for (const bool checked : {true, false})
    {
        const auto scenario = root / (checked ? L"checked" : L"unchecked");
        const std::array<std::filesystem::path, 2u> sources{{scenario / L"source" / L"first.bin", scenario / L"source" / L"second.bin"}};
        const auto destination = scenario / L"destination";
        if (! RecreateEmptyDirectory(destination) || ! WriteFilledTestFile(sources[0], 16u * 1024u, 0x31u) ||
            ! WriteFilledTestFile(sources[1], 32u * 1024u, 0x73u))
            return reject(L"the two source roots could not be prepared");
        const auto taskId = StartFileOperationAndGetId(operations,
                                                       FILESYSTEM_COPY,
                                                       FolderWindow::Pane::Left,
                                                       FolderWindow::Pane::Right,
                                                       state.fsLocal,
                                                       {sources[0], sources[1]},
                                                       destination,
                                                       FILESYSTEM_FLAG_NONE,
                                                       false,
                                                       0u,
                                                       FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                       false,
                                                       state.fsLocal);
        if (! taskId.has_value())
            return reject(L"the actual two-root bridge Copy was not admitted");
        workerQuiet                 = false;
        const auto retire           = wil::scope_exit([&]() noexcept
        {
            workerQuiet = CancelOutstandingSelfTestOperations(operations);
            if (! workerQuiet)
                operations->Shutdown();
            operations->DismissCompletedTask(taskId.value());
            operations->DebugRemoveDiagnosticsForTask(taskId.value());
            state.completedTasks.erase(taskId.value());
        });
        uint64_t answeredGeneration = 0u;
        uint64_t prompts            = 0u;
        bool refused                = false;
        std::wstring expectedRoot;
        const auto completed = RedSalamander::TestSupport::PumpMessagesUntil(
            [&]() noexcept
        {
            auto* task        = operations->FindTask(taskId.value());
            const auto prompt = TryGetConflictPromptCopy(task);
            if (task && prompt.has_value() && prompt->generation != answeredGeneration)
            {
                const auto plans = task->LoadPlans();
                bool qualified   = plans != nullptr;
                if (plans)
                    for (const auto& plan : *plans)
                        if (const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan))
                        {
                            if (expectedRoot.empty())
                                expectedRoot = transfer->destinationEndpoint.rootId;
                            qualified = qualified && ! expectedRoot.empty() && transfer->destinationEndpoint.rootId == expectedRoot;
                        }
                if (! qualified || ! prompt->deferredConsent || prompt->decisionScope.destinationRootId != expectedRoot ||
                    (prompt->bucket != Task::ConflictBucket::PlaceholderHydration && prompt->bucket != Task::ConflictBucket::SparseInflation) ||
                    ! prompt->applyToAllEligible || ! PromptHasAction(prompt.value(), Task::ConflictAction::Proceed))
                {
                    refused = true;
                    task->RequestCancel();
                }
                else
                {
                    if (checked)
                        task->ToggleConflictApplyToAllChecked(prompt->generation);
                    answeredGeneration = prompt->generation;
                    ++prompts;
                    if (! task->SubmitConflictDecision(prompt->generation, Task::ConflictAction::Proceed, std::nullopt))
                    {
                        refused = true;
                        task->RequestCancel();
                    }
                }
            }
            return state.completedTasks.contains(taskId.value()) && ! operations->HasActiveOperations();
        },
            {.timeout = SelfTest::Scale(std::chrono::seconds(20)), .operationName = L"two selected metadata roots complete"});
        const auto result = state.completedTasks.find(taskId.value());
        if (! completed.conditionMet || result == state.completedTasks.end() || FAILED(result->second.hr) || refused || prompts != (checked ? 2u : 4u) ||
            ! FilesEqualBytes(sources[0], destination / sources[0].filename()) || ! FilesEqualBytes(sources[1], destination / sources[1].filename()))
            return reject(L"checked scope, independent risk, fresh unchecked task or exact Copy payload failed");
        operations->DebugFlushDiagnosticsForSelfTest();
        std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> diagnostics;
        operations->CollectDiagnostics(diagnostics);
        const auto cached = std::ranges::count_if(diagnostics,
                                                  [&](const auto& entry) noexcept
        {
            return entry.taskId == taskId.value() && entry.acceptedDecision.has_value() &&
                   entry.acceptedDecision->origin == FileOperations::ConflictDecisionOrigin::CachedScope &&
                   entry.acceptedDecision->scope.destinationRootId == expectedRoot;
        });
        if (cached != (checked ? 2 : 0))
            return reject(L"the second selected root did not retain the expected cached-origin records");
        Debug::Perf::Emit(L"FileOps.SelfTest.MetadataConsentSelectedRoots",
                          checked ? L"checked; two risks" : L"unchecked; two risks",
                          0u,
                          prompts,
                          static_cast<uint64_t>(cached),
                          S_OK);
        AppendLog(
            std::format(L"FileOps metadata consent roots: checked={} selectedRoots=2 risks=2 prompts={} cached={} exactBytes=true", checked, prompts, cached));
    }
    return true;
}

[[nodiscard]] bool RunFileOpsKeepBothNestedRepeatChecks(SelfTestState& state) noexcept
{
    using Task        = FolderWindow::FileOperationState::Task;
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"FileOps nested Keep Both repeat: FAILED: {}", reason));
        return false;
    };
    auto* operations = TryGetFileOps(TryGetFolderWindow(state.mainWindow));
    if (! operations || ! state.fsLocal || state.tempRoot.empty() || operations->HasActiveOperations())
        return reject(L"requires an idle owner, Local provider and owned test root");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");
    const auto previousSettings                               = g_settings.fileOperations;
    const auto restoreSettings                                = wil::scope_exit([&]() noexcept { g_settings.fileOperations = previousSettings; });
    EnsureFileOperationsSettingsForSelfTest().verifyAfterCopy = false;
    const auto root                                           = state.tempRoot / L"keep-both-descendant-repeat";
    bool workerQuiet                                          = true;
    const auto cleanup                                        = wil::scope_exit([&]() noexcept
    {
        if (workerQuiet)
            static_cast<void>(SelfTest::RemoveAll(root));
        else
            state.retainTempRootOnFailure = true;
    });
    if (! RecreateEmptyDirectory(root))
        return reject(L"fixture root unavailable");
    const auto source      = root / L"source" / L"tree";
    const auto destination = root / L"destination";
    const std::array<std::wstring_view, 2u> parents{{L"first", L"second"}};
    for (size_t index = 0u; index < parents.size(); ++index)
    {
        if (! WriteFilledTestFile(source / parents[index] / L"leaf.bin", (index + 1u) * 16u * 1024u, static_cast<unsigned char>(0x31u + index)) ||
            ! WriteFilledTestFile(destination / L"tree" / parents[index] / L"leaf.bin", 8u * 1024u, static_cast<unsigned char>(0xA0u + index)) ||
            ! WriteFilledTestFile(root / std::format(L"original-{}.bin", index), 8u * 1024u, static_cast<unsigned char>(0xA0u + index)))
            return reject(L"the two distinct child collisions could not be prepared");
    }
    const auto taskId = StartFileOperationAndGetId(operations,
                                                   FILESYSTEM_COPY,
                                                   FolderWindow::Pane::Left,
                                                   FolderWindow::Pane::Right,
                                                   state.fsLocal,
                                                   {source},
                                                   destination,
                                                   FILESYSTEM_FLAG_RECURSIVE,
                                                   false,
                                                   0u,
                                                   FolderWindow::FileOperationState::ExecutionMode::PerItem,
                                                   false);
    if (! taskId.has_value())
        return reject(L"recursive Local Copy was not admitted");
    workerQuiet                 = false;
    const auto retire           = wil::scope_exit([&]() noexcept
    {
        workerQuiet = CancelOutstandingSelfTestOperations(operations);
        if (! workerQuiet)
            operations->Shutdown();
        operations->DismissCompletedTask(taskId.value());
        operations->DebugRemoveDiagnosticsForTask(taskId.value());
        state.completedTasks.erase(taskId.value());
    });
    uint64_t answeredGeneration = 0u;
    uint64_t prompts            = 0u;
    bool refused                = false;
    const auto completed        = RedSalamander::TestSupport::PumpMessagesUntil(
        [&]() noexcept
    {
        auto* task        = operations->FindTask(taskId.value());
        const auto prompt = TryGetConflictPromptCopy(task);
        if (task && prompt.has_value() && prompt->generation != answeredGeneration)
        {
            if (prompts != 0u || prompt->bucket != Task::ConflictBucket::RegularFileExists || ! prompt->applyToAllEligible ||
                prompt->decisionScope.keepBothContext != FileOperations::ConflictRepeatContext::Descendant ||
                ! PromptHasAction(prompt.value(), Task::ConflictAction::KeepBoth))
            {
                refused = true;
                task->RequestCancel();
            }
            else
            {
                answeredGeneration = prompt->generation;
                ++prompts;
                task->ToggleConflictApplyToAllChecked(prompt->generation);
                if (! task->SubmitConflictDecision(prompt->generation, Task::ConflictAction::KeepBoth, std::nullopt))
                {
                    refused = true;
                    task->RequestCancel();
                }
            }
        }
        return state.completedTasks.contains(taskId.value()) && ! operations->HasActiveOperations();
    },
        {.timeout = SelfTest::Scale(std::chrono::seconds(20)), .operationName = L"nested Keep Both scoped Copy completion"});
    const auto result = state.completedTasks.find(taskId.value());
    if (! completed.conditionMet || result == state.completedTasks.end() || FAILED(result->second.hr) || refused || prompts != 1u)
        return reject(L"the checked descendant choice did not cover both child collisions");
    for (size_t index = 0u; index < parents.size(); ++index)
        if (! FilesEqualBytes(source / parents[index] / L"leaf.bin", destination / L"tree" / parents[index] / L"leaf (2).bin") ||
            ! FilesEqualBytes(root / std::format(L"original-{}.bin", index), destination / L"tree" / parents[index] / L"leaf.bin"))
            return reject(L"a child sibling was misplaced or original/source bytes changed");
    operations->DebugFlushDiagnosticsForSelfTest();
    std::vector<FolderWindow::FileOperationState::TaskDiagnosticEntry> diagnostics;
    operations->CollectDiagnostics(diagnostics);
    const auto cached = std::ranges::count_if(diagnostics,
                                              [&](const auto& entry) noexcept
    {
        return entry.taskId == taskId.value() && entry.acceptedDecision.has_value() &&
               entry.acceptedDecision->origin == FileOperations::ConflictDecisionOrigin::CachedScope &&
               entry.acceptedDecision->action == Task::ConflictAction::KeepBoth &&
               entry.acceptedDecision->scope.keepBothContext == FileOperations::ConflictRepeatContext::Descendant;
    });
    if (cached != 1)
        return reject(L"the matching descendant did not retain one scoped cached decision");
    Debug::Perf::Emit(L"FileOps.SelfTest.KeepBothDescendantRepeat",
                      L"two nested parents; originals preserved; siblings exact",
                      0u,
                      prompts,
                      static_cast<uint64_t>(cached),
                      S_OK);
    AppendLog(L"FileOps nested Keep Both repeat: prompts=1 cached=1 parents=2 exactOriginalsAndSiblings=true");
    return true;
}
