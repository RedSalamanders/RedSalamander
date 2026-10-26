// I35 S1 (Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-11.md, section 5):
// questions, decisions and expected outcomes are Information notes, never counted warnings; the
// informational-notes setting changes only what is shown; metadata outcomes classify alike on the
// host bridge and the provider-native Copy routes.

namespace
{
using NotesTask           = FolderWindow::FileOperationState::Task;
using NotesSeverity       = FolderWindow::FileOperationState::DiagnosticSeverity;
using NotesSummary        = FolderWindow::FileOperationState::CompletedTaskSummary;
using NotesDiagnostic     = FolderWindow::FileOperationState::TaskDiagnosticEntry;
constexpr auto kNotesWait = std::chrono::seconds(30);

[[nodiscard]] std::optional<NotesSummary> FindNotesCompletedSummary(FolderWindow::FileOperationState* operations, uint64_t taskId)
{
    std::vector<NotesSummary> summaries;
    operations->CollectCompletedTasks(summaries);
    const auto found = std::ranges::find_if(summaries, [taskId](const NotesSummary& summary) noexcept { return summary.taskId == taskId; });
    if (found == summaries.end())
        return std::nullopt;
    return *found;
}

[[nodiscard]] size_t CountNotesDiagnostics(const NotesSummary& summary, std::wstring_view category, NotesSeverity severity) noexcept
{
    return static_cast<size_t>(std::ranges::count_if(
        summary.issueDiagnostics, [&](const NotesDiagnostic& issue) noexcept { return issue.category == category && issue.severity == severity; }));
}

[[nodiscard]] bool HasNotesProblemDiagnostic(const NotesSummary& summary, std::wstring_view prefix) noexcept
{
    return std::ranges::any_of(summary.issueDiagnostics, [&](const NotesDiagnostic& issue) noexcept {
        return issue.category.starts_with(prefix) && (issue.severity == NotesSeverity::Warning || issue.severity == NotesSeverity::Error);
    });
}

[[nodiscard]] std::wstring DescribeNotesDiagnostics(const NotesSummary& summary)
{
    std::wstring text;
    for (const NotesDiagnostic& issue : summary.issueDiagnostics)
    {
        if (! text.empty())
            text += L"; ";
        text += std::format(L"{}:{}:0x{:08X}", issue.category, static_cast<unsigned int>(issue.severity), static_cast<unsigned long>(issue.status));
    }
    return text;
}

[[nodiscard]] bool WriteNotesStream(const std::filesystem::path& file, std::wstring_view stream, std::string_view payload) noexcept
{
    const std::wstring target = MakeExtendedPathForSelfTest(file) + L":" + std::wstring(stream);
    wil::unique_hfile handle(
        CreateFileW(target.c_str(), GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_DELETE, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    DWORD written = 0u;
    return handle && WriteFile(handle.get(), payload.data(), static_cast<DWORD>(payload.size()), &written, nullptr) != FALSE &&
           written == static_cast<DWORD>(payload.size());
}

[[nodiscard]] std::optional<std::string> ReadNotesStream(const std::filesystem::path& file, std::wstring_view stream)
{
    const std::wstring target = MakeExtendedPathForSelfTest(file) + L":" + std::wstring(stream);
    wil::unique_hfile handle(CreateFileW(
        target.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! handle)
        return std::nullopt;
    std::string payload(4096u, '\0');
    DWORD read = 0u;
    if (ReadFile(handle.get(), payload.data(), static_cast<DWORD>(payload.size()), &read, nullptr) == FALSE)
        return std::nullopt;
    payload.resize(read);
    return payload;
}

// Grants a test-only principal (Guests) read access that every new child inherits, so a copied
// file's DACL differs from its source's only by destination inheritance.
[[nodiscard]] bool AddInheritableGuestsReadForNotes(const std::filesystem::path& directory) noexcept
{
    std::array<std::byte, SECURITY_MAX_SID_SIZE> sid{};
    DWORD sidBytes = static_cast<DWORD>(sid.size());
    if (! CreateWellKnownSid(WinBuiltinGuestsSid, nullptr, sid.data(), &sidBytes))
        return false;
    PACL existing                 = nullptr;
    PSECURITY_DESCRIPTOR rawOwner = nullptr;
    if (GetNamedSecurityInfoW(directory.c_str(), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr, nullptr, &existing, nullptr, &rawOwner) != ERROR_SUCCESS ||
        rawOwner == nullptr)
        return false;
    wil::unique_hlocal_ptr<void> descriptor(rawOwner);
    EXPLICIT_ACCESSW grant{};
    grant.grfAccessPermissions = FILE_GENERIC_READ;
    grant.grfAccessMode        = GRANT_ACCESS;
    grant.grfInheritance       = SUB_CONTAINERS_AND_OBJECTS_INHERIT;
    grant.Trustee.TrusteeForm  = TRUSTEE_IS_SID;
    grant.Trustee.TrusteeType  = TRUSTEE_IS_WELL_KNOWN_GROUP;
    grant.Trustee.ptstrName    = reinterpret_cast<wchar_t*>(sid.data());
    PACL merged                = nullptr;
    if (SetEntriesInAclW(1u, &grant, existing, &merged) != ERROR_SUCCESS || merged == nullptr)
        return false;
    wil::unique_hlocal_ptr<ACL> ownedMerged(merged);
    return SetNamedSecurityInfoW(const_cast<wchar_t*>(directory.c_str()), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nullptr, nullptr, merged, nullptr) ==
           ERROR_SUCCESS;
}

// Replaces a source file's DACL with a protected one (Everyone: full control, nothing inherited),
// the shape a user creates with "Disable inheritance".
[[nodiscard]] bool ProtectDaclForNotes(const std::filesystem::path& file) noexcept
{
    std::array<std::byte, SECURITY_MAX_SID_SIZE> sid{};
    DWORD sidBytes = static_cast<DWORD>(sid.size());
    if (! CreateWellKnownSid(WinWorldSid, nullptr, sid.data(), &sidBytes))
        return false;
    EXPLICIT_ACCESSW grant{};
    grant.grfAccessPermissions = FILE_ALL_ACCESS;
    grant.grfAccessMode        = SET_ACCESS;
    grant.grfInheritance       = NO_INHERITANCE;
    grant.Trustee.TrusteeForm  = TRUSTEE_IS_SID;
    grant.Trustee.TrusteeType  = TRUSTEE_IS_WELL_KNOWN_GROUP;
    grant.Trustee.ptstrName    = reinterpret_cast<wchar_t*>(sid.data());
    PACL dacl                  = nullptr;
    if (SetEntriesInAclW(1u, &grant, nullptr, &dacl) != ERROR_SUCCESS || dacl == nullptr)
        return false;
    wil::unique_hlocal_ptr<ACL> ownedDacl(dacl);
    return SetNamedSecurityInfoW(const_cast<wchar_t*>(file.c_str()),
                                 SE_FILE_OBJECT,
                                 DACL_SECURITY_INFORMATION | PROTECTED_DACL_SECURITY_INFORMATION,
                                 nullptr,
                                 nullptr,
                                 dacl,
                                 nullptr) == ERROR_SUCCESS;
}

// Owns the settings, prompt and popup state a Notes case changes, and every task it starts.
class NotesCaseScope final
{
public:
    NotesCaseScope(SelfTestState& state, FolderWindow* folder, FolderWindow::FileOperationState* operations) noexcept
        : _state(state),
          _folder(folder),
          _operations(operations)
    {
        auto& settings = EnsureFileOperationsSettingsForSelfTest();
        _showNotes     = settings.showInformationalNotes;
        _infoEnabled   = settings.diagnosticsInfoEnabled;
        _verify        = settings.verifyAfterCopy;
        _autoDismiss   = operations->GetAutoDismissSuccess();
        _originalPopup = operations->GetPopupHwndForSelfTest();
        operations->SetAutoDismissSuccess(false);
    }
    NotesCaseScope(const NotesCaseScope&)            = delete;
    NotesCaseScope& operator=(const NotesCaseScope&) = delete;
    ~NotesCaseScope()
    {
        const bool quiet = CancelOutstandingSelfTestOperations(_operations);
        for (const uint64_t taskId : _tasks)
        {
            if (quiet)
            {
                _operations->DismissCompletedTask(taskId);
                _operations->DebugRemoveDiagnosticsForTask(taskId);
                _state.completedTasks.erase(taskId);
            }
        }
        auto& settings                  = EnsureFileOperationsSettingsForSelfTest();
        settings.showInformationalNotes = _showNotes;
        settings.diagnosticsInfoEnabled = _infoEnabled;
        settings.verifyAfterCopy        = _verify;
        _operations->SetAutoDismissSuccess(_autoDismiss);
        static_cast<void>(_operations->DebugFlushPendingSettingsSaveForSelfTest(5'000u));
        if (quiet && ! _originalPopup && _operations->GetPopupHwndForSelfTest())
            _operations->DebugResetPopupForSelfTest();
    }

    // Starts a Copy, answers its one conflict prompt when asked, and waits for the retired task.
    [[nodiscard]] uint64_t RunCopy(std::vector<std::filesystem::path> sources,
                                   const std::filesystem::path& destination,
                                   std::optional<NotesTask::ConflictAction> answer = std::nullopt,
                                   FileSystemFlags flags                           = FILESYSTEM_FLAG_RECURSIVE)
    {
        const auto admitted = StartFileOperationAndGetId(_operations,
                                                         FILESYSTEM_COPY,
                                                         FolderWindow::Pane::Left,
                                                         FolderWindow::Pane::Right,
                                                         _state.fsLocal,
                                                         std::move(sources),
                                                         destination,
                                                         flags,
                                                         false,
                                                         0u,
                                                         FolderWindow::FileOperationState::ExecutionMode::PerItem);
        if (! admitted.has_value())
            return 0u;
        const uint64_t taskId = admitted.value();
        _tasks.push_back(taskId);
        if (answer.has_value())
        {
            std::optional<NotesTask::ConflictPromptState> prompt;
            const auto asked = RedSalamander::TestSupport::PumpMessagesUntil(
                [&]() noexcept
            {
                prompt = TryGetConflictPromptCopy(_operations->FindTask(taskId));
                return prompt.has_value() || _state.completedTasks.contains(taskId);
            },
                {.timeout = SelfTest::Scale(kNotesWait), .operationName = L"Notes conflict prompt"});
            NotesTask* task = _operations->FindTask(taskId);
            if (! asked.conditionMet || ! prompt.has_value() || task == nullptr || ! task->SubmitConflictDecision(prompt->generation, answer.value(), false))
                return 0u;
        }
        const auto done = RedSalamander::TestSupport::PumpMessagesUntil([&]() noexcept {
            return _state.completedTasks.contains(taskId) && _operations->FindTask(taskId) == nullptr;
        }, {.timeout = SelfTest::Scale(kNotesWait), .operationName = L"Notes task completion"});
        return done.conditionMet ? taskId : 0u;
    }

private:
    SelfTestState& _state;
    FolderWindow* _folder                         = nullptr;
    FolderWindow::FileOperationState* _operations = nullptr;
    std::vector<uint64_t> _tasks;
    HWND _originalPopup = nullptr;
    bool _showNotes     = false;
    bool _infoEnabled   = true;
    bool _verify        = false;
    bool _autoDismiss   = false;
};

// Whether the actual Issues pane projects any row for the task under the current setting.
[[nodiscard]] bool NotesIssuesPaneShowsTask(SelfTestState& state, FolderWindow* folder, FolderWindow::FileOperationState* operations, uint64_t taskId) noexcept
{
    auto lifetime = std::make_shared<int>(1);
    wil::unique_hwnd issues(FileOperationsIssuesPane::Create(operations, folder, state.mainWindow, lifetime));
    const auto close = wil::scope_exit([&]() noexcept
    {
        lifetime.reset();
        issues.reset();
    });
    return issues && FileOperationsIssuesPane::SelfTestSelectTask(issues.get(), taskId);
}

[[nodiscard]] bool NotesProjectableTask(FolderWindow::FileOperationState* operations, uint64_t taskId) noexcept
{
    std::vector<uint64_t> projectable;
    operations->CollectProjectableIssueTaskIds(projectable);
    return std::ranges::find(projectable, taskId) != projectable.end();
}
} // namespace

// R5.1-R5.3, R5.6: a question answered without a failure, a skip the user chose and a log-verbosity
// setting never become counted warnings; a real failure still does.
[[nodiscard]] bool RunFileOpsNotesSeverityChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Notes severity: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"notes-severity";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    NotesCaseScope scope(state, folder, operations);
    auto& settings = EnsureFileOperationsSettingsForSelfTest();

    // 1. Copy onto an existing file, answer Replace: the question is a note, the task is Done and
    //    auto-dismissable, and the question row is hidden unless the setting shows notes.
    const auto replaceSource      = root / L"replace-src" / L"payload.bin";
    const auto replaceDestination = root / L"replace-dst";
    if (! WriteTestFile(replaceSource, 4096u) || ! RecreateEmptyDirectory(replaceDestination) || ! WriteTestFile(replaceDestination / L"payload.bin", 512u))
        return reject(L"Replace fixture unavailable");
    settings.showInformationalNotes = false;
    const uint64_t replaceTask      = scope.RunCopy({replaceSource}, replaceDestination, NotesTask::ConflictAction::Overwrite);
    const auto replaced             = replaceTask != 0u ? FindNotesCompletedSummary(operations, replaceTask) : std::nullopt;
    if (! replaced.has_value())
        return reject(L"Replace Copy did not complete");
    const bool questionKept = std::ranges::any_of(replaced->issueDiagnostics, [](const NotesDiagnostic& issue) noexcept {
        return issue.conflictPrompt.has_value() && issue.severity == NotesSeverity::Info;
    });
    if (replaced->resultHr != S_OK || replaced->warningCount != 0u || replaced->errorCount != 0u || replaced->noteCount == 0u || ! questionKept ||
        ! IsAutoDismissableFileOperationCompletion(replaced->resultHr, replaced->warningCount, replaced->errorCount, false) ||
        ! FilesEqualBytes(replaceSource, replaceDestination / L"payload.bin"))
        return reject(std::format(L"answered Replace counted as a problem: hr=0x{:08X} warnings={} errors={} notes={} diagnostics=[{}]",
                                  static_cast<unsigned long>(replaced->resultHr),
                                  replaced->warningCount,
                                  replaced->errorCount,
                                  replaced->noteCount,
                                  DescribeNotesDiagnostics(replaced.value())));
    if (NotesProjectableTask(operations, replaceTask) || NotesIssuesPaneShowsTask(state, folder, operations, replaceTask))
        return reject(L"an answered question projected an Issues row with informational notes hidden");
    settings.showInformationalNotes = true;
    if (! NotesProjectableTask(operations, replaceTask) || ! NotesIssuesPaneShowsTask(state, folder, operations, replaceTask))
        return reject(L"an answered question did not project its Issues row with informational notes shown");
    settings.showInformationalNotes = false;

    // 2. Skip on that same question, with Info logging off: Partial, card kept, zero counters, the
    //    typed skip note recorded anyway (notes bypass the log-verbosity gate).
    const auto skipDestination = root / L"skip-dst";
    if (! RecreateEmptyDirectory(skipDestination) || ! WriteTestFile(skipDestination / L"payload.bin", 512u))
        return reject(L"Skip fixture unavailable");
    settings.diagnosticsInfoEnabled = false;
    const uint64_t skipTask         = scope.RunCopy({replaceSource}, skipDestination, NotesTask::ConflictAction::Skip);
    settings.diagnosticsInfoEnabled = true;
    const auto skipped              = skipTask != 0u ? FindNotesCompletedSummary(operations, skipTask) : std::nullopt;
    if (! skipped.has_value())
        return reject(L"Skip Copy did not complete");
    const size_t skipNotes = CountNotesDiagnostics(skipped.value(), L"task.skippedAtRequest", NotesSeverity::Info);
    if ((skipped->resultHr != S_FALSE && skipped->resultHr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)) || skipped->warningCount != 0u ||
        skipped->errorCount != 0u || skipped->skippedItemCount != 1u || ! skipped->partialExplainedByUserSkips || skipNotes != 1u ||
        IsAutoDismissableFileOperationCompletion(skipped->resultHr, skipped->warningCount, skipped->errorCount, true) ||
        ! FileSizeEquals(skipDestination / L"payload.bin", 512u))
        return reject(std::format(L"a user Skip was not a retained Information note: hr=0x{:08X} warnings={} errors={} skipped={} explained={} "
                                  L"skipNotes={} diagnostics=[{}]",
                                  static_cast<unsigned long>(skipped->resultHr),
                                  skipped->warningCount,
                                  skipped->errorCount,
                                  skipped->skippedItemCount,
                                  skipped->partialExplainedByUserSkips,
                                  skipNotes,
                                  DescribeNotesDiagnostics(skipped.value())));

    // 3. A real failure without any skip still counts: one root cannot be opened (another program holds
    //    it exclusively) and continue-on-error keeps going without a prompt; the other completes. The
    //    task is Partial with a counted problem.
    const auto failSourceA     = root / L"fail-src" / L"a.bin";
    const auto failSourceB     = root / L"fail-src" / L"b.bin";
    const auto failDestination = root / L"fail-dst";
    if (! WriteTestFile(failSourceA, 2048u) || ! WriteTestFile(failSourceB, 2048u) || ! RecreateEmptyDirectory(failDestination))
        return reject(L"failure fixture unavailable");
    wil::unique_hfile exclusive(
        CreateFileW(MakeExtendedPathForSelfTest(failSourceB).c_str(), GENERIC_READ, 0u, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! exclusive)
        return reject(L"failure fixture could not hold its source exclusively");
    const uint64_t failTask = scope.RunCopy(
        {failSourceA, failSourceB}, failDestination, std::nullopt, static_cast<FileSystemFlags>(FILESYSTEM_FLAG_RECURSIVE | FILESYSTEM_FLAG_CONTINUE_ON_ERROR));
    exclusive.reset();
    const auto failed = failTask != 0u ? FindNotesCompletedSummary(operations, failTask) : std::nullopt;
    if (! failed.has_value())
        return reject(L"failure Copy did not complete");
    if (failed->resultHr != HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) || failed->warningCount + failed->errorCount == 0u || failed->partialExplainedByUserSkips ||
        CountNotesDiagnostics(failed.value(), L"task.skippedAtRequest", NotesSeverity::Info) != 0u ||
        IsAutoDismissableFileOperationCompletion(failed->resultHr, failed->warningCount, failed->errorCount, false))
        return reject(std::format(L"a partial with a real failure lost its counted problem: hr=0x{:08X} warnings={} errors={} explained={} diagnostics=[{}]",
                                  static_cast<unsigned long>(failed->resultHr),
                                  failed->warningCount,
                                  failed->errorCount,
                                  failed->partialExplainedByUserSkips,
                                  DescribeNotesDiagnostics(failed.value())));

    Debug::Perf::Emit(L"FileOps.SelfTest.NotesSeverity",
                      L"replace=note+done+autodismiss;issues-hidden-then-shown;skip=partial-note-info-off;failure=counted",
                      0u,
                      3u,
                      skipNotes,
                      S_OK);
    return true;
}

// R5.4/R5.5 and S1b (R5.9): inherited permissions are one Information note per task on both Copy
// routes, a protected source DACL replaced by inheritance is a Warning on both, the native route
// reports metadata it could not keep, and an Overwrite keeps named streams and MOTW.
[[nodiscard]] bool RunFileOpsNotesMetadataChecks(SelfTestState& state) noexcept
{
    const auto reject = [](std::wstring_view reason) noexcept
    {
        AppendLog(std::format(L"Notes metadata: {}", reason));
        return false;
    };
    auto* folder     = TryGetFolderWindow(state.mainWindow);
    auto* operations = TryGetFileOps(folder);
    if (! folder || ! operations || ! state.fsLocal || ! CancelOutstandingSelfTestOperations(operations))
        return reject(L"owner, Local provider or initial worker quiet point unavailable");
    RedSalamander::TestSupport::ScopedWindowActivationBlocker blocker;
    if (! RedSalamander::TestSupport::ScopedWindowActivationBlocker::IsActiveForCurrentThread() && ! blocker.Start())
        return reject(L"no-activation guard unavailable");

    const std::filesystem::path root = state.tempRoot / L"notes-metadata";
    if (! RecreateEmptyDirectory(root))
        return reject(L"case root unavailable");
    const auto removeRoot                    = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(root)); });
    constexpr std::wstring_view kFailMaskEnv = L"REDSALAMANDER_FILEOPS_METADATA_FAIL_MASK";
    const auto clearFailMask                 = wil::scope_exit([&]() noexcept { static_cast<void>(SetEnvironmentVariableW(kFailMaskEnv.data(), nullptr)); });
    NotesCaseScope scope(state, folder, operations);
    auto& settings = EnsureFileOperationsSettingsForSelfTest();

    // Three plain files in an ordinary folder; each route copies them into its own folder whose
    // inheritable Guests ACE the sources lack.
    const auto sourceDir = root / L"inherit-src";
    std::vector<std::filesystem::path> sources;
    for (const wchar_t* leaf : {L"one.bin", L"two.bin", L"three.bin"})
    {
        sources.push_back(sourceDir / leaf);
        if (! WriteTestFile(sources.back(), 4096u))
            return reject(L"inheritance source fixture unavailable");
    }
    struct RouteCase final
    {
        const wchar_t* name;
        bool verify;
    };
    constexpr std::array<RouteCase, 2> routes{{{L"native", false}, {L"bridge", true}}};
    for (const RouteCase& route : routes)
    {
        const auto destination = root / std::format(L"inherit-dst-{}", route.name);
        if (! RecreateEmptyDirectory(destination) || ! AddInheritableGuestsReadForNotes(destination))
            return reject(std::format(L"{} inheritable destination fixture unavailable", route.name));
        settings.verifyAfterCopy = route.verify;
        const uint64_t taskId    = scope.RunCopy(sources, destination);
        const auto summary       = taskId != 0u ? FindNotesCompletedSummary(operations, taskId) : std::nullopt;
        if (! summary.has_value())
            return reject(std::format(L"{} inherited-ACL Copy did not complete", route.name));
        const auto inherited = std::ranges::find_if(summary->issueDiagnostics, [](const NotesDiagnostic& issue) noexcept {
            return issue.category == L"metadata.security.inherited" && issue.severity == NotesSeverity::Info;
        });
        const auto* count    = inherited != summary->issueDiagnostics.end() ? std::get_if<unsigned long>(&inherited->userText.argument) : nullptr;
        if (summary->resultHr != S_OK || summary->warningCount != 0u || summary->errorCount != 0u ||
            CountNotesDiagnostics(summary.value(), L"metadata.security.inherited", NotesSeverity::Info) != 1u || count == nullptr || *count != 3u ||
            inherited->userText.messageId != IDS_FMT_FILEOPS_NOTE_PERMISSIONS_INHERITED || HasNotesProblemDiagnostic(summary.value(), L"metadata.") ||
            (route.verify && summary->verifiedItemCount != 3u) || (! route.verify && summary->verifiedItemCount != 0u))
            return reject(std::format(L"{} route did not record exactly one inherited-permissions note for 3 files: hr=0x{:08X} warnings={} "
                                      L"errors={} verified={} diagnostics=[{}]",
                                      route.name,
                                      static_cast<unsigned long>(summary->resultHr),
                                      summary->warningCount,
                                      summary->errorCount,
                                      summary->verifiedItemCount,
                                      DescribeNotesDiagnostics(summary.value())));
    }

    // A source whose DACL was protected lands on an inheriting destination: a Warning on both routes.
    const auto protectedSource = root / L"protected-src" / L"private.bin";
    if (! WriteTestFile(protectedSource, 2048u) || ! ProtectDaclForNotes(protectedSource))
        return reject(L"protected-DACL source fixture unavailable");
    for (const RouteCase& route : routes)
    {
        const auto destination = root / std::format(L"protected-dst-{}", route.name);
        if (! RecreateEmptyDirectory(destination))
            return reject(std::format(L"{} protected destination fixture unavailable", route.name));
        settings.verifyAfterCopy = route.verify;
        const uint64_t taskId    = scope.RunCopy({protectedSource}, destination);
        const auto summary       = taskId != 0u ? FindNotesCompletedSummary(operations, taskId) : std::nullopt;
        if (! summary.has_value())
            return reject(std::format(L"{} protected-DACL Copy did not complete", route.name));
        const bool replacedWarning = std::ranges::any_of(summary->issueDiagnostics,
                                                         [](const NotesDiagnostic& issue) noexcept
        {
            return issue.category == L"metadata.security.changed" && issue.severity == NotesSeverity::Warning &&
                   issue.userText.messageId == IDS_FILEOPS_NOTE_METADATA_SECURITY_REPLACED;
        });
        if (FAILED(summary->resultHr) || summary->warningCount == 0u || ! replacedWarning ||
            CountNotesDiagnostics(summary.value(), L"metadata.security.inherited", NotesSeverity::Info) != 0u)
            return reject(std::format(L"{} route did not warn that a protected DACL was replaced: hr=0x{:08X} warnings={} diagnostics=[{}]",
                                      route.name,
                                      static_cast<unsigned long>(summary->resultHr),
                                      summary->warningCount,
                                      DescribeNotesDiagnostics(summary.value())));
    }
    settings.verifyAfterCopy = false;

    // S1b: the native route reports metadata it could not keep (deterministic loss injection).
    constexpr std::string_view kMotw   = "[ZoneTransfer]\r\nZoneId=3\r\n";
    constexpr std::string_view kStream = "named stream payload";
    const auto streamSource            = root / L"streams-src" / L"downloaded.bin";
    if (! WriteTestFile(streamSource, 4096u) || ! WriteNotesStream(streamSource, L"Zone.Identifier", kMotw) ||
        ! WriteNotesStream(streamSource, L"note", kStream))
        return reject(L"stream source fixture unavailable");
    {
        const auto destination  = root / L"streams-lost-dst";
        const std::wstring mask = std::format(L"0x{:X}", static_cast<uint32_t>(FILESYSTEM_METADATA_MOTW | FILESYSTEM_METADATA_ALTERNATE_STREAMS));
        if (! RecreateEmptyDirectory(destination) || SetEnvironmentVariableW(kFailMaskEnv.data(), mask.c_str()) == FALSE)
            return reject(L"native loss fixture unavailable");
        const uint64_t taskId = scope.RunCopy({streamSource}, destination);
        static_cast<void>(SetEnvironmentVariableW(kFailMaskEnv.data(), nullptr));
        const auto summary = taskId != 0u ? FindNotesCompletedSummary(operations, taskId) : std::nullopt;
        if (! summary.has_value())
            return reject(L"native loss Copy did not complete");
        if (FAILED(summary->resultHr) || summary->warningCount == 0u ||
            CountNotesDiagnostics(summary.value(), L"metadata.motw.lost", NotesSeverity::Warning) != 1u ||
            CountNotesDiagnostics(summary.value(), L"metadata.alternateStreams.lost", NotesSeverity::Warning) != 1u ||
            ReadNotesStream(destination / streamSource.filename(), L"Zone.Identifier").has_value())
            return reject(std::format(L"native Copy lost MOTW/streams silently: hr=0x{:08X} warnings={} diagnostics=[{}]",
                                      static_cast<unsigned long>(summary->resultHr),
                                      summary->warningCount,
                                      DescribeNotesDiagnostics(summary.value())));
    }

    // S1b: an Overwrite (staged route) keeps named streams and MOTW on NTFS.
    {
        const auto destination = root / L"streams-overwrite-dst";
        const auto target      = destination / streamSource.filename();
        if (! RecreateEmptyDirectory(destination) || ! WriteTestFile(target, 128u))
            return reject(L"Overwrite stream fixture unavailable");
        const uint64_t taskId = scope.RunCopy({streamSource}, destination, NotesTask::ConflictAction::Overwrite);
        const auto summary    = taskId != 0u ? FindNotesCompletedSummary(operations, taskId) : std::nullopt;
        if (! summary.has_value())
            return reject(L"Overwrite stream Copy did not complete");
        const auto motw   = ReadNotesStream(target, L"Zone.Identifier");
        const auto stream = ReadNotesStream(target, L"note");
        if (summary->resultHr != S_OK || summary->warningCount != 0u || ! motw.has_value() || motw.value() != kMotw || ! stream.has_value() ||
            stream.value() != kStream || ! FilesEqualBytes(streamSource, target) || HasNotesProblemDiagnostic(summary.value(), L"metadata."))
            return reject(std::format(L"Overwrite did not keep MOTW and named streams: hr=0x{:08X} warnings={} motw={} stream={} diagnostics=[{}]",
                                      static_cast<unsigned long>(summary->resultHr),
                                      summary->warningCount,
                                      motw.has_value(),
                                      stream.has_value(),
                                      DescribeNotesDiagnostics(summary.value())));
    }

    // S1b witness on a real destination that may not keep named streams (the registered SMB share):
    // every class is either kept or reported, never silently dropped.
    const std::wstring smbRootText = GetEnvVarTrimmed(L"REDSALAMANDER_SELFTEST_SMB_ROOT");
    size_t smbWitnessed            = 0u;
    if (smbRootText.empty())
    {
        AppendLog(L"Notes metadata: SMB witness skipped (path.smb.primary is not configured).");
    }
    else
    {
        const std::filesystem::path smbDestination = std::filesystem::path(smbRootText) / std::format(L"notes-metadata-{}", GetCurrentProcessId());
        const auto removeSmb                       = wil::scope_exit([&]() noexcept { static_cast<void>(SelfTest::RemoveAll(smbDestination)); });
        if (! RecreateEmptyDirectory(smbDestination))
            return reject(std::format(L"SMB destination '{}' unavailable", smbDestination.native()));
        const uint64_t taskId = scope.RunCopy({streamSource}, smbDestination);
        const auto summary    = taskId != 0u ? FindNotesCompletedSummary(operations, taskId) : std::nullopt;
        if (! summary.has_value() || FAILED(summary->resultHr))
            return reject(
                std::format(L"native Copy to the SMB share failed: hr=0x{:08X}", summary.has_value() ? static_cast<unsigned long>(summary->resultHr) : 0ul));
        const auto target = smbDestination / streamSource.filename();
        for (const auto& [stream, category] : {std::pair{L"Zone.Identifier", L"metadata.motw.lost"}, std::pair{L"note", L"metadata.alternateStreams.lost"}})
        {
            const bool kept     = ReadNotesStream(target, stream).has_value();
            const bool reported = CountNotesDiagnostics(summary.value(), category, NotesSeverity::Warning) == 1u;
            if (kept == reported)
                return reject(std::format(L"SMB stream '{}' kept={} reported={}: a class must be kept or reported, diagnostics=[{}]",
                                          stream,
                                          kept,
                                          reported,
                                          DescribeNotesDiagnostics(summary.value())));
            AppendLog(std::format(L"Notes metadata: SMB stream '{}' {}.", stream, kept ? L"kept" : L"reported lost"));
        }
        smbWitnessed = 1u;
    }

    Debug::Perf::Emit(L"FileOps.SelfTest.NotesMetadata",
                      L"inherited=one-note-per-task(native,bridge);protected=warning(native,bridge);native-loss=reported;overwrite=streams-kept",
                      0u,
                      6u,
                      smbWitnessed,
                      S_OK);
    return true;
}
