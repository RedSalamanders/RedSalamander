#include "FolderWindow.FileOperations.State.Private.h"
#include "WindowActivation.h"

#include "FileOperationConfirmation.h"
#include "FolderWindow.FileOperations.Popup.h"
#include "HostServices.h"
#include "NavigationLocation.h"
#include "SessionState.h"
#include "SettingsHotReload.h"

#include <algorithm>
#include <limits>
#include <numeric>
#include <tuple>

using namespace FolderWindowFileOperationsStateInternal;

namespace
{
[[nodiscard]] uint64_t SaturatingAdd(const uint64_t left, const size_t right) noexcept
{
    const uint64_t converted =
        right > static_cast<size_t>((std::numeric_limits<uint64_t>::max)()) ? (std::numeric_limits<uint64_t>::max)() : static_cast<uint64_t>(right);
    return converted > (std::numeric_limits<uint64_t>::max)() - left ? (std::numeric_limits<uint64_t>::max)() : left + converted;
}

[[nodiscard]] HRESULT BuildMoveBreadcrumbRecord(const FolderWindow::FileOperationState::Task& task,
                                                const FolderWindow::FileOperationState::SelectedAdmissionView& selectedAdmission,
                                                FileOperationMoveBreadcrumb::Record& out)
{
    const auto& plans = selectedAdmission.plans;
    if (task._operation != FILESYSTEM_MOVE || ! selectedAdmission.destination || ! plans)
    {
        return E_INVALIDARG;
    }

    FileOperationMoveBreadcrumb::Record record{};
    record.taskId = task._taskId;
    record.sourcePane =
        task._sourcePane == FolderWindow::Pane::Right ? FileOperationMoveBreadcrumb::PaneHint::Right : FileOperationMoveBreadcrumb::PaneHint::Left;
    const FolderWindow::Pane destinationPane =
        selectedAdmission.destination->pane.value_or(task._sourcePane == FolderWindow::Pane::Left ? FolderWindow::Pane::Right : FolderWindow::Pane::Left);
    record.destinationPane =
        destinationPane == FolderWindow::Pane::Right ? FileOperationMoveBreadcrumb::PaneHint::Right : FileOperationMoveBreadcrumb::PaneHint::Left;
    std::vector<std::tuple<std::wstring, std::wstring, std::wstring, std::wstring>> uniqueSourceRoots;
    for (const FileOperations::FileOperationPlan& plan : *plans)
    {
        const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
        if (transfer == nullptr || transfer->intent != FileOperations::TransferIntent::Move || transfer->selectedItems.empty())
        {
            return E_INVALIDARG;
        }

        const size_t itemCount = transfer->selectedItems.size();
        switch (transfer->strategy)
        {
            case FileOperations::OperationStrategy::Native:
                record.admittedStrategies.nativeItems = SaturatingAdd(record.admittedStrategies.nativeItems, itemCount);
                break;
            case FileOperations::OperationStrategy::Managed:
                record.admittedStrategies.managedItems = SaturatingAdd(record.admittedStrategies.managedItems, itemCount);
                break;
            case FileOperations::OperationStrategy::CopyOnly:
                record.admittedStrategies.copyOnlyItems = SaturatingAdd(record.admittedStrategies.copyOnlyItems, itemCount);
                break;
            case FileOperations::OperationStrategy::Copy: return E_INVALIDARG;
        }

        const auto key = std::make_tuple(
            transfer->sourceEndpoint.pluginId, transfer->sourceEndpoint.instanceId, transfer->sourceEndpoint.profileId, transfer->sourceEndpoint.rootId);
        if (std::ranges::find(uniqueSourceRoots, key) == uniqueSourceRoots.end())
        {
            uniqueSourceRoots.push_back(key);
            if (record.sourceRootSamples.size() < FileOperationMoveBreadcrumb::kMaximumPersistedSourceRootSamples)
            {
                record.sourceRootSamples.push_back(FileOperationMoveBreadcrumb::QualifiedLocation{
                    .pluginId           = transfer->sourceEndpoint.pluginId,
                    .pluginShortId      = task._sourcePluginShortId,
                    .instanceId         = transfer->sourceEndpoint.instanceId,
                    .profileId          = transfer->sourceEndpoint.profileId,
                    .rootId             = transfer->sourceEndpoint.rootId,
                    .representativePath = transfer->selectedItems.front().providerPath,
                });
            }
        }

        if (record.destination.pluginId.empty())
        {
            record.destination = FileOperationMoveBreadcrumb::QualifiedLocation{
                .pluginId           = transfer->destinationEndpoint.pluginId,
                .pluginShortId      = selectedAdmission.destination->pluginShortId,
                .instanceId         = transfer->destinationEndpoint.instanceId,
                .profileId          = transfer->destinationEndpoint.profileId,
                .rootId             = transfer->destinationEndpoint.rootId,
                .representativePath = transfer->destination.providerFolderPath,
            };
        }
    }
    record.sourceRootCount            = static_cast<uint64_t>(uniqueSourceRoots.size());
    record.sourceRootSamplesTruncated = uniqueSourceRoots.size() > record.sourceRootSamples.size();
    out                               = std::move(record);
    return S_OK;
}

[[nodiscard]] FileOperationArtifacts::Endpoint ToArtifactEndpoint(const FileOperations::QualifiedEndpoint& endpoint)
{
    return FileOperationArtifacts::Endpoint{
        .pluginId   = endpoint.pluginId,
        .instanceId = endpoint.instanceId,
        .profileId  = endpoint.profileId,
        .rootId     = endpoint.rootId,
    };
}

[[nodiscard]] bool SameArtifactEndpoint(const FileOperationArtifacts::Endpoint& left, const FileOperationArtifacts::Endpoint& right) noexcept
{
    return left.pluginId == right.pluginId && left.instanceId == right.instanceId && left.profileId == right.profileId && left.rootId == right.rootId;
}

HRESULT AppendArtifactTouchCandidate(IFileSystem* fileSystem,
                                     const FileOperations::QualifiedEndpoint& endpoint,
                                     const std::wstring_view providerPath,
                                     std::vector<FileOperationArtifacts::Candidate>& out,
                                     const FileOperationArtifacts::DeleteHandoff deleteHandoff = FileOperationArtifacts::DeleteHandoff::None)
{
    if (fileSystem == nullptr || providerPath.empty() || ! endpoint.pathIdentity.has_value())
    {
        return E_INVALIDARG;
    }

    FileOperationArtifacts::Candidate candidate{
        .endpoint      = ToArtifactEndpoint(endpoint),
        .pathIdentity  = endpoint.pathIdentity.value(),
        .path          = std::wstring(providerPath),
        .probeState    = FileOperationArtifacts::ProbeState::Indeterminate,
        .deleteHandoff = deleteHandoff,
    };
    if (FileOperationArtifacts::ClassifyCandidate(candidate).classification == FileOperationArtifacts::Classification::Ordinary)
    {
        return S_FALSE;
    }

    constexpr FileSystemBindFlags bindFlags   = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
    FileOperations::ObjectBindingResult bound = FileOperations::BindObjectAuthority(fileSystem, providerPath, endpoint.profileId, bindFlags);
    if (bound.state == FileOperations::ObjectBindingState::Missing)
    {
        // A missing publication name is not an artifact object being touched. The normal operation
        // still owns its creation/conflict checks.
        return S_FALSE;
    }
    candidate.identityProbe = FileOperations::ToArtifactIdentityProbe(bound);
    if (bound.state == FileOperations::ObjectBindingState::Bound)
    {
        candidate.probeState      = FileOperationArtifacts::ProbeState::Present;
        candidate.currentIdentity = FileOperationArtifacts::Identity{
            .objectId      = std::move(bound.authority.identity.objectId),
            .revisionId    = std::move(bound.authority.identity.revisionId),
            .pathProfileId = std::move(bound.authority.identity.pathProfileId),
            .kind          = bound.authority.kind,
        };
    }

    const auto duplicate = std::ranges::find_if(out,
                                                [&](const FileOperationArtifacts::Candidate& existing) noexcept
    {
        return SameArtifactEndpoint(existing.endpoint, candidate.endpoint) &&
               EquivalentPath(candidate.pathIdentity, existing.path.native(), candidate.path.native());
    });
    if (duplicate == out.end())
    {
        out.push_back(std::move(candidate));
    }
    else if (duplicate->deleteHandoff != deleteHandoff)
    {
        // The same object reached by two routes keeps the guard unless both hand it off alike.
        duplicate->deleteHandoff = FileOperationArtifacts::DeleteHandoff::None;
    }
    return S_OK;
}

// I35 R7.3/R7.6: who confirms a Delete item that has no exact identity, from the plan and the
// provider's interfaces alone (no provider I/O). A provider without object binding confirms every
// Delete on the permanent-delete card (C1/C10); the Local provider confirms an identity-less object by
// name on that card (Permanent) or on its Recycle escalation (Recycle). An archive cleanup carries a
// consent captured before the task and asks nothing more, so its items keep the guard's refusal, as does
// a provider with object binding but no by-name route.
[[nodiscard]] FileOperationArtifacts::DeleteHandoff DeleteHandoffForPlan(const FileOperations::DeletePlan& plan, IFileSystem* fileSystem) noexcept
{
    if (fileSystem == nullptr || plan.origin == FileOperations::DeleteOrigin::PackCleanup || plan.origin == FileOperations::DeleteOrigin::UnpackCleanup)
    {
        return FileOperationArtifacts::DeleteHandoff::None;
    }
    const FileOperationArtifacts::DeleteHandoff byMode = plan.mode == FileOperations::DeleteMode::Permanent
                                                             ? FileOperationArtifacts::DeleteHandoff::PermanentCard
                                                             : FileOperationArtifacts::DeleteHandoff::RecycleWarn;
    wil::com_ptr<IFileSystemObjectBinding> binding;
    const HRESULT bindingHr = fileSystem->QueryInterface(__uuidof(IFileSystemObjectBinding), binding.put_void());
    if (bindingHr == E_NOINTERFACE && ! binding)
    {
        return byMode;
    }
    // The by-name route of a binding provider is the Local provider's (D-7.3), as Preparing honors it.
    wil::com_ptr<IFileSystemIdentityDelete> identityDelete;
    if (FAILED(bindingHr) || ! binding || ! FileOperations::IsLocalFileSystemEndpoint(plan.endpoint) ||
        FAILED(fileSystem->QueryInterface(__uuidof(IFileSystemIdentityDelete), identityDelete.put_void())) || ! identityDelete)
    {
        return FileOperationArtifacts::DeleteHandoff::None;
    }
    return byMode;
}

HRESULT CollectArtifactTouchCandidates(const FileOperations::FileOperationPlanGroup& plans,
                                       IFileSystem* sourceFileSystem,
                                       IFileSystem* destinationFileSystem,
                                       std::vector<FileOperationArtifacts::Candidate>& out)
{
    Debug::Perf::Scope perf(L"fileops.artifact.touch.capture.us");
    out.clear();
    IFileSystem* const effectiveDestination = destinationFileSystem != nullptr ? destinationFileSystem : sourceFileSystem;
    for (const FileOperations::FileOperationPlan& plan : plans)
    {
        const HRESULT planHr = std::visit(
            [&](const auto& typedPlan) -> HRESULT
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::TransferPlan>)
            {
                if (! typedPlan.sourceEndpoint.pathIdentity.has_value() || ! typedPlan.destinationEndpoint.pathIdentity.has_value())
                {
                    return E_INVALIDARG;
                }
                for (size_t index = 0u; index < typedPlan.selectedItems.size(); ++index)
                {
                    const FileOperations::QualifiedSourceItem& item = typedPlan.selectedItems[index];
                    if (typedPlan.intent == FileOperations::TransferIntent::Move)
                    {
                        const HRESULT sourceHr = AppendArtifactTouchCandidate(sourceFileSystem, typedPlan.sourceEndpoint, item.providerPath, out);
                        if (FAILED(sourceHr))
                        {
                            return sourceHr;
                        }
                    }

                    std::wstring destinationPath;
                    const auto explicitMapping = std::ranges::find_if(typedPlan.explicitMappings,
                                                                      [index](const FileOperations::TransferDestinationMapping& mapping) noexcept
                    { return mapping.sourceIndex == index; });
                    if (explicitMapping != typedPlan.explicitMappings.end())
                    {
                        destinationPath = explicitMapping->destinationProviderPath;
                    }
                    else
                    {
                        std::wstring leaf;
                        if (! TryGetFileSystemLeafName(typedPlan.sourceEndpoint.pathIdentity.value(), item.providerPath, leaf))
                        {
                            return E_INVALIDARG;
                        }
                        destinationPath =
                            JoinFileSystemPath(typedPlan.destinationEndpoint.pathIdentity.value(), typedPlan.destination.providerFolderPath, leaf);
                    }
                    const HRESULT destinationHr = AppendArtifactTouchCandidate(effectiveDestination, typedPlan.destinationEndpoint, destinationPath, out);
                    if (FAILED(destinationHr))
                    {
                        return destinationHr;
                    }
                }
                return S_OK;
            }
            else if constexpr (std::is_same_v<Plan, FileOperations::RenamePlan>)
            {
                if (! typedPlan.endpoint.pathIdentity.has_value())
                {
                    return E_INVALIDARG;
                }
                for (const FileOperations::RenameStep& step : typedPlan.finalMappings)
                {
                    const HRESULT sourceHr = AppendArtifactTouchCandidate(sourceFileSystem, typedPlan.endpoint, step.source.providerPath, out);
                    if (FAILED(sourceHr))
                    {
                        return sourceHr;
                    }
                    // An inline rename publishes with its provider join pending until Preparing
                    // fills it (C1). Until then the destination candidate is the host join of the
                    // source's parent and the final leaf, as a transfer's implicit destination is.
                    std::wstring destinationPath = step.providerJoinedPath;
                    if (destinationPath.empty())
                    {
                        std::wstring parent;
                        if (! TryGetFileSystemParentPath(typedPlan.endpoint.pathIdentity.value(), step.source.providerPath, parent))
                        {
                            return E_INVALIDARG;
                        }
                        destinationPath = JoinFileSystemPath(typedPlan.endpoint.pathIdentity.value(), parent, step.finalLeafName);
                    }
                    const HRESULT destinationHr = AppendArtifactTouchCandidate(sourceFileSystem, typedPlan.endpoint, destinationPath, out);
                    if (FAILED(destinationHr))
                    {
                        return destinationHr;
                    }
                }
                return S_OK;
            }
            else
            {
                const FileOperationArtifacts::DeleteHandoff deleteHandoff = DeleteHandoffForPlan(typedPlan, sourceFileSystem);
                for (const FileOperations::QualifiedSourceItem& item : typedPlan.selectedItems)
                {
                    const HRESULT sourceHr = AppendArtifactTouchCandidate(sourceFileSystem, typedPlan.endpoint, item.providerPath, out, deleteHandoff);
                    if (FAILED(sourceHr))
                    {
                        return sourceHr;
                    }
                }
                return S_OK;
            }
        },
            plan);
        if (FAILED(planHr))
        {
            perf.SetHr(planHr);
            return planHr;
        }
    }
    perf.SetValue0(static_cast<uint64_t>(out.size()));
    perf.SetHr(S_OK);
    return out.empty() ? S_FALSE : S_OK;
}

// I35 R7.1: the touch guard's fail-closed answer names each item that has no current exact
// identity, says why in plain words and logs both HRESULTs with the path. The system message of a
// failed open is quoted without its trailing period, inside the sentence that explains it.
[[nodiscard]] std::wstring FormatArtifactTouchBlockedReason(const FileOperationArtifacts::TouchGuardItem& item)
{
    using FileOperationArtifacts::IdentityProbeOutcome;
    if (item.identityProbe.outcome == IdentityProbeOutcome::OpenFailed)
    {
        std::wstring systemText = TryFormatSystemMessage(GetSystemMessageIdFromHResult(item.identityProbe.openStatus));
        while (! systemText.empty() && (systemText.back() == L'.' || systemText.back() == L'\r' || systemText.back() == L'\n' || systemText.back() == L' '))
            systemText.pop_back();
        // A status without system text is still named by the item line's (0x...) code.
        if (systemText.empty())
            return LoadStringResource(nullptr, IDS_FILEOPS_ARTIFACT_BLOCKED_REASON_OPEN_UNKNOWN);
        return FormatStringResource(nullptr, IDS_FMT_FILEOPS_ARTIFACT_BLOCKED_REASON_OPEN, systemText);
    }
    return LoadStringResource(nullptr, IDS_FILEOPS_ARTIFACT_BLOCKED_REASON_UNSUPPORTED);
}

[[nodiscard]] HRESULT ArtifactTouchBlockedStatus(const FileOperationArtifacts::TouchGuardItem& item) noexcept
{
    if (item.identityProbe.outcome == FileOperationArtifacts::IdentityProbeOutcome::OpenFailed)
        return item.identityProbe.openStatus;
    return FAILED(item.identityProbe.identityStatus) ? item.identityProbe.identityStatus : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

// S_OK accepted with a receipt; S_FALSE declined; HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) blocked and
// already explained to the user; HRESULT_FROM_WIN32(ERROR_RETRY) when the caller offered Retry and the
// user chose it (the caller captures the candidates again and asks again).
HRESULT ConfirmArtifactTouch(HWND owner,
                             const FileOperationArtifacts::TouchGuardRequest& request,
                             FileOperationArtifacts::TouchGuardReceipt& receipt,
                             const bool retryAvailable = false)
{
    receipt                  = {};
    const uint64_t total     = static_cast<uint64_t>(request.items.size());
    const std::wstring title = LoadStringResource(nullptr, IDS_FILEOPS_ARTIFACT_TOUCH_TITLE);
    const bool blocked       = ! request.allItemsRevalidatable;

    std::wstring message;
    bool offerRetry = false;
    // The File Explorer hint is given only for an item the location itself cannot identify; an item
    // that could not be opened is explained by its own Retry line (R7.1).
    bool identityUnsupportedWin32Item = false;
    std::optional<std::filesystem::path> openLocationPath;
    if (blocked)
    {
        constexpr size_t kMaxNamedItems = 3u;
        std::wstring itemLines;
        size_t namedItems   = 0u;
        size_t blockedItems = 0u;
        for (const FileOperationArtifacts::TouchGuardItem& item : request.items)
        {
            // A handed-off Recycle item (I35 R7.3) is not what blocks the request.
            if (item.revalidatable || item.handedOff)
                continue;
            ++blockedItems;
            const HRESULT blockedStatus = ArtifactTouchBlockedStatus(item);
            const bool openFailed       = item.identityProbe.outcome == FileOperationArtifacts::IdentityProbeOutcome::OpenFailed;
            // A name-shape-only item (no object binding on this location) was never opened or queried.
            const bool notProbed = item.identityProbe.outcome == FileOperationArtifacts::IdentityProbeOutcome::NotProbed;
            const std::wstring openText =
                notProbed ? std::wstring(L"not probed") : std::format(L"0x{:08X}", static_cast<unsigned long>(item.identityProbe.openStatus));
            const std::wstring identityText = notProbed    ? std::wstring(L"not probed")
                                              : openFailed ? std::wstring(L"not queried")
                                                           : std::format(L"0x{:08X}", static_cast<unsigned long>(item.identityProbe.identityStatus));
            Debug::Warning(L"File Operations refused to touch a possible leftover without a current exact identity: path='{}' openHr={} identityHr={}.",
                           item.path.native(),
                           openText,
                           identityText);
            Debug::Perf::Emit(L"fileops.artifact.touch.blocked",
                              item.path.filename().native(),
                              static_cast<uint64_t>(static_cast<uint32_t>(item.identityProbe.openStatus)),
                              static_cast<uint64_t>(static_cast<uint32_t>(item.identityProbe.identityStatus)),
                              static_cast<uint64_t>(item.identityProbe.outcome),
                              blockedStatus);
            offerRetry           = offerRetry || (retryAvailable && openFailed);
            const bool win32Item = FileOperationArtifacts::IsWin32FileSystemEndpoint(item.endpoint);
            if (! openLocationPath.has_value() && win32Item)
                openLocationPath = item.path;
            identityUnsupportedWin32Item =
                identityUnsupportedWin32Item || (win32Item && item.identityProbe.outcome == FileOperationArtifacts::IdentityProbeOutcome::IdentityUnsupported);
            if (namedItems < kMaxNamedItems)
            {
                if (! itemLines.empty())
                    itemLines += L"\r\n";
                itemLines += FormatStringResource(nullptr,
                                                  IDS_FMT_FILEOPS_ARTIFACT_BLOCKED_ITEM,
                                                  item.path.filename().native(),
                                                  static_cast<unsigned long>(blockedStatus),
                                                  FormatArtifactTouchBlockedReason(item));
                ++namedItems;
            }
        }
        if (blockedItems > namedItems)
        {
            itemLines += L"\r\n";
            itemLines += FormatStringResource(nullptr, IDS_FMT_FILEOPS_ARTIFACT_BLOCKED_MORE, static_cast<unsigned long>(blockedItems - namedItems));
        }
        if (identityUnsupportedWin32Item)
        {
            itemLines += L"\r\n\r\n";
            itemLines += LoadStringResource(nullptr, IDS_FILEOPS_ARTIFACT_BLOCKED_NEXT_OPEN_LOCATION);
        }
        message = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ARTIFACT_TOUCH_BLOCKED, total, itemLines);
    }
    else
    {
        message = FormatStringResource(nullptr, IDS_FMT_FILEOPS_ARTIFACT_TOUCH_WARNING, total);
    }

    // The blocked presentation labels Retry (Yes), Open location in File Explorer (Ok, or No when
    // Retry is also offered) and Close (Cancel, or No without Open location); Close is the default and
    // Escape action. None of them grants a receipt.
    HostPromptButtons blockedButtons = HOST_PROMPT_BUTTONS_OK;
    if (offerRetry && openLocationPath.has_value())
        blockedButtons = HOST_PROMPT_BUTTONS_YES_NO_CANCEL;
    else if (openLocationPath.has_value())
        blockedButtons = HOST_PROMPT_BUTTONS_OK_CANCEL;
    else if (offerRetry)
        blockedButtons = HOST_PROMPT_BUTTONS_YES_NO;
    HostPromptResult blockedClose = HOST_PROMPT_RESULT_CANCEL;
    if (blockedButtons == HOST_PROMPT_BUTTONS_OK)
        blockedClose = HOST_PROMPT_RESULT_OK;
    else if (blockedButtons == HOST_PROMPT_BUTTONS_YES_NO)
        blockedClose = HOST_PROMPT_RESULT_NO;

    HostPromptRequest prompt{};
    prompt.sizeBytes     = sizeof(prompt);
    prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
    prompt.severity      = HOST_ALERT_WARNING;
    prompt.buttons       = blocked ? blockedButtons : HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow  = owner;
    prompt.title         = title.c_str();
    prompt.message       = message.c_str();
    prompt.defaultResult = blocked ? blockedClose : HOST_PROMPT_RESULT_CANCEL;
    prompt.presentation  = blocked ? HOST_PROMPT_PRESENTATION_ARTIFACT_BLOCKED : HOST_PROMPT_PRESENTATION_ARTIFACT_TOUCH;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT promptHr        = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(promptHr))
    {
        return promptHr;
    }
    if (blocked)
    {
        if (offerRetry && promptResult == HOST_PROMPT_RESULT_YES)
            return HRESULT_FROM_WIN32(ERROR_RETRY);
        const bool openLocationChosen =
            openLocationPath.has_value() && ((blockedButtons == HOST_PROMPT_BUTTONS_OK_CANCEL && promptResult == HOST_PROMPT_RESULT_OK) ||
                                             (blockedButtons == HOST_PROMPT_BUTTONS_YES_NO_CANCEL && promptResult == HOST_PROMPT_RESULT_NO));
        if (openLocationChosen)
            static_cast<void>(FileOperationArtifacts::OpenLocationInFileExplorer(openLocationPath.value()));
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    if (promptResult != HOST_PROMPT_RESULT_OK)
    {
        return S_FALSE;
    }
    return FileOperationArtifacts::AcceptTouchGuard(request, receipt);
}

[[nodiscard]] bool ValidateRetainedSourceActionItems(IFileSystem* fileSystem,
                                                     std::span<const FolderWindow::FileOperationState::CompletedTaskSummary::RetainedSourceActionItem> items,
                                                     std::vector<std::filesystem::path>* exactPaths) noexcept
{
    if (fileSystem == nullptr || items.empty())
    {
        return false;
    }
    if (exactPaths)
    {
        exactPaths->clear();
        exactPaths->reserve(items.size());
    }
    for (const auto& item : items)
    {
        constexpr FileSystemBindFlags bindFlags = static_cast<FileSystemBindFlags>(FILESYSTEM_BIND_NO_FOLLOW | FILESYSTEM_BIND_READ_METADATA);
        FileOperations::ObjectBindingResult current =
            FileOperations::BindObjectAuthority(fileSystem, item.providerPath.native(), item.identity.pathProfileId, bindFlags);
        if (current.state != FileOperations::ObjectBindingState::Bound || current.authority.identity.pathProfileId != item.identity.pathProfileId ||
            current.authority.identity.objectId != item.identity.objectId)
        {
            return false;
        }
        if (exactPaths)
        {
            exactPaths->push_back(item.providerPath);
        }
    }
    return true;
}
} // namespace

HRESULT FolderWindow::FileOperationState::Task::CreateMoveBreadcrumbAfterAcceptance() noexcept
{
    if (_operation != FILESYSTEM_MOVE || _moveBreadcrumb.has_value())
    {
        return S_OK;
    }

    FileOperationMoveBreadcrumb::Record record;
    const auto selected    = LoadSelectedAdmission();
    const HRESULT recordHr = BuildMoveBreadcrumbRecord(*this, *selected, record);
    if (FAILED(recordHr))
    {
        return recordHr;
    }

    FileOperationMoveBreadcrumb::Breadcrumb breadcrumb;
    const HRESULT breadcrumbHr = FileOperationMoveBreadcrumb::Breadcrumb::Create(std::move(record), breadcrumb);
    if (FAILED(breadcrumbHr))
    {
        return breadcrumbHr;
    }
    _moveBreadcrumb.emplace(std::move(breadcrumb));
    return S_OK;
}

bool IsAutoDismissableFileOperationCompletion(HRESULT resultHr, unsigned long warningCount, unsigned long errorCount, bool partialOutcome) noexcept
{
    // Recorded issues must stay reachable: even a cancelled task keeps its card when it already
    // collected warnings or errors, so the user can open the diagnostics. A Partial without a counted
    // warning (skipped items, a Move that kept sources: I35 D-5.2) keeps its card too.
    if (warningCount != 0 || errorCount != 0 || partialOutcome)
    {
        return false;
    }

    if (IsCancellationStatus(resultHr))
    {
        return true;
    }

    return SUCCEEDED(resultHr);
}

FolderWindow::FileOperationState::FileOperationState(FolderWindow& owner) : _owner(owner)
{
    _uiLifetime = std::make_shared<int>(0);
    LoadInterruptedMoveBreadcrumbs();
}

FolderWindow::FileOperationState::~FileOperationState()
{
    Shutdown();
}

void FolderWindow::FileOperationState::LoadInterruptedMoveBreadcrumbs() noexcept
{
    std::vector<FileOperationMoveBreadcrumb::LoadedRecord> loaded;
    FileOperationMoveBreadcrumb::LoadStats stats{};
    const HRESULT loadHr = FileOperationMoveBreadcrumb::Breadcrumb::LoadInterrupted(loaded, &stats);
    if (FAILED(loadHr) && loaded.empty())
    {
        Debug::Warning(L"File Operations could not load interrupted-Move breadcrumbs (hr=0x{:08X}).", static_cast<unsigned long>(loadHr));
        return;
    }
    if (FAILED(loadHr))
    {
        Debug::Warning(L"File Operations loaded a partial interrupted-Move breadcrumb set (hr=0x{:08X}); valid records remain visible.",
                       static_cast<unsigned long>(loadHr));
    }
    if (loaded.empty())
    {
        return;
    }

    std::ranges::sort(loaded, [](const auto& left, const auto& right) noexcept { return left.record.createdFileTime > right.record.createdFileTime; });

    for (FileOperationMoveBreadcrumb::LoadedRecord& loadedRecord : loaded)
    {
        const FileOperationMoveBreadcrumb::Record& record = loadedRecord.record;
        if (record.sourceRootSamples.empty())
        {
            continue;
        }
        const FileOperationMoveBreadcrumb::QualifiedLocation& source = record.sourceRootSamples.front();
        const std::optional<std::wstring> sourceInstanceContext      = FileOperationMoveBreadcrumb::TryDecodeNavigationInstanceContext(source.instanceId);
        const std::optional<std::wstring> destinationInstanceContext =
            FileOperationMoveBreadcrumb::TryDecodeNavigationInstanceContext(record.destination.instanceId);
        if (! sourceInstanceContext.has_value() || ! destinationInstanceContext.has_value())
        {
            Debug::Warning(L"File Operations ignored an interrupted-Move breadcrumb with a noncanonical provider instance ID.");
            continue;
        }
        const UINT phaseStringId   = record.phase == FileOperationMoveBreadcrumb::DurablePhase::Executing ? IDS_FILEOPS_INTERRUPTED_MOVE_PHASE_EXECUTING
                                                                                                          : IDS_FILEOPS_INTERRUPTED_MOVE_PHASE_ADMITTED;
        const std::wstring message = FormatStringResource(nullptr,
                                                          IDS_FMT_FILEOPS_INTERRUPTED_MOVE_SUMMARY,
                                                          record.taskId,
                                                          LoadStringResource(nullptr, phaseStringId),
                                                          record.admittedStrategies.nativeItems,
                                                          record.admittedStrategies.managedItems,
                                                          record.admittedStrategies.copyOnlyItems,
                                                          record.sourceRootCount);

        CompletedTaskSummary summary{};
        summary.taskId                 = _nextTaskId++;
        summary.interruptedOperationId = record.taskId;
        summary.selection.itemCount    = record.sourceRootCount;
        if (record.sourceRootCount == 1u)
            summary.selection.singleName = source.representativePath.native(); // Recovery retains an exact path, without a live provider profile or probe.
        summary.operation           = FILESYSTEM_MOVE;
        summary.sourcePane          = record.sourcePane == FileOperationMoveBreadcrumb::PaneHint::Right ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
        summary.sourcePluginId      = source.pluginId;
        summary.sourcePluginShortId = source.pluginShortId;
        summary.sourceInstanceContext = sourceInstanceContext.value();
        summary.destinationPane = record.destinationPane == FileOperationMoveBreadcrumb::PaneHint::Right ? FolderWindow::Pane::Right : FolderWindow::Pane::Left;
        summary.destinationPluginId         = record.destination.pluginId;
        summary.destinationPluginShortId    = record.destination.pluginShortId;
        summary.destinationInstanceContext  = destinationInstanceContext.value();
        summary.destinationFolder           = record.destination.representativePath;
        summary.resultHr                    = HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED);
        summary.discoveryClosed             = true;
        summary.sourcePath                  = source.representativePath.native();
        summary.destinationPath             = record.destination.representativePath.native();
        summary.warningCount                = 0u;
        summary.lastDiagnosticMessage       = message;
        summary.resultSummary               = message;
        summary.unknownPublicationItemCount = 1u;
        summary.unknownSourceCount          = 1u;
        summary.indeterminateItemCount      = 1u;
        summary.unknownSourcePaths.push_back(source.representativePath);
        summary.completedTick                 = GetTickCount64();
        summary.interruptedMoveBreadcrumbPath = std::move(loadedRecord.path);
        summary.interruptedMoveNotice         = true;
        _completedTasks.push_back(std::move(summary));
        _issuesRevision.fetch_add(1u, std::memory_order_relaxed);
        if (_completedTasks.size() >= kMaxCompletedTaskSummaries)
        {
            break;
        }
    }
    _interruptedMoveSummariesPendingPresentation.store(! _completedTasks.empty(), std::memory_order_release);
}

void FolderWindow::FileOperationState::QueueSettingsSave(std::wstring_view context) noexcept
{
    if (! _owner._settings)
    {
        return;
    }

    const uint64_t enqueueStartUs = PerfNowUs();
    const HRESULT queueHr         = SettingsHotReload::QueueSettingsSave(kFileOpsAppId, *_owner._settings, L"FileOps.Settings.SaveUs", context);
    if (queueHr == S_OK)
    {
        _hasQueuedSettingsSave.store(true, std::memory_order_release);
    }
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Settings.EnqueueUs", context, PerfElapsedUs(enqueueStartUs), SUCCEEDED(queueHr) ? 1u : 0u, 0u, queueHr);
    }
    if (FAILED(queueHr))
    {
        Debug::Error(L"FileOperations: failed to queue settings persistence ({}) hr=0x{:08X}", context, static_cast<unsigned long>(queueHr));
    }
}

void FolderWindow::FileOperationState::FlushPendingSettingsSave() noexcept
{
    constexpr DWORD kShutdownFlushTimeoutMs = 5000u;
    if (! SettingsHotReload::FlushQueuedSettingsSaves(kShutdownFlushTimeoutMs))
    {
        Debug::Warning(L"FileOperations: queued settings persistence exceeded the {} ms shutdown deadline; copied snapshots remain owned by the worker.",
                       kShutdownFlushTimeoutMs);
    }
}

#ifdef ENABLE_TESTS
bool FolderWindow::FileOperationState::DebugFlushPendingSettingsSaveForSelfTest(DWORD timeoutMs) noexcept
{
    return SettingsHotReload::FlushQueuedSettingsSaves(timeoutMs);
}

FolderWindow::FileOperationState::SettingsSaveDebugSnapshot FolderWindow::FileOperationState::DebugGetSettingsSaveSnapshotForSelfTest() noexcept
{
    const SettingsHotReload::SettingsSaveDebugSnapshot source = SettingsHotReload::DebugGetSettingsSaveSnapshotForSelfTest();
    SettingsSaveDebugSnapshot result{};
    result.queuedGeneration    = source.queuedGeneration;
    result.completedGeneration = source.completedGeneration;
    result.coalescedCount      = source.coalescedCount;
    result.lastQueueThreadId   = source.lastQueueThreadId;
    result.lastSaveThreadId    = source.lastSaveThreadId;
    result.pending             = source.pending;
    result.saveInProgress      = source.saveInProgress;
    return result;
}
#endif

HostFileOperationVerificationAvailability FolderWindow::FileOperationState::PlannedVerificationAvailability(
    const FileOperations::FileOperationPlanGroup& plans) noexcept
{
    const auto transfers = [&](auto predicate) noexcept
    {
        return std::ranges::any_of(plans,
                                   [&](const FileOperations::FileOperationPlan& plan) noexcept
        {
            const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
            return transfer != nullptr && transfer->strategy != FileOperations::OperationStrategy::Native && predicate(*transfer);
        });
    };
    const bool allTransferPlansNative = std::ranges::all_of(plans,
                                                            [](const FileOperations::FileOperationPlan& plan) noexcept
    {
        const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
        return transfer != nullptr && transfer->strategy == FileOperations::OperationStrategy::Native;
    });
    if (allTransferPlansNative)
    {
        return HOST_FILE_OPERATION_VERIFICATION_NOT_APPLICABLE;
    }
    if (transfers([](const FileOperations::TransferPlan& transfer) noexcept { return transfer.destinationEndpoint.verificationCapabilityCheckDeferred; }))
    {
        return HOST_FILE_OPERATION_VERIFICATION_CHECK_DURING_OPERATION;
    }
    return transfers(
               [](const FileOperations::TransferPlan& transfer) noexcept
    {
        return transfer.destinationEndpoint.verificationHostReadback || transfer.destinationEndpoint.verificationProviderBlake3Proof ||
               transfer.destinationEndpoint.verificationWriterDigestProof;
    })
               ? HOST_FILE_OPERATION_VERIFICATION_SUPPORTED
               : HOST_FILE_OPERATION_VERIFICATION_UNSUPPORTED;
}

HRESULT FolderWindow::FileOperationState::StartOperation(OperationAdmission admission, uint64_t* taskIdOut)
{
    const auto uiAdmissionStartedAt = std::chrono::steady_clock::now();
    if (taskIdOut)
    {
        *taskIdOut = 0;
    }
    if (_completionShutdown.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
    }
    // An accepted "Cancel all and exit" refuses new work before any prompt, clipboard use or breadcrumb.
    if (_owner.IsFileOperationsExitPending())
    {
        Debug::Warning(L"File Operations refused a new admission because {} is pending.", _owner.FileOperationsExitPendingReason());
        return FolderWindow::kFileOperationRefusedForPendingExit;
    }

    FileOperations::FileOperationPlanGroup plans = std::move(admission.plans);
    if (plans.empty())
    {
        Debug::Error(L"FolderWindow StartOperation received an empty plan group.");
        return E_INVALIDARG;
    }

    const auto planOptions = [](const FileOperations::FileOperationPlan& plan) noexcept
    { return std::visit([](const auto& typedPlan) noexcept { return typedPlan.options; }, plan); };
    const auto planOperation = [](const FileOperations::FileOperationPlan& plan) noexcept -> FileSystemOperation
    {
        return std::visit(
            [](const auto& typedPlan) noexcept -> FileSystemOperation
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::TransferPlan>)
                return typedPlan.intent == FileOperations::TransferIntent::Move ? FILESYSTEM_MOVE : FILESYSTEM_COPY;
            else if constexpr (std::is_same_v<Plan, FileOperations::DeletePlan>)
                return FILESYSTEM_DELETE;
            else
                return FILESYSTEM_RENAME;
        },
            plan);
    };
    const auto planSourcePluginId = [](const FileOperations::FileOperationPlan& plan) -> std::wstring
    {
        return std::visit(
            [](const auto& typedPlan) -> std::wstring
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::TransferPlan>)
                return typedPlan.sourceEndpoint.pluginId;
            else
                return typedPlan.endpoint.pluginId;
        },
            plan);
    };
    const auto planDestinationPluginId = [](const FileOperations::FileOperationPlan& plan) -> std::wstring
    {
        return std::visit(
            [](const auto& typedPlan) -> std::wstring
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::TransferPlan>)
                return typedPlan.destinationEndpoint.pluginId;
            else
                return {};
        },
            plan);
    };
    const auto appendPlanSourcePaths = [](const FileOperations::FileOperationPlan& plan, std::vector<std::filesystem::path>& paths)
    {
        std::visit(
            [&paths](const auto& typedPlan)
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::RenamePlan>)
            {
                for (const auto& step : typedPlan.finalMappings)
                    paths.emplace_back(step.source.providerPath);
            }
            else
            {
                for (const auto& item : typedPlan.selectedItems)
                    paths.emplace_back(item.providerPath);
            }
        },
            plan);
    };
    const auto planDestinationFolder = [](const FileOperations::FileOperationPlan& plan) -> std::filesystem::path
    {
        return std::visit(
            [](const auto& typedPlan) -> std::filesystem::path
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::TransferPlan>)
                return typedPlan.destination.providerFolderPath;
            else
                return {};
        },
            plan);
    };

    const FileOperations::FileOperationPlan& firstPlan = plans.front();
    FileOperations::OperationOptions options           = planOptions(firstPlan);
    const FileSystemOperation operation                = planOperation(firstPlan);
    const bool archiveDeleteConsentCaptured            = admission.capturedConsentKind == FileOperations::ConsentKind::ArchiveDeleteAfter;
    if (admission.capturedConsentKind.has_value())
    {
        const bool validCapturedConsent = archiveDeleteConsentCaptured && operation == FILESYSTEM_DELETE &&
                                          std::ranges::all_of(plans,
                                                              [](const FileOperations::FileOperationPlan& plan) noexcept
        {
            const auto* deletion = std::get_if<FileOperations::DeletePlan>(&plan);
            return deletion != nullptr && deletion->mode == FileOperations::DeleteMode::Permanent &&
                   (deletion->origin == FileOperations::DeleteOrigin::PackCleanup || deletion->origin == FileOperations::DeleteOrigin::UnpackCleanup);
        });
        if (! validCapturedConsent)
        {
            Debug::Error(L"FolderWindow StartOperation rejected a captured destructive consent outside archive cleanup.");
            return E_INVALIDARG;
        }
    }
    const bool inlineRename = std::visit(
        [](const auto& typedPlan) noexcept
    {
        using Plan = std::remove_cvref_t<decltype(typedPlan)>;
        if constexpr (std::is_same_v<Plan, FileOperations::RenamePlan>)
        {
            return typedPlan.origin == FileOperations::RenameOrigin::InlineRename;
        }
        else
        {
            return false;
        }
    },
        firstPlan);
    const std::wstring sourcePluginId       = planSourcePluginId(firstPlan);
    std::wstring destinationPluginId        = planDestinationPluginId(firstPlan);
    std::filesystem::path destinationFolder = planDestinationFolder(firstPlan);
    const auto planSourceEndpoint           = [](const FileOperations::FileOperationPlan& plan) noexcept -> const FileOperations::QualifiedEndpoint&
    {
        return std::visit(
            [](const auto& typedPlan) noexcept -> const FileOperations::QualifiedEndpoint&
        {
            using Plan = std::remove_cvref_t<decltype(typedPlan)>;
            if constexpr (std::is_same_v<Plan, FileOperations::TransferPlan>)
                return typedPlan.sourceEndpoint;
            else
                return typedPlan.endpoint;
        },
            plan);
    };
    const auto& firstSourceEndpoint                     = planSourceEndpoint(firstPlan);
    const auto* firstTransfer                           = std::get_if<FileOperations::TransferPlan>(&firstPlan);
    const FileSystemPathIdentity* sourceDisplayIdentity = firstSourceEndpoint.pathIdentity.has_value() ? &firstSourceEndpoint.pathIdentity.value() : nullptr;
    const FileSystemPathIdentity* destinationDisplayIdentity =
        firstTransfer && firstTransfer->destinationEndpoint.pathIdentity.has_value() ? &firstTransfer->destinationEndpoint.pathIdentity.value() : nullptr;
    std::vector<std::filesystem::path> sourcePaths;
    std::optional<uint32_t> groupMoveClipboardSequence;
    bool childMissingMoveClipboardSequence = false;
    for (const FileOperations::FileOperationPlan& plan : plans)
    {
        // Display grouping borrows only the already captured endpoint contracts.
        // A multi-root/profile admission cannot lend one child's rules to another.
        if (sourceDisplayIdentity && ! FileOperations::QualifiedEndpointsHaveSamePathContract(firstSourceEndpoint, planSourceEndpoint(plan)))
            sourceDisplayIdentity = nullptr;
        if (destinationDisplayIdentity)
        {
            const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
            if (! transfer || ! FileOperations::QualifiedEndpointsHaveSamePathContract(firstTransfer->destinationEndpoint, transfer->destinationEndpoint))
                destinationDisplayIdentity = nullptr;
        }
        const FileOperations::OperationOptions childOptions = planOptions(plan);
        if (planOperation(plan) != operation || planSourcePluginId(plan) != sourcePluginId || planDestinationPluginId(plan) != destinationPluginId ||
            planDestinationFolder(plan) != destinationFolder || childOptions.linkPolicy != options.linkPolicy ||
            childOptions.verifyAfterCopy != options.verifyAfterCopy || childOptions.executionMode != options.executionMode ||
            childOptions.bandwidthLimitBytesPerSecond != options.bandwidthLimitBytesPerSecond)
        {
            Debug::Error(L"FolderWindow StartOperation rejected an inconsistent child plan group.");
            return E_INVALIDARG;
        }
        if (const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan); transfer && operation == FILESYSTEM_MOVE)
        {
            if (! transfer->moveClipboardSequence.has_value())
            {
                childMissingMoveClipboardSequence = true;
            }
            else if (! groupMoveClipboardSequence.has_value())
            {
                groupMoveClipboardSequence = transfer->moveClipboardSequence->windowsSequenceNumber;
            }
            else if (groupMoveClipboardSequence.value() != transfer->moveClipboardSequence->windowsSequenceNumber)
            {
                Debug::Error(L"FolderWindow StartOperation rejected inconsistent clipboard sequences in one plan group.");
                return E_INVALIDARG;
            }
        }
        appendPlanSourcePaths(plan, sourcePaths);
    }
    if (groupMoveClipboardSequence.has_value() && childMissingMoveClipboardSequence)
    {
        Debug::Error(L"FolderWindow StartOperation rejected a partial clipboard-Move plan group.");
        return E_INVALIDARG;
    }
    if (groupMoveClipboardSequence.has_value() != static_cast<bool>(admission.preWorkerReleaseBarrier))
    {
        Debug::Error(L"FolderWindow StartOperation rejected a clipboard-Move group without its worker-release barrier.");
        return E_INVALIDARG;
    }
    if (groupMoveClipboardSequence.has_value())
    {
        std::scoped_lock lock(_mutex);
        if (std::ranges::find(_acceptedMoveClipboardSequences, groupMoveClipboardSequence.value()) != _acceptedMoveClipboardSequences.end())
        {
            Debug::Warning(L"File Operations rejected an already-consumed clipboard-Move sequence before confirmation.");
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
    }
    FolderWindow::Pane sourcePane                                      = admission.sourcePane;
    std::optional<FolderWindow::Pane> destinationPane                  = admission.destinationPane;
    wil::com_ptr<IFileSystem> fileSystem                               = std::move(admission.fileSystem);
    wil::com_ptr<IFileSystem> destinationFileSystem                    = std::move(admission.destinationFileSystem);
    const FileSystemFlags flags                                        = admission.flags;
    bool waitForOthers                                                 = ! inlineRename && options.executionMode == FileOperations::ExecutionMode::Queue;
    uint64_t initialSpeedLimitBytesPerSecond                           = options.bandwidthLimitBytesPerSecond.value_or(0u);
    const ExecutionMode executionMode                                  = admission.executionMode;
    const bool requireConfirmation                                     = admission.requireConfirmation;
    std::vector<FolderWindow::ResolvedFileOperationItem> resolvedItems = std::move(admission.resolvedItems);
    std::wstring confirmationMessage                                   = std::move(admission.confirmationMessage);
    std::wstring sourcePluginShortId                                   = std::move(admission.sourcePluginShortId);
    std::wstring sourceInstanceContext                                 = std::move(admission.sourceInstanceContext);
    std::wstring destinationPluginShortId                              = std::move(admission.destinationPluginShortId);
    std::wstring destinationInstanceContext                            = std::move(admission.destinationInstanceContext);

    const bool useResolvedItems = ! resolvedItems.empty();
    if (useResolvedItems)
    {
        if (operation != FILESYSTEM_COPY && operation != FILESYSTEM_MOVE)
        {
            Debug::Error(L"FolderWindow StartOperation resolved items require copy/move op={}", static_cast<unsigned int>(operation));
            return E_INVALIDARG;
        }

        if (resolvedItems.size() != sourcePaths.size())
        {
            Debug::Error(L"FolderWindow StartOperation resolved item/source count mismatch resolved={} source={}", resolvedItems.size(), sourcePaths.size());
            return E_INVALIDARG;
        }

        for (size_t index = 0; index < resolvedItems.size(); ++index)
        {
            if (resolvedItems[index].sourcePath.empty() || resolvedItems[index].destinationPath.empty() ||
                resolvedItems[index].sourcePath != sourcePaths[index])
            {
                Debug::Error(L"FolderWindow StartOperation rejected invalid resolved item index={}", index);
                return E_INVALIDARG;
            }
        }

        Debug::Perf::Emit(L"compare.sync.manifest.submitted_items",
                          operation == FILESYSTEM_MOVE ? L"move" : L"copy",
                          0,
                          static_cast<uint64_t>(resolvedItems.size()),
                          0,
                          S_OK);
    }

    if (! fileSystem)
    {
        Debug::Error(L"FolderWindow StartOperation null filesystem");
        return E_POINTER;
    }

    if (sourcePaths.empty())
    {
        Debug::Error(L"FolderWindow StartOperation sourcePath empty");
        return S_FALSE;
    }

    // Publication needs the owner recovery heartbeat. Arm it before any prompt, so a failure is
    // reported before the user consents instead of dropping a confirmed operation. A modal
    // prompt can let the idle heartbeat stop, so publication arms it again below.
    if (const HRESULT preConsentRecoveryHr = ArmCompletionRecoveryTimer(); FAILED(preConsentRecoveryHr))
    {
        Debug::Error(L"File Operations refused admission before any prompt: the recovery heartbeat could not be armed (hr=0x{:08X}).",
                     static_cast<unsigned long>(preConsentRecoveryHr));
        return preConsentRecoveryHr;
    }

    std::optional<FileOperationArtifacts::TouchGuardReceipt> artifactTouchReceipt;
    for (;;)
    {
        std::vector<FileOperationArtifacts::Candidate> artifactCandidates;
        const HRESULT candidatesHr = CollectArtifactTouchCandidates(plans, fileSystem.get(), destinationFileSystem.get(), artifactCandidates);
        if (FAILED(candidatesHr))
        {
            return candidatesHr;
        }
        if (candidatesHr != S_OK)
        {
            break;
        }
        FileOperationArtifacts::TouchGuardRequest touchRequest;
        const HRESULT requestHr = FileOperationArtifacts::BuildTouchGuardRequest(artifactCandidates, touchRequest);
        if (FAILED(requestHr))
        {
            return requestHr;
        }
        if (requestHr != S_OK)
        {
            break;
        }
        FileOperationArtifacts::TouchGuardReceipt accepted;
        const HRESULT confirmHr = ConfirmArtifactTouch(IngressPromptAnchor(admission.sourcePane), touchRequest, accepted, true);
        if (_completionShutdown.load(std::memory_order_acquire))
        {
            return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
        }
        if (confirmHr == HRESULT_FROM_WIN32(ERROR_RETRY))
        {
            // R7.1: an item that could not be opened may open now; capture every candidate again.
            continue;
        }
        if (! touchRequest.allItemsRevalidatable && confirmHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        {
            // The blocked warning already said the operation was not started.
            return FolderWindow::kFileOperationRefusedAlreadyReported;
        }
        if (confirmHr != S_OK)
        {
            return confirmHr;
        }
        // I35 R7.3: a request of handed-off Recycle items only was warned about and has nothing to
        // revalidate; their own route corroborates them.
        if (! accepted.items.empty())
        {
            artifactTouchReceipt = std::move(accepted);
        }
        break;
    }

    if (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE)
    {
        SessionState::UpdateActiveFileSystemPluginIdsAndOperation({sourcePluginId, destinationPluginId}, SessionState::OperationKind::Copy);
    }

    const bool sourceIsLocalFilePlugin  = NavigationLocation::IsFilePluginShortId(sourcePluginShortId);
    const bool deleteBypassesRecycleBin = operation == FILESYSTEM_DELETE && ((flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) == 0 || ! sourceIsLocalFilePlugin);
    const bool supportsBandwidthLimit   = operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE;
    uint64_t taskDesiredSpeedLimit      = supportsBandwidthLimit ? initialSpeedLimitBytesPerSecond : 0u;

    std::vector<DWORD> sourcePathAttributesHint;

    const uint64_t knownCopyOnlyMoveCount   = std::accumulate(plans.begin(),
                                                              plans.end(),
                                                              uint64_t{0u},
                                                              [](const uint64_t current, const FileOperations::FileOperationPlan& plan) noexcept
    {
        const auto* transfer = std::get_if<FileOperations::TransferPlan>(&plan);
        if (transfer == nullptr || transfer->intent != FileOperations::TransferIntent::Move ||
            transfer->strategy != FileOperations::OperationStrategy::CopyOnly)
        {
            return current;
        }
        return SaturatingAdd(current, transfer->selectedItems.size());
    });
    const bool transferConfirmationRequired = requireConfirmation || knownCopyOnlyMoveCount > 0u;
    bool copyOnlyMoveDisclosed              = false;

    if (operation == FILESYSTEM_COPY || operation == FILESYSTEM_MOVE)
    {
        unsigned long long fileCount    = 0;
        unsigned long long folderCount  = 0;
        unsigned long long unknownCount = 0;
        std::filesystem::path sampleFile;
        bool hasSampleFile = false;

        const FolderView& sourceFolderView = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.folderView : _owner._rightPane.folderView;
        const std::vector<FolderView::PathAttributes> selected = sourceFolderView.GetSelectedOrFocusedPathAttributes();
        bool selectionMatches                                  = ! selected.empty() && selected.size() == sourcePaths.size();
        if (selectionMatches)
        {
            for (size_t i = 0; i < selected.size(); ++i)
            {
                if (selected[i].path != sourcePaths[i])
                {
                    selectionMatches = false;
                    break;
                }
            }
        }

        if (selectionMatches)
        {
            for (const auto& item : selected)
            {
                const bool isDirectory = (item.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (isDirectory)
                {
                    ++folderCount;
                    continue;
                }

                ++fileCount;
                if (! hasSampleFile)
                {
                    sampleFile    = item.path;
                    hasSampleFile = true;
                }
            }

            sourcePathAttributesHint.reserve(selected.size());
            for (const auto& item : selected)
            {
                sourcePathAttributesHint.push_back(item.fileAttributes);
            }
        }
        else
        {
            unknownCount = static_cast<unsigned long long>(sourcePaths.size());
        }

        // I35 XP-04: verification availability is a property of the plans, decided once here whether or
        // not the confirmation is shown. A destination writer that proves its own content digest is a
        // verification route; with no route at all the task copies without verification instead of
        // refusing to publish.
        const HostFileOperationVerificationAvailability verificationAvailability = PlannedVerificationAvailability(plans);
        const bool verificationImpossible                                        = verificationAvailability == HOST_FILE_OPERATION_VERIFICATION_UNSUPPORTED ||
                                                                                   verificationAvailability == HOST_FILE_OPERATION_VERIFICATION_NOT_APPLICABLE;
        if (! transferConfirmationRequired && verificationImpossible && options.verifyAfterCopy)
        {
            options.verifyAfterCopy = false;
            for (FileOperations::FileOperationPlan& plan : plans)
            {
                std::visit([](auto& typedPlan) noexcept { typedPlan.options.verifyAfterCopy = false; }, plan);
            }
        }

        // The source attribute hints feed the popup's completed file/folder counters for every
        // transfer, so they are captured whenever the pane selection matches the sources. Only the
        // non-revertable confirmation itself depends on transferConfirmationRequired.
        if (transferConfirmationRequired)
        {
            NonRevertableFileOperationPromptCounts counts{};
            counts.fileCount                 = fileCount;
            counts.folderCount               = folderCount;
            counts.unknownCount              = unknownCount;
            counts.sampleFile                = std::move(sampleFile);
            counts.hasSampleFile             = hasSampleFile;
            const bool consumesClipboardMove = groupMoveClipboardSequence.has_value();
            HostFileOperationPromptOptions promptOptions{};
            promptOptions.sizeBytes = sizeof(promptOptions);
            promptOptions.linkPolicy =
                options.linkPolicy == FileOperations::LinkPolicy::Skip ? HOST_FILE_OPERATION_LINK_SKIP : HOST_FILE_OPERATION_LINK_PRESERVE;
            promptOptions.verifyAfterCopy = options.verifyAfterCopy ? 1u : 0u;
            promptOptions.executionMode =
                options.executionMode == FileOperations::ExecutionMode::Parallel ? HOST_FILE_OPERATION_EXECUTION_PARALLEL : HOST_FILE_OPERATION_EXECUTION_QUEUE;
            promptOptions.clipboardMoveConsumesCutList = consumesClipboardMove ? 1u : 0u;
            promptOptions.bandwidthLimitBytesPerSecond = options.bandwidthLimitBytesPerSecond.value_or(0u);
            promptOptions.verificationAvailability     = verificationAvailability;
            if (verificationImpossible)
            {
                promptOptions.verifyAfterCopy = 0u;
            }

            // I35 R5.6: a known Copy-only Move says so before it starts; the kept sources are then the
            // disclosed outcome, recorded as one Information note.
            const std::wstring copyOnlyNotice =
                knownCopyOnlyMoveCount > 0u
                    ? FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFIRM_COPY_ONLY_MOVE, static_cast<unsigned long>(knownCopyOnlyMoveCount))
                    : std::wstring{};
            const bool confirmed = ConfirmNonRevertableFileOperation(IngressPromptAnchor(sourcePane),
                                                                     operation,
                                                                     sourcePaths,
                                                                     destinationFolder,
                                                                     counts,
                                                                     &promptOptions,
                                                                     confirmationMessage,
                                                                     sourceDisplayIdentity,
                                                                     destinationDisplayIdentity,
                                                                     copyOnlyNotice);
            if (_completionShutdown.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
            }
            if (! confirmed)
            {
                return S_FALSE;
            }
            copyOnlyMoveDisclosed = ! copyOnlyNotice.empty();

            options.linkPolicy =
                promptOptions.linkPolicy == HOST_FILE_OPERATION_LINK_SKIP ? FileOperations::LinkPolicy::Skip : FileOperations::LinkPolicy::Preserve;
            options.verifyAfterCopy = promptOptions.verifyAfterCopy != 0u &&
                                      promptOptions.verificationAvailability != HOST_FILE_OPERATION_VERIFICATION_UNSUPPORTED &&
                                      promptOptions.verificationAvailability != HOST_FILE_OPERATION_VERIFICATION_NOT_APPLICABLE;
            options.executionMode   = promptOptions.executionMode == HOST_FILE_OPERATION_EXECUTION_PARALLEL ? FileOperations::ExecutionMode::Parallel
                                                                                                            : FileOperations::ExecutionMode::Queue;
            options.bandwidthLimitBytesPerSecond =
                promptOptions.bandwidthLimitBytesPerSecond == 0u ? std::nullopt : std::optional<uint64_t>(promptOptions.bandwidthLimitBytesPerSecond);
            for (FileOperations::FileOperationPlan& plan : plans)
            {
                std::visit([&options](auto& typedPlan) noexcept { typedPlan.options = options; }, plan);
            }
            waitForOthers                   = ! inlineRename && options.executionMode == FileOperations::ExecutionMode::Queue;
            initialSpeedLimitBytesPerSecond = options.bandwidthLimitBytesPerSecond.value_or(0u);
            taskDesiredSpeedLimit           = supportsBandwidthLimit ? initialSpeedLimitBytesPerSecond : 0u;
        }
    }
    else if (operation == FILESYSTEM_DELETE && ! archiveDeleteConsentCaptured && (requireConfirmation || deleteBypassesRecycleBin))
    {
        unsigned long long fileCount    = 0;
        unsigned long long folderCount  = 0;
        unsigned long long unknownCount = 0;
        std::filesystem::path sampleFile;
        bool hasSampleFile = false;

        const FolderView& sourceFolderView = sourcePane == FolderWindow::Pane::Left ? _owner._leftPane.folderView : _owner._rightPane.folderView;
        const std::vector<FolderView::PathAttributes> selected = sourceFolderView.GetSelectedOrFocusedPathAttributes();
        bool selectionMatches                                  = ! selected.empty() && selected.size() == sourcePaths.size();
        if (selectionMatches)
        {
            for (size_t i = 0; i < selected.size(); ++i)
            {
                if (selected[i].path != sourcePaths[i])
                {
                    selectionMatches = false;
                    break;
                }
            }
        }

        if (selectionMatches)
        {
            for (const auto& item : selected)
            {
                const bool isDirectory = (item.fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (isDirectory)
                {
                    ++folderCount;
                    continue;
                }

                ++fileCount;
                if (! hasSampleFile)
                {
                    sampleFile    = item.path;
                    hasSampleFile = true;
                }
            }
        }
        else
        {
            unknownCount = static_cast<unsigned long long>(sourcePaths.size());
        }

        auto suffixFor = [](unsigned long long count) noexcept -> std::wstring_view
        { return count == 1ull ? std::wstring_view(L"") : std::wstring_view(L"s"); };

        const unsigned long long itemCount = static_cast<unsigned long long>(sourcePaths.size());
        std::wstring what;
        if (unknownCount > 0)
        {
            const std::wstring_view itemSuffix = suffixFor(itemCount);
            what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_ITEM, itemCount, itemSuffix);
        }
        else if (fileCount > 0 && folderCount > 0)
        {
            const std::wstring_view fileSuffix   = suffixFor(fileCount);
            const std::wstring_view folderSuffix = suffixFor(folderCount);
            what = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILES_FOLDERS, fileCount, fileSuffix, folderCount, folderSuffix);
        }
        else if (fileCount > 0)
        {
            const std::wstring_view fileSuffix = suffixFor(fileCount);
            what                               = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FILE, fileCount, fileSuffix);
        }
        else
        {
            const std::wstring_view folderSuffix = suffixFor(folderCount);
            what                                 = FormatStringResource(nullptr, IDS_FMT_FILEOPS_COUNT_FOLDER, folderCount, folderSuffix);
        }

        NonRevertableFileOperationPromptCounts counts{};
        counts.fileCount      = fileCount;
        counts.folderCount    = folderCount;
        counts.unknownCount   = unknownCount;
        counts.sampleFile     = std::move(sampleFile);
        counts.hasSampleFile  = hasSampleFile;
        std::wstring fromText = FormatFileOperationConfirmationSource(sourcePaths, counts, sourceDisplayIdentity);

        if (deleteBypassesRecycleBin)
        {
            // C1: the permanent-delete confirmation moves to the card, after Preparing has pinned
            // every selected root on the task thread; the words stay the same, gathered here from
            // the pane listing without a provider call.
            admission.permanentDeleteConfirmationPending = true;
            admission.permanentDeleteConsentDetail       = std::move(what);
            admission.permanentDeleteConsentFrom         = std::move(fromText);
#ifdef ENABLE_TESTS
            // Self-tests answered the former modal through the host prompt override, or through the
            // run's auto-accept default when none is set; the card honors what was in effect at
            // admission.
            const HostPromptResult testOverride = HostGetTestPromptResultOverride();
            if (testOverride == HOST_PROMPT_RESULT_OK || (testOverride == HOST_PROMPT_RESULT_NONE && HostGetAutoAcceptPrompts()))
            {
                admission.permanentDeleteTestOverride = 1;
            }
            else if (testOverride == HOST_PROMPT_RESULT_CANCEL)
            {
                admission.permanentDeleteTestOverride = 2;
            }
#endif
        }
        else
        {
            const std::wstring message = FormatStringResource(nullptr, IDS_FMT_FILEOPS_CONFIRM_PERMANENT_DELETE, what, fromText);
            const std::wstring caption = LoadStringResource(nullptr, IDS_FILEOP_OPERATION_DELETE);

            HostPromptRequest prompt{};
            prompt.sizeBytes     = sizeof(prompt);
            prompt.scope         = HOST_ALERT_SCOPE_WINDOW;
            prompt.severity      = HOST_ALERT_WARNING;
            prompt.buttons       = HOST_PROMPT_BUTTONS_OK_CANCEL;
            prompt.targetWindow  = IngressPromptAnchor(sourcePane);
            prompt.title         = caption.c_str();
            prompt.message       = message.c_str();
            prompt.defaultResult = HOST_PROMPT_RESULT_CANCEL;
            prompt.presentation  = HOST_PROMPT_PRESENTATION_DELETE;

            HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
            const HRESULT hrPrompt        = HostShowPrompt(prompt, nullptr, &promptResult);
            if (_completionShutdown.load(std::memory_order_acquire))
            {
                return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
            }
            if (FAILED(hrPrompt) || promptResult != HOST_PROMPT_RESULT_OK)
            {
                return S_FALSE;
            }
        }
    }
    // A prompt pumps messages, so the exit may have been accepted while it was open.
    if (_owner.IsFileOperationsExitPending())
    {
        Debug::Warning(L"File Operations refused an admission after its prompt because {} became pending.", _owner.FileOperationsExitPendingReason());
        return FolderWindow::kFileOperationRefusedForPendingExit;
    }

    uint64_t admittedTaskId = 0u;
    {
        std::scoped_lock lock(_mutex);
        admittedTaskId = _nextTaskId++;
    }

    for (FileOperations::FileOperationPlan& plan : plans)
    {
        if (auto* deletion = std::get_if<FileOperations::DeletePlan>(&plan); deletion && deletion->mode == FileOperations::DeleteMode::Permanent &&
                                                                             ! deletion->initialConsent.has_value() &&
                                                                             ! admission.permanentDeleteConfirmationPending)
        {
            const FileOperations::ConsentKind consentKind = admission.capturedConsentKind.value_or(
                deletion->origin == FileOperations::DeleteOrigin::PackCleanup || deletion->origin == FileOperations::DeleteOrigin::UnpackCleanup
                    ? FileOperations::ConsentKind::ArchiveDeleteAfter
                    : FileOperations::ConsentKind::PermanentDelete);
            deletion->initialConsent = FileOperations::DestructiveConsentReceipt{
                .kind      = consentKind,
                .taskNonce = admittedTaskId,
            };
        }

        FileOperations::PlanRejectionBucket finalRejection = FileOperations::PlanRejectionBucket::None;
        const HRESULT finalValidationHr                    = FileOperations::ValidatePlan(plan, &finalRejection, admission.permanentDeleteConfirmationPending);
        if (FAILED(finalValidationHr))
        {
            Debug::Error(L"File Operations rejected a child plan after confirmation but before publication (bucket={}, hr=0x{:08X}).",
                         static_cast<unsigned int>(finalRejection),
                         static_cast<unsigned long>(finalValidationHr));
            return finalValidationHr;
        }
    }

    const ULONGLONG presentationAdmissionTick = TaskPresentationNowTick();
    auto task                                 = std::make_unique<Task>(*this);
    AdmissionDestinationInputs initialDestination;
    {
        std::scoped_lock lock(_mutex);
        task->_taskId = admittedTaskId;
        task->_admissionKey.store(admittedTaskId, std::memory_order_release);
        task->_queueOrderKey.store(task->_taskId, std::memory_order_release);
        task->_operation                   = operation;
        task->_executionMode               = executionMode;
        task->_sourcePane                  = sourcePane;
        initialDestination.pane            = destinationPane;
        task->_sourcePluginId              = sourcePluginId;
        task->_sourcePluginShortId         = std::move(sourcePluginShortId);
        task->_sourceInstanceContext       = std::move(sourceInstanceContext);
        initialDestination.pluginId        = std::move(destinationPluginId);
        initialDestination.pluginShortId   = std::move(destinationPluginShortId);
        initialDestination.instanceContext = std::move(destinationInstanceContext);
        task->_fileSystem                  = fileSystem;
        initialDestination.fileSystem      = std::move(destinationFileSystem);
        task->_sourcePaths                 = std::move(sourcePaths);
        task->_sourcePathAttributesHint    = std::move(sourcePathAttributesHint);
        task->_selection.itemCount         = task->_sourcePaths.size();
        task->_selection.kindsKnown        = task->_sourcePathAttributesHint.size() == task->_sourcePaths.size();
        for (const DWORD attributes : task->_sourcePathAttributesHint)
        {
            if (attributes == 0u || attributes == INVALID_FILE_ATTRIBUTES)
                task->_selection.kindsKnown = false;
            else if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
                ++task->_selection.folderCount;
            else
                ++task->_selection.fileCount;
        }
        if (task->_sourcePaths.size() == 1u)
        {
            const auto& path = task->_sourcePaths.front().native();
            if (! sourceDisplayIdentity || ! TryGetFileSystemLeafName(*sourceDisplayIdentity, path, task->_selection.singleName))
                task->_selection.singleName = path; // Exact provider path when its captured profile cannot derive a leaf.
        }
        initialDestination.folder                 = std::move(destinationFolder);
        task->_flags                              = flags;
        initialDestination.resolvedItems          = std::move(resolvedItems);
        task->_permanentDeleteConfirmationPending = admission.permanentDeleteConfirmationPending;
        task->_permanentDeleteConsentDetail       = std::move(admission.permanentDeleteConsentDetail);
        task->_permanentDeleteConsentFrom         = std::move(admission.permanentDeleteConsentFrom);
        task->_permanentDeleteTestOverride        = admission.permanentDeleteTestOverride;
        task->_crossFsBridgeBufferBytes           = GetCrossFsBridgeBufferBytesFromSettings(_owner._settings);
        task->_waitForOthers.store(waitForOthers, std::memory_order_release);
        task->_desiredSpeedLimitBytesPerSecond.store(taskDesiredSpeedLimit, std::memory_order_release);
        task->_clipboardMoveAdmission     = groupMoveClipboardSequence.has_value();
        task->_copyOnlyMoveDisclosed      = copyOnlyMoveDisclosed;
        task->_preWorkerReleaseBarrier    = std::move(admission.preWorkerReleaseBarrier);
        task->_preConsumptionDecisionGate = std::move(admission.preConsumptionDecisionGate);
        task->_preparationObserver        = std::move(admission.preparationObserver);
        task->_externalProgressCallback   = std::move(admission.progressCallback);
        task->InitializeAdmissionInputs(
            std::make_shared<const FileOperations::FileOperationPlanGroup>(std::move(plans)), std::move(initialDestination), std::move(artifactTouchReceipt));
        task->_presentationState.store(Task::TaskPresentationState::Hidden, std::memory_order_release);
        task->_presentationDeadlineTick       = presentationAdmissionTick > std::numeric_limits<ULONGLONG>::max() - FileOperations::kTaskCardRevealDelayMs
                                                    ? std::numeric_limits<ULONGLONG>::max()
                                                    : presentationAdmissionTick + FileOperations::kTaskCardRevealDelayMs;
        task->_suppressCleanCompletionSummary = inlineRename;
        // Mark as waiting in queue immediately if queuing, so UI shows "Waiting..." right away
        task->SetWaitingInQueue(waitForOthers);
    }

    // Allocate every selected-root result builder before the task becomes observable in
    // _tasks or a worker can publish callback truth.
    task->InitializeSourceItemResultBuilders();

    {
        const size_t itemCount = task->_sourcePaths.size();
        {
            std::scoped_lock lock(task->_topLevelCompletionMutex);
            task->_topLevelItemKinds.assign(itemCount, Task::TopLevelItemKind::Unknown);
            task->_topLevelItemCompleted.assign(itemCount, 0);
            task->_plannedTopLevelFiles     = 0;
            task->_plannedTopLevelFolders   = 0;
            task->_completedTopLevelFiles   = 0;
            task->_completedTopLevelFolders = 0;
        }
        task->_publishedCompletedTopLevelFiles.store(0, std::memory_order_release);
        task->_publishedCompletedTopLevelFolders.store(0, std::memory_order_release);

        const bool haveHint = task->_sourcePathAttributesHint.size() == itemCount;
        if (haveHint)
        {
            std::scoped_lock lock(task->_topLevelCompletionMutex);
            for (size_t i = 0; i < itemCount; ++i)
            {
                const bool isDir = (task->_sourcePathAttributesHint[i] & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (isDir)
                {
                    task->_topLevelItemKinds[i] = Task::TopLevelItemKind::Folder;
                    if (task->_plannedTopLevelFolders < std::numeric_limits<unsigned long>::max())
                    {
                        ++task->_plannedTopLevelFolders;
                    }
                }
                else
                {
                    task->_topLevelItemKinds[i] = Task::TopLevelItemKind::File;
                    if (task->_plannedTopLevelFiles < std::numeric_limits<unsigned long>::max())
                    {
                        ++task->_plannedTopLevelFiles;
                    }
                }
            }
        }
    }

    {
        std::scoped_lock lock(task->_progressPathMutex);
        if (! task->_sourcePaths.empty())
        {
            task->_progressSourcePath = task->_sourcePaths.front().native();
        }

        const auto selected = task->LoadSelectedAdmission();
        if (! selected->destination->folder.empty())
        {
            task->_progressDestinationPath = selected->destination->folder.native();
        }

        PublishDiagnosticPathSnapshotLocked(*task);
    }

    Task* rawTask                = task.get();
    const uint64_t startedTaskId = rawTask->_taskId;

    // Admission seeds recovery synchronously before any worker can mutate. A full
    // message queue must not prevent recovery, even before a popup exists.
    const HRESULT recoveryHr = ArmCompletionRecoveryTimer();
    if (FAILED(recoveryHr))
        return recoveryHr;

    {
        std::scoped_lock lock(_mutex);
        if (groupMoveClipboardSequence.has_value())
        {
            if (std::ranges::find(_acceptedMoveClipboardSequences, groupMoveClipboardSequence.value()) != _acceptedMoveClipboardSequences.end())
            {
                Debug::Warning(L"File Operations rejected a duplicate clipboard-Move sequence after an earlier accepted admission.");
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }
            _acceptedMoveClipboardSequences.push_back(groupMoveClipboardSequence.value());
            while (_acceptedMoveClipboardSequences.size() > kMaxAcceptedMoveClipboardSequences)
            {
                _acceptedMoveClipboardSequences.pop_front();
            }
        }
        _tasks.emplace_back(std::move(task));
    }

    // Queue admission is complete only after the worker exists. For clipboard Move, the worker
    // first publishes selected-root/interlock readiness, then remains blocked until the exact
    // accepted sequence has either been consumed successfully or failed terminally.
    try
    {
        rawTask->_thread = std::jthread([rawTask](std::stop_token stopToken) noexcept { rawTask->ThreadMain(stopToken); });
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::system_error& error)
    {
        // std::jthread reports OS thread-admission failure through std::system_error. Roll back
        // the unpublished worker and sequence receipt; the clipboard has not been touched yet.
        Debug::Error(L"File Operations could not create the admitted worker (taskId={}, code={}).", startedTaskId, error.code().value());
        std::scoped_lock lock(_mutex);
        const auto taskIt = std::ranges::find_if(_tasks, [rawTask](const auto& candidate) noexcept { return candidate.get() == rawTask; });
        if (taskIt != _tasks.end())
        {
            _tasks.erase(taskIt);
        }
        if (groupMoveClipboardSequence.has_value())
        {
            const auto sequenceIt = std::ranges::find(_acceptedMoveClipboardSequences, groupMoveClipboardSequence.value());
            if (sequenceIt != _acceptedMoveClipboardSequences.end())
            {
                _acceptedMoveClipboardSequences.erase(sequenceIt);
            }
        }
        return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY);
    }

    RequestTaskPresentationRefresh();

    if (! rawTask->_clipboardMoveAdmission)
    {
        rawTask->_workerReleased.store(true, std::memory_order_release);
        rawTask->_workerReleased.notify_all();
    }
    if (taskIdOut)
    {
        *taskIdOut = startedTaskId;
    }
    Debug::Perf::Emit(L"FileOps.Preparing.UiAdmissionUs",
                      OperationToString(operation),
                      Debug::Perf::ElapsedUs(uiAdmissionStartedAt),
                      static_cast<uint64_t>(rawTask->_sourcePaths.size()),
                      rawTask->_clipboardMoveAdmission ? 1u : 0u,
                      S_OK);
    return S_OK;
}

void FolderWindow::FileOperationState::Task::CompleteClipboardMoveAdmission(const HRESULT status, const bool consumed) noexcept
{
    bool expected = false;
    if (! _clipboardMoveBarrierCompleted.compare_exchange_strong(expected, true, std::memory_order_acq_rel))
    {
        return;
    }

    _clipboardMoveConsumptionStatus.store(status, std::memory_order_release);
    _clipboardMoveConsumed.store(consumed && SUCCEEDED(status), std::memory_order_release);
    if (FAILED(status))
    {
        Debug::Warning(L"File Operations could not clear the accepted clipboard-Move cut list; the accepted task will terminate before mutation (hr=0x{:08X}).",
                       static_cast<unsigned long>(status));
        LogDiagnostic(DiagnosticSeverity::Error, status, L"clipboard.move.consumeFailed", LocalizedDiagnosticText{IDS_FILEOPS_CLIPBOARD_MOVE_CLEAR_FAILED});
    }
    else
    {
        LogDiagnostic(DiagnosticSeverity::Info, S_OK, L"clipboard.move.consumed", LocalizedDiagnosticText{IDS_FILEOPS_CLIPBOARD_MOVE_QUEUED});
    }

    _workerReleased.store(true, std::memory_order_release);
    _workerReleased.notify_all();
}

bool FolderWindow::FileOperationState::IsPendingConflictInspectionCurrent(
    uint64_t taskId, uint64_t promptGeneration, const std::shared_ptr<const FileOperations::ConflictInspectionContext>& context) noexcept
{
    // Task removal belongs to the owning UI thread. Do not retain a borrowed Task
    // across a pane's asynchronous navigation; resolve its ID again on completion.
    if (! context || promptGeneration == 0u || GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) != GetCurrentThreadId())
        return false;
    Task* const task = FindTask(taskId);
    if (! task)
        return false;
    std::scoped_lock lock(task->_conflictArbiter.mutex);
    // An accepted decision answers the prompt even before the worker retires it.
    return ! task->_cancelled.load(std::memory_order_acquire) && ! task->_taskFinished.load(std::memory_order_acquire) && ! task->_stopToken.stop_requested() &&
           task->_conflictArbiter.prompt.active && task->_conflictArbiter.prompt.generation == promptGeneration &&
           task->_conflictArbiter.prompt.pathDisplay.inspection == context && ! task->_conflictArbiter.decisionAction.has_value();
}

HRESULT FolderWindow::FileOperationState::ResolvePendingConflictInspectionTarget(uint64_t taskId,
                                                                                 uint64_t promptGeneration,
                                                                                 std::shared_ptr<const FileOperations::ConflictInspectionContext> context,
                                                                                 bool destination,
                                                                                 FileOperations::ConflictInspectionTarget& out) noexcept
{
    out = {};
    if (! IsPendingConflictInspectionCurrent(taskId, promptGeneration, context))
        return E_ABORT;
    std::weak_ptr<void> uiLifetime;
    {
        std::scoped_lock lock(_mutex);
        uiLifetime = _uiLifetime;
    }
    if (uiLifetime.expired())
        return E_ABORT;
    const auto& location = destination ? context->destination : context->source;
    if (! location)
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    Task* const task = FindTask(taskId);
    if (! task)
        return E_ABORT;
    const auto selectedAdmission = task->LoadSelectedAdmission();
    FileOperations::ConflictInspectionTarget candidate;
    candidate.context          = context;
    candidate.promptGeneration = promptGeneration;
    candidate.destination      = destination;
    candidate.fileSystem       = destination && selectedAdmission->destination->fileSystem ? selectedAdmission->destination->fileSystem : task->_fileSystem;
    candidate.pluginId         = destination ? selectedAdmission->destination->pluginId : task->_sourcePluginId;
    candidate.pluginShortId    = destination ? selectedAdmission->destination->pluginShortId : task->_sourcePluginShortId;
    candidate.instanceContext  = destination ? selectedAdmission->destination->instanceContext : task->_sourceInstanceContext;
    {
        std::scoped_lock lock(task->_conflictArbiter.mutex);
        const auto& prompt   = task->_conflictArbiter.prompt;
        const auto& metadata = destination ? prompt.destinationMetadata : prompt.sourceMetadata;
        if (prompt.generation == promptGeneration && ! prompt.metadataLoading && metadata.available && metadata.objectKindKnown)
            candidate.itemIsDirectory = metadata.isDirectory;
    }
    // Admission fields are immutable. Retain all needed ownership before entering
    // provider code; no borrowed task survives a potentially reentrant COM call.
    const HRESULT status = FileOperations::QualifyConflictInspectionLocation(
        location.value(), context->operation, candidate.fileSystem.get(), candidate.pluginId, candidate.instanceContext);
    if (uiLifetime.expired() || ! IsPendingConflictInspectionCurrent(taskId, promptGeneration, context))
        return E_ABORT;
    if (FAILED(status))
        return status;
    out = std::move(candidate);
    return S_OK;
}

HRESULT FolderWindow::FileOperationState::BeginPendingConflictInspection(uint64_t taskId,
                                                                         uint64_t promptGeneration,
                                                                         std::shared_ptr<const FileOperations::ConflictInspectionContext> context,
                                                                         bool destination,
                                                                         FileOperations::ConflictInspectionCompletion completion) noexcept
{
    if (GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) != GetCurrentThreadId())
        return RPC_E_WRONG_THREAD;
    if (! completion)
        return E_INVALIDARG;
    if (_pendingConflictInspection)
        return HRESULT_FROM_WIN32(ERROR_BUSY);
    const auto pane       = _owner.GetActivePane();
    const HWND foreground = GetForegroundWindow();
    const HWND focus      = GetFocus();
    LASTINPUTINFO input{sizeof(LASTINPUTINFO)};
    const bool inputKnown = GetLastInputInfo(&input) != FALSE;
    std::weak_ptr<void> lifetime;
    {
        std::scoped_lock lock(_mutex);
        lifetime = _uiLifetime;
    }
    FileOperations::ConflictInspectionTarget target;
    const HRESULT qualified = ResolvePendingConflictInspectionTarget(taskId, promptGeneration, std::move(context), destination, target);
    if (lifetime.expired())
        return E_ABORT;
    if (FAILED(qualified))
        return qualified;
    if (target.pluginShortId.empty())
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    if (_pendingConflictInspection || _owner.GetActivePane() != pane)
        return E_ABORT;
    auto& state         = pane == FolderWindow::Pane::Left ? _owner._leftPane : _owner._rightPane;
    const auto previous = state.folderView.GetFolderPath();
    if (! previous || state.pluginId.empty() || state.pluginShortId.empty())
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    auto pending                     = std::make_unique<PendingConflictInspection>();
    const auto generation            = ++_inspectionRequestGeneration;
    pending->generation              = generation;
    pending->taskId                  = taskId;
    pending->pane                    = pane;
    pending->target                  = std::move(target);
    pending->completion              = std::move(completion);
    pending->previousQualifiedPath   = NavigationLocation::FormatHistoryPath(state.pluginShortId, state.instanceContext, previous.value());
    pending->previousProviderPath    = previous.value();
    pending->previousPluginId        = state.pluginId;
    pending->previousInstanceContext = state.instanceContext;
    pending->foreground              = foreground;
    pending->focus                   = focus;
    pending->inputTick               = input.dwTime;
    pending->inputKnown              = inputKnown;
    const auto& location             = destination ? pending->target.context->destination.value() : pending->target.context->source.value();
    const auto path =
        NavigationLocation::FormatHistoryPath(pending->target.pluginShortId, pending->target.instanceContext, std::filesystem::path(location.parentPath));
    _completedConflictInspection.reset();
    _pendingConflictInspection = std::move(pending);
    // I35 R8: the user inspects this prompt's item in the pane from the card; the mirror stays off it.
    NoteDecisionMirrorInspection(taskId, promptGeneration);
    _owner.InstallFolderViewEnumerationCompletedCallback(pane);
    _owner.SetFolderPath(pane, path);
    if (lifetime.expired())
        return E_ABORT;
    if (! _pendingConflictInspection || _pendingConflictInspection->generation != generation)
        return E_ABORT;
    _pendingConflictInspection->visit = state.folderView.CaptureInspectionVisit();
    if (! _pendingConflictInspection->visit)
    {
        FinishPendingConflictInspection({FileOperations::ConflictInspectionNavigationStatus::Unavailable, E_ABORT});
        return S_OK;
    }
    TryCompletePendingConflictInspection(); // Also covers an already-enumerated same-folder request.
    return S_OK;
}

FolderView::EnumerationSettledCallback FolderWindow::FileOperationState::GetConflictInspectionEnumerationCallback(FolderWindow::Pane pane)
{
    if (! _pendingConflictInspection || _pendingConflictInspection->pane != pane)
        return {};
    const std::weak_ptr<void> lifetime = _uiLifetime;
    return [this, lifetime, pane](const FolderView::InspectionVisit& settled, const std::filesystem::path&, HRESULT status)
    {
        if (lifetime.expired())
            return;
        // A watcher or cache refresh (or a user cancel) of the target folder supersedes the
        // requested enumeration without changing the navigation: follow its settlement. No visit
        // is engaged while Begin or a rollback is still dispatching its navigation.
        if (_pendingConflictInspection && _pendingConflictInspection->pane == pane && _pendingConflictInspection->visit &&
            FolderView::IsSameNavigationVisit(_pendingConflictInspection->visit.value(), settled))
            _pendingConflictInspection->visit = settled;
        TryCompletePendingConflictInspection(status);
    };
}

void FolderWindow::FileOperationState::FinishPendingConflictInspection(FileOperations::ConflictInspectionNavigationResult result) noexcept
{
    auto completed = std::move(_pendingConflictInspection);
    if (! completed)
        return;
    _owner.InstallFolderViewEnumerationCompletedCallback(completed->pane);
    Debug::Perf::Emit(L"FileOps.Inspection.NavigationUs",
                      completed->target.destination ? L"destination" : L"source",
                      Debug::Perf::ElapsedUs(completed->startedAt),
                      static_cast<uint64_t>(result.status),
                      result.focusUnchanged ? 1u : 0u,
                      result.error);
    auto completion = std::move(completed->completion);
    if (result.status == FileOperations::ConflictInspectionNavigationStatus::Selected && result.focusUnchanged && completed->visit)
    {
        result.focusTicket           = completed->generation;
        _completedConflictInspection = std::move(completed);
    }
    // No access to this after the caller callback: it may close the owner or begin another request.
    completion(result);
}

void FolderWindow::FileOperationState::CancelPendingConflictInspection() noexcept
{
    _completedConflictInspection.reset();
    FinishPendingConflictInspection({FileOperations::ConflictInspectionNavigationStatus::Stale, E_ABORT});
}

void FolderWindow::FileOperationState::CancelConflictInspectionForTask(uint64_t taskId) noexcept
{
    // FinishPendingConflictInspection reinstalls pane listeners and runs the popup completion:
    // UI thread only, with no FileOperationState or arbiter lock held by the caller.
    if (GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) != GetCurrentThreadId())
        return;
    if (_completedConflictInspection && _completedConflictInspection->taskId == taskId)
        _completedConflictInspection.reset();
    if (_pendingConflictInspection && _pendingConflictInspection->taskId == taskId)
        FinishPendingConflictInspection({FileOperations::ConflictInspectionNavigationStatus::Stale, E_ABORT});
}

void FolderWindow::FileOperationState::CancelConflictInspectionForPane(FolderWindow::Pane pane) noexcept
{
    if (GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) != GetCurrentThreadId())
        return;
    if (_pendingConflictInspection && _pendingConflictInspection->pane == pane)
        FinishPendingConflictInspection({FileOperations::ConflictInspectionNavigationStatus::Stale, E_ABORT});
}

bool FolderWindow::FileOperationState::CommitPendingConflictInspectionFocus(uint64_t ticket, HWND popup) noexcept
{
    if (GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) != GetCurrentThreadId() || ! ticket || ! _completedConflictInspection ||
        _completedConflictInspection->generation != ticket)
        return false;
    auto completed                     = std::move(_completedConflictInspection); // Consume even when a final witness fails.
    const std::weak_ptr<void> lifetime = _uiLifetime;
    const HWND ownerWindow             = GetAncestor(_owner.GetHwnd(), GA_ROOT);
    const auto current                 = [&]() noexcept
    {
        if (lifetime.expired())
            return false;
        const auto& pane = completed->pane == FolderWindow::Pane::Left ? _owner._leftPane : _owner._rightPane;
        LASTINPUTINFO input{sizeof(LASTINPUTINFO)};
        return _inspectionRequestGeneration == ticket && ! _pendingConflictInspection && _popup.get() == popup && IsWindowVisible(popup) &&
               _owner.GetActivePane() == completed->pane && completed->visit && pane.folderView.IsInspectionVisitCurrent(completed->visit.value()) &&
               pane.fileSystem.get() == completed->selectedFileSystem.get() && pane.pluginId == completed->selectedPluginId &&
               pane.instanceContext == completed->selectedInstanceContext && completed->inputKnown && GetLastInputInfo(&input) &&
               input.dwTime == completed->inputTick &&
               IsPendingConflictInspectionCurrent(completed->taskId, completed->target.promptGeneration, completed->target.context);
    };
    if (! current() || completed->foreground != popup || GetForegroundWindow() != popup || GetFocus() != completed->focus ||
        (completed->focus != popup && ! IsChild(popup, completed->focus)))
        return false;
    const auto startedAt = std::chrono::steady_clock::now();
    // Activation and focus can synchronously destroy or replace the owner/prompt.
    // Revalidate after each Windows call, with lifetime checked before member access.
    if (! Common::RequestForegroundWindow(ownerWindow) || ! current() || GetForegroundWindow() != ownerWindow)
        return false;
    _owner.FocusPaneFolderView(completed->pane);
    const bool committed = current() && GetForegroundWindow() == ownerWindow && GetFocus() == _owner.GetFolderViewHwnd(completed->pane);
    Debug::Perf::Emit(
        L"FileOps.Inspection.FocusCommitUs", L"guarded one-shot", Debug::Perf::ElapsedUs(startedAt), ticket, committed ? 1u : 0u, committed ? S_OK : E_ABORT);
    return committed;
}

void FolderWindow::FileOperationState::TryCompletePendingConflictInspection(std::optional<HRESULT> settlement) noexcept
{
    using Status    = FileOperations::ConflictInspectionNavigationStatus;
    using Selection = FolderView::InspectionSelectionStatus;
    if (! _pendingConflictInspection || ! _pendingConflictInspection->visit)
        return; // Dispatch is still installing the requested visit.
    const auto generation              = _pendingConflictInspection->generation;
    const std::weak_ptr<void> lifetime = _uiLifetime;
    const auto pane                    = _pendingConflictInspection->pane;
    auto& state                        = pane == FolderWindow::Pane::Left ? _owner._leftPane : _owner._rightPane;
    const auto visit                   = _pendingConflictInspection->visit.value();
    if (! state.folderView.IsInspectionVisitCurrent(visit))
    {
        FinishPendingConflictInspection({Status::Stale, E_ABORT});
        return;
    }
    if (_pendingConflictInspection->rollingBack)
    {
        // The settlement callback supplies readiness; this branch is never called
        // synchronously after dispatching the rollback.
        const auto current = state.folderView.GetFolderPath();
        if (! settlement)
            return;
        const bool restored = SUCCEEDED(settlement.value()) && current && current.value() == _pendingConflictInspection->previousProviderPath &&
                              state.pluginId == _pendingConflictInspection->previousPluginId &&
                              state.instanceContext == _pendingConflictInspection->previousInstanceContext;
        const bool canceled = _pendingConflictInspection->navigationError == HRESULT_FROM_WIN32(ERROR_CANCELLED);
        FinishPendingConflictInspection({restored ? (canceled ? Status::Canceled : Status::NavigationFailed) : Status::RollbackFailed,
                                         FAILED(settlement.value()) ? settlement.value() : _pendingConflictInspection->navigationError});
        return;
    }
    FileOperations::ConflictInspectionTarget target;
    const HRESULT qualified = ResolvePendingConflictInspectionTarget(_pendingConflictInspection->taskId,
                                                                     _pendingConflictInspection->target.promptGeneration,
                                                                     _pendingConflictInspection->target.context,
                                                                     _pendingConflictInspection->target.destination,
                                                                     target);
    if (lifetime.expired() || ! _pendingConflictInspection || _pendingConflictInspection->generation != generation)
        return;
    if (FAILED(qualified))
    {
        FinishPendingConflictInspection({qualified == E_ABORT ? Status::Stale : Status::Unavailable, qualified});
        return;
    }
    const auto& location = target.destination ? target.context->destination.value() : target.context->source.value();
    // Retain the live pane adapter and its strings across reentrant provider calls.
    const auto liveFileSystem              = state.fileSystem;
    const std::wstring livePluginId        = state.pluginId;
    const std::wstring liveInstanceContext = state.instanceContext;
    const HRESULT live =
        FileOperations::QualifyConflictInspectionLocation(location, target.context->operation, liveFileSystem.get(), livePluginId, liveInstanceContext);
    if (lifetime.expired() || ! _pendingConflictInspection || _pendingConflictInspection->generation != generation)
        return;
    if (state.fileSystem.get() != liveFileSystem.get() || state.pluginId != livePluginId || state.instanceContext != liveInstanceContext)
    {
        FinishPendingConflictInspection({Status::Stale, E_ABORT});
        return;
    }
    if (FAILED(live))
    {
        FinishPendingConflictInspection({Status::Unavailable, live});
        return;
    }
    if (! IsPendingConflictInspectionCurrent(_pendingConflictInspection->taskId, target.promptGeneration, target.context))
    {
        FinishPendingConflictInspection({Status::Stale, E_ABORT});
        return;
    }
    const auto selected =
        state.folderView.SelectItemForInspection(visit, location.endpoint.pathIdentity.value(), location.providerPath, target.itemIsDirectory);
    if (lifetime.expired() || ! _pendingConflictInspection || _pendingConflictInspection->generation != generation)
        return;
    if (! IsPendingConflictInspectionCurrent(_pendingConflictInspection->taskId, target.promptGeneration, target.context))
    {
        FinishPendingConflictInspection({Status::Stale, E_ABORT});
        return;
    }
    switch (selected.status)
    {
        case Selection::Pending: return;
        case Selection::EnumerationFailed:
        {
            _pendingConflictInspection->rollingBack     = true;
            _pendingConflictInspection->navigationError = selected.enumerationStatus;
            const auto previous                         = _pendingConflictInspection->previousQualifiedPath;
            _pendingConflictInspection->visit.reset();
            _owner.SetFolderPath(pane, previous);
            if (lifetime.expired() || ! _pendingConflictInspection || _pendingConflictInspection->generation != generation)
                return;
            _pendingConflictInspection->visit = state.folderView.CaptureInspectionVisit();
            if (! _pendingConflictInspection->visit)
                FinishPendingConflictInspection({Status::RollbackFailed, selected.enumerationStatus});
            return;
        }
        case Selection::Selected:
        {
            LASTINPUTINFO input{sizeof(LASTINPUTINFO)};
            const bool focusUnchanged = _pendingConflictInspection->inputKnown && GetLastInputInfo(&input) &&
                                        input.dwTime == _pendingConflictInspection->inputTick &&
                                        GetForegroundWindow() == _pendingConflictInspection->foreground && GetFocus() == _pendingConflictInspection->focus &&
                                        _owner.GetActivePane() == pane;
            // Selection changes its ownership epoch. Retain a fresh visit and the
            // qualified live adapter for the separate synchronous focus commit.
            _pendingConflictInspection->visit                   = state.folderView.CaptureInspectionVisit();
            _pendingConflictInspection->selectedFileSystem      = liveFileSystem;
            _pendingConflictInspection->selectedPluginId        = livePluginId;
            _pendingConflictInspection->selectedInstanceContext = liveInstanceContext;
            FinishPendingConflictInspection({Status::Selected, S_OK, focusUnchanged});
            return;
        }
        case Selection::MissingOrFiltered: FinishPendingConflictInspection({Status::MissingOrFiltered, HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)}); return;
        case Selection::Ambiguous: FinishPendingConflictInspection({Status::Ambiguous, HRESULT_FROM_WIN32(ERROR_DUP_NAME)}); return;
        case Selection::Stale: FinishPendingConflictInspection({Status::Stale, E_ABORT}); return;
        case Selection::Unavailable: FinishPendingConflictInspection({Status::Unavailable, HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)}); return;
    }
}

void FolderWindow::FileOperationState::CompleteClipboardMoveAdmissionByTaskId(const uint64_t taskId, const HRESULT status) noexcept
{
    if (Task* const task = FindTask(taskId))
    {
        task->CompleteClipboardMoveAdmission(status, false);
    }
}

void FolderWindow::FileOperationState::OnClipboardMoveReady(std::unique_ptr<ClipboardMoveReadyPayload> payload) noexcept
{
    if (! payload)
    {
        return;
    }

    const uint64_t taskId = payload->taskId;
    Task* task            = FindTask(taskId);
    if (! task)
    {
        return;
    }

    const HRESULT readinessHr = task->_selectedRootReadinessStatus.load(std::memory_order_acquire);
    if (_completionShutdown.load(std::memory_order_acquire) || task->_cancelled.load(std::memory_order_acquire))
    {
        task->CompleteClipboardMoveAdmission(HRESULT_FROM_WIN32(ERROR_CANCELLED), false);
        return;
    }
    if (FAILED(readinessHr))
    {
        task->CompleteClipboardMoveAdmission(readinessHr, false);
        return;
    }

    // Clipboard/OLE calls may pump messages. Move the one-shot callback out, invoke it without
    // retaining a raw Task pointer, then re-find by id before publishing the result.
    std::function<HRESULT()> barrier = std::move(task->_preWorkerReleaseBarrier);
    if (! barrier)
    {
        task->CompleteClipboardMoveAdmission(E_UNEXPECTED, false);
        return;
    }
    const HRESULT barrierHr = barrier();

    task = FindTask(taskId);
    if (! task)
    {
        return;
    }
    if (_completionShutdown.load(std::memory_order_acquire) || task->_cancelled.load(std::memory_order_acquire))
    {
        task->CompleteClipboardMoveAdmission(HRESULT_FROM_WIN32(ERROR_CANCELLED), false);
        return;
    }
    const HRESULT terminalBarrierHr = barrierHr == S_OK ? S_OK : (FAILED(barrierHr) ? barrierHr : HRESULT_FROM_WIN32(ERROR_CANCELLED));
    task->CompleteClipboardMoveAdmission(terminalBarrierHr, barrierHr == S_OK);
}

HRESULT FolderWindow::FileOperationState::Task::PrepareBatchRenameArtifactGuard() noexcept
{
    const auto selectedAdmission = LoadSelectedAdmission();
    const auto plans             = selectedAdmission->plans;
    if (! plans || plans->size() != 1u || ! _fileSystem || ! _state)
        return E_UNEXPECTED;
    std::optional<FileOperationArtifacts::TouchGuardReceipt> prepared;
    const HRESULT hr = PrepareArtifactTouchGuard(*plans, _fileSystem.get(), selectedAdmission->destination->fileSystem.get(), prepared);
    if (hr == S_OK)
        StoreInitialArtifactTouchReceipt(std::move(prepared));
    return hr;
}

HRESULT FolderWindow::FileOperationState::Task::PrepareArtifactTouchGuard(const FileOperations::FileOperationPlanGroup& plans,
                                                                          IFileSystem* const sourceFileSystem,
                                                                          IFileSystem* const destinationFileSystem,
                                                                          std::optional<FileOperationArtifacts::TouchGuardReceipt>& out) noexcept
{
    if (! _state || ! sourceFileSystem)
        return E_POINTER;
    const auto stopped = [this]() noexcept
    { return _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested() || _state->_completionShutdown.load(std::memory_order_acquire); };
    if (stopped())
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);

    std::vector<FileOperationArtifacts::Candidate> candidates;
    const HRESULT candidatesHr = CollectArtifactTouchCandidates(plans, sourceFileSystem, destinationFileSystem, candidates);
    if (stopped())
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    if (FAILED(candidatesHr))
        return candidatesHr;
    if (candidatesHr == S_FALSE)
    {
        out.reset();
        return S_OK;
    }

    FileOperationArtifacts::TouchGuardRequest request;
    const HRESULT requestHr = FileOperationArtifacts::BuildTouchGuardRequest(candidates, request);
    if (stopped())
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    if (FAILED(requestHr))
        return requestHr;
    if (requestHr == S_FALSE)
    {
        out.reset();
        return S_OK;
    }

    auto payload = std::unique_ptr<BatchRenameArtifactPromptPayload>(new (std::nothrow) BatchRenameArtifactPromptPayload{});
    if (! payload)
        return E_OUTOFMEMORY;
    payload->taskId          = _taskId;
    payload->request         = std::move(request);
    const uint64_t requestId = [&]() noexcept -> uint64_t
    {
        std::scoped_lock lock(_batchRenameArtifactPromptMutex);
        // Zero is reserved. Do not roll over and accept an old posted reply.
        // A completed result still belongs to its waiting caller until consumed.
        if (_batchRenameArtifactPromptPendingRequestId != 0u || _batchRenameArtifactPromptNextRequestId == (std::numeric_limits<uint64_t>::max)())
            return 0u;
        const uint64_t next                        = ++_batchRenameArtifactPromptNextRequestId;
        _batchRenameArtifactPromptPendingRequestId = next;
        _batchRenameArtifactPromptStatus           = E_PENDING;
        _batchRenameArtifactPromptCompleted        = false;
        _batchRenameArtifactPromptResultReceipt.reset();
        return next;
    }();
    if (requestId == 0u)
        return HRESULT_FROM_WIN32(ERROR_BUSY); // overflow is also fail-closed
    payload->requestId = requestId;
    if (! PostMessagePayload(_state->_owner.GetHwnd(), WndMsg::kFileOperationBatchRenameArtifactPrompt, static_cast<WPARAM>(_taskId), std::move(payload)))
    {
        std::scoped_lock lock(_batchRenameArtifactPromptMutex);
        if (_batchRenameArtifactPromptPendingRequestId == requestId)
        {
            _batchRenameArtifactPromptPendingRequestId = 0u;
            _batchRenameArtifactPromptResultReceipt.reset();
        }
        return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
    }

    std::unique_lock lock(_batchRenameArtifactPromptMutex);
    _batchRenameArtifactPromptCv.wait(
        lock, [this]() noexcept { return _batchRenameArtifactPromptCompleted || _cancelled.load(std::memory_order_acquire) || _stopToken.stop_requested(); });
    if (stopped())
    {
        if (_batchRenameArtifactPromptPendingRequestId == requestId)
        {
            _batchRenameArtifactPromptPendingRequestId = 0u;
            _batchRenameArtifactPromptResultReceipt.reset();
        }
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }
    if (_batchRenameArtifactPromptPendingRequestId != requestId)
        return E_UNEXPECTED;
    const HRESULT status                       = _batchRenameArtifactPromptStatus;
    auto accepted                              = std::move(_batchRenameArtifactPromptResultReceipt);
    _batchRenameArtifactPromptPendingRequestId = 0u;
    _batchRenameArtifactPromptResultReceipt.reset();
    if (status != S_OK)
        return status;
    if (! accepted)
        return E_UNEXPECTED;
    out = std::move(accepted);
    return S_OK;
}

bool FolderWindow::FileOperationState::Task::IsBatchRenameArtifactPromptPending(const uint64_t requestId) noexcept
{
    std::scoped_lock lock(_batchRenameArtifactPromptMutex);
    return requestId != 0u && requestId == _batchRenameArtifactPromptPendingRequestId && ! _batchRenameArtifactPromptCompleted &&
           ! _cancelled.load(std::memory_order_acquire) && ! _stopToken.stop_requested();
}

void FolderWindow::FileOperationState::Task::CompleteBatchRenameArtifactPrompt(const uint64_t requestId,
                                                                               HRESULT status,
                                                                               std::optional<FileOperationArtifacts::TouchGuardReceipt> receipt) noexcept
{
    {
        std::scoped_lock lock(_batchRenameArtifactPromptMutex);
        if (requestId == 0u || requestId != _batchRenameArtifactPromptPendingRequestId || _batchRenameArtifactPromptCompleted)
            return;
        if (status == S_OK && ! receipt)
            status = E_UNEXPECTED;
        else if (status != S_OK && SUCCEEDED(status))
            status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        _batchRenameArtifactPromptStatus = status;
        if (status == S_OK)
            _batchRenameArtifactPromptResultReceipt = std::move(receipt);
        else
            _batchRenameArtifactPromptResultReceipt.reset();
        _batchRenameArtifactPromptCompleted = true;
    }
    _batchRenameArtifactPromptCv.notify_all();
}

void FolderWindow::FileOperationState::Task::FailBatchRenameArtifactPromptWithoutPayload(const HRESULT status) noexcept
{
    {
        std::scoped_lock lock(_batchRenameArtifactPromptMutex);
        if (_batchRenameArtifactPromptCompleted)
            return;
        _batchRenameArtifactPromptStatus = FAILED(status) ? status : E_UNEXPECTED;
        _batchRenameArtifactPromptResultReceipt.reset();
        _batchRenameArtifactPromptCompleted = true;
    }
    _batchRenameArtifactPromptCv.notify_all();
}

void FolderWindow::FileOperationState::CompleteBatchRenameArtifactPromptByTaskId(const uint64_t taskId,
                                                                                 const uint64_t requestId,
                                                                                 const HRESULT status,
                                                                                 std::optional<FileOperationArtifacts::TouchGuardReceipt> receipt) noexcept
{
    if (Task* const task = FindTask(taskId))
        task->CompleteBatchRenameArtifactPrompt(requestId, status, std::move(receipt));
}

void FolderWindow::FileOperationState::CompleteBatchRenameArtifactPromptByTaskId(const uint64_t taskId, const HRESULT status) noexcept
{
    if (Task* const task = FindTask(taskId))
        task->FailBatchRenameArtifactPromptWithoutPayload(status);
}

void FolderWindow::FileOperationState::OnBatchRenameArtifactPrompt(std::unique_ptr<BatchRenameArtifactPromptPayload> payload) noexcept
{
    if (! payload)
        return;
    const uint64_t taskId = payload->taskId;

    const uint64_t requestId = payload->requestId;
    Task* const task         = FindTask(taskId);
    if (! task || ! task->IsBatchRenameArtifactPromptPending(requestId))
        return;
    if (_completionShutdown.load(std::memory_order_acquire) || task->_cancelled.load(std::memory_order_acquire))
    {
        CompleteBatchRenameArtifactPromptByTaskId(taskId, requestId, HRESULT_FROM_WIN32(ERROR_CANCELLED));
        return;
    }
    FileOperationArtifacts::TouchGuardReceipt receipt;
    HRESULT promptHr = ConfirmArtifactTouch(_owner.GetHwnd(), payload->request, receipt);
    if (promptHr == S_FALSE)
        promptHr = HRESULT_FROM_WIN32(ERROR_CANCELLED);
    // ConfirmArtifactTouch may pump teardown or another prompt. The dispatch scope in
    // FolderWindow's WndProc keeps FileOperationState alive; do not use `task` again.
    CompleteBatchRenameArtifactPromptByTaskId(
        taskId, requestId, promptHr, promptHr == S_OK ? std::optional<FileOperationArtifacts::TouchGuardReceipt>(std::move(receipt)) : std::nullopt);
}

HRESULT FolderWindow::FileOperationState::Task::RevalidateArtifactTouchGuard(const SelectedAdmissionView& admission) noexcept
{
    if (! admission.artifactTouchReceipt)
    {
        return S_OK;
    }
    const std::shared_ptr<const FileOperations::FileOperationPlanGroup> plans = admission.plans;
    if (! plans || ! _fileSystem)
    {
        return E_UNEXPECTED;
    }

    Debug::Perf::Scope perf(L"fileops.artifact.touch.revalidate.us");
    std::vector<FileOperationArtifacts::Candidate> current;
    const HRESULT captureHr = CollectArtifactTouchCandidates(*plans, _fileSystem.get(), admission.destination->fileSystem.get(), current);
    if (captureHr != S_OK)
    {
        const HRESULT status = FAILED(captureHr) ? captureHr : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        perf.SetHr(status);
        return status;
    }
    const HRESULT revalidateHr = FileOperationArtifacts::RevalidateTouchGuard(*admission.artifactTouchReceipt, current);
    perf.SetValue0(static_cast<uint64_t>(current.size()));
    perf.SetHr(revalidateHr);
    return revalidateHr;
}

bool FolderWindow::FileOperationState::OpenCompletedTaskSource(uint64_t taskId) noexcept
{
    CompletedTaskSummary summary{};
    {
        std::scoped_lock lock(_mutex);
        const auto found = std::ranges::find_if(_completedTasks, [taskId](const CompletedTaskSummary& value) noexcept { return value.taskId == taskId; });
        if (found == _completedTasks.end())
        {
            return false;
        }
        summary = *found;
    }
    if ((summary.retainedSourcePaths.empty() && summary.unknownSourcePaths.empty()) || summary.sourcePluginId.empty() || summary.sourcePluginShortId.empty())
    {
        return false;
    }

    const std::filesystem::path sourcePath = ! summary.retainedSourcePaths.empty() ? summary.retainedSourcePaths.front() : summary.unknownSourcePaths.front();
    const std::filesystem::path parent     = sourcePath.parent_path();
    const std::wstring leaf                = sourcePath.filename().wstring();
    return ! parent.empty() &&
           SUCCEEDED(_owner.ExecuteInPaneLocation(
               summary.sourcePane, summary.sourcePluginId, summary.sourcePluginShortId, summary.sourceInstanceContext, parent, leaf, 0u, true));
}

bool FolderWindow::FileOperationState::SelectCompletedTaskRetainedSources(uint64_t taskId) noexcept
{
    return ExecuteCompletedRetainedSourceAction(taskId, false);
}

bool FolderWindow::FileOperationState::CutCompletedTaskRetainedSourcesAgain(uint64_t taskId) noexcept
{
    return ExecuteCompletedRetainedSourceAction(taskId, true);
}

bool FolderWindow::FileOperationState::ExecuteCompletedRetainedSourceAction(uint64_t taskId, bool cutAgain) noexcept
{
    CompletedTaskSummary summary{};
    {
        std::scoped_lock lock(_mutex);
        const auto found = std::ranges::find_if(_completedTasks, [taskId](const CompletedTaskSummary& value) noexcept { return value.taskId == taskId; });
        if (found == _completedTasks.end())
        {
            return false;
        }
        summary = *found;
    }

    const bool exactCompleteSet = summary.clipboardMoveAdmission && summary.operation == FILESYSTEM_MOVE && summary.retainedSourceCount > 0u &&
                                  summary.unknownSourceCount == 0u && summary.exactRetainedSourceItems.size() == summary.retainedSourceCount;
    if (! exactCompleteSet || ! NavigationLocation::IsFilePluginShortId(summary.sourcePluginShortId))
    {
        return false;
    }

    const std::filesystem::path firstPath   = summary.exactRetainedSourceItems.front().providerPath;
    const std::filesystem::path firstParent = firstPath.parent_path();
    if (firstParent.empty() || FAILED(_owner.ExecuteInPaneLocation(summary.sourcePane,
                                                                   summary.sourcePluginId,
                                                                   summary.sourcePluginShortId,
                                                                   summary.sourceInstanceContext,
                                                                   firstParent,
                                                                   firstPath.filename().wstring(),
                                                                   0u,
                                                                   true)))
    {
        return false;
    }

    FolderWindow::PaneState& paneState = summary.sourcePane == FolderWindow::Pane::Left ? _owner._leftPane : _owner._rightPane;
    if (! paneState.fileSystem || ! NavigationLocation::EqualsNoCase(paneState.pluginId, summary.sourcePluginId) ||
        ! NavigationLocation::EqualsNoCase(paneState.pluginShortId, summary.sourcePluginShortId) ||
        ! NavigationLocation::EqualsNoCase(paneState.instanceContext, summary.sourceInstanceContext))
    {
        return false;
    }

    std::vector<std::filesystem::path> exactPaths;
    if (! ValidateRetainedSourceActionItems(paneState.fileSystem.get(), summary.exactRetainedSourceItems, &exactPaths))
    {
        _owner.ShowPaneAlertOverlay(summary.sourcePane,
                                    FolderView::ErrorOverlayKind::Operation,
                                    FolderView::OverlaySeverity::Warning,
                                    LoadStringResource(nullptr, IDS_CAPTION_WARNING),
                                    LoadStringResource(nullptr, IDS_FILEOPS_RETAINED_SOURCE_CHANGED),
                                    HRESULT_FROM_WIN32(ERROR_RETRY),
                                    true,
                                    false);
        return false;
    }

    if (cutAgain && ! paneState.folderView.CutPathsToClipboard(exactPaths))
    {
        _owner.ShowPaneAlertOverlay(summary.sourcePane,
                                    FolderView::ErrorOverlayKind::Operation,
                                    FolderView::OverlaySeverity::Warning,
                                    LoadStringResource(nullptr, IDS_CAPTION_WARNING),
                                    LoadStringResource(nullptr, IDS_MSG_CLIPBOARD_WRITE_FAILED),
                                    E_FAIL,
                                    true,
                                    false);
        return false;
    }

    std::vector<std::wstring> visibleLeaves;
    for (const std::filesystem::path& path : exactPaths)
    {
        if (NavigationLocation::EqualsNoCase(path.parent_path().native(), firstParent.native()))
        {
            visibleLeaves.push_back(path.filename().wstring());
        }
    }
    if (! visibleLeaves.empty())
    {
        paneState.folderView.SelectDisplayNamesAfterNextEnumeration(firstParent, std::move(visibleLeaves));
        paneState.folderView.ForceRefresh();
    }
    return true;
}

void FolderWindow::FileOperationState::ApplyTheme(const AppTheme& /*theme*/)
{
    HWND popup      = nullptr;
    HWND issuesPane = nullptr;
    {
        std::scoped_lock lock(_mutex);
        popup      = _popup.get();
        issuesPane = _issuesPane.get();
    }

    if (popup)
    {
        PostMessageW(popup, WM_THEMECHANGED, 0, 0);
    }

    if (issuesPane)
    {
        PostMessageW(issuesPane, WM_THEMECHANGED, 0, 0);
    }
}

void FolderWindow::FileOperationState::Shutdown() noexcept
{
    if (_completionShutdown.exchange(true, std::memory_order_acq_rel))
    {
        return;
    }

    // I35 R8.8: the modeless decision mirror goes first, before the UI lifetime ends: hidden without
    // handing focus back, with no callback left that could reach this state.
    ShutdownDecisionMirror();

    std::vector<Task*> tasks;
    wil::unique_hwnd popupToClose;
    wil::unique_hwnd issuesPaneToClose;
    HWND popupHwnd      = nullptr;
    HWND issuesPaneHwnd = nullptr;
    {
        std::scoped_lock lock(_mutex);
        _uiLifetime.reset();
        _pendingConflictInspection.reset();
        _completedConflictInspection.reset();
        tasks.reserve(_tasks.size());
        for (const std::unique_ptr<Task>& task : _tasks)
        {
            if (task)
            {
                tasks.push_back(task.get());
            }
        }
        popupHwnd         = _popup.get();
        issuesPaneHwnd    = _issuesPane.get();
        popupToClose      = std::move(_popup);
        issuesPaneToClose = std::move(_issuesPane);
    }

    if (_owner._settings)
    {
        if (popupHwnd)
        {
            SavePopupPlacement(popupHwnd);
        }

        if (issuesPaneHwnd)
        {
            SaveIssuesPanePlacement(issuesPaneHwnd);
        }

        if (popupHwnd || issuesPaneHwnd)
        {
            QueueSettingsSave(L"shutdown");
        }
    }

    if (_hasQueuedSettingsSave.load(std::memory_order_acquire))
    {
        FlushPendingSettingsSave();
    }

    for (Task* task : tasks)
    {
        if (task)
        {
            task->RequestCancel();
        }
    }

    std::vector<std::jthread> completionWorkers;
    {
        std::scoped_lock lock(_completionWorkerTransferMutex);
        completionWorkers.reserve(tasks.size());
        for (Task* task : tasks)
        {
            if (task && task->_thread.joinable())
            {
                completionWorkers.push_back(std::move(task->_thread));
            }
        }
    }
    for (std::jthread& worker : completionWorkers)
    {
        if (worker.joinable())
        {
            worker.join();
        }
    }

    {
        std::unique_lock lock(_orphanedCompletionDrainMutex);
        _orphanedCompletionDrainCv.wait(lock, [this]() noexcept { return _orphanedCompletionDrainsOutstanding == 0u; });
    }

    std::vector<FallbackCompletion> pendingCompletions;
    {
        std::scoped_lock lock(_fallbackCompletedMutex);
        pendingCompletions.swap(_fallbackCompletedPayloads);
    }
    for (const FallbackCompletion& entry : pendingCompletions)
    {
        const TaskCompletedPayload& completed = entry.payload;
        _owner.ApplyFileOperationCompletion(completed.taskId, completed.hr, completed.warningCount, completed.errorCount, completed.partialOutcome);
    }

    std::vector<std::unique_ptr<Task>> remainingTasks;
    {
        std::scoped_lock lock(_mutex);
        remainingTasks.swap(_tasks);
    }

    // All task and reaper worker handles are joined before pending completion is
    // applied or remaining task storage is destroyed.
    FlushDiagnostics(true);
}

void FolderWindow::FileOperationState::NotifyQueueChanged()
{
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Queue.NotifyAllCount", L"", 0, 1u, 0u, S_OK);
    }
    _queueCv.notify_all();
}

HRESULT FolderWindow::FileOperationState::ConfirmExternalArtifactTouch(const wil::com_ptr<IFileSystem>& fileSystem,
                                                                       const std::wstring_view pluginId,
                                                                       const std::wstring_view instanceContext,
                                                                       const std::span<const std::filesystem::path> providerPaths,
                                                                       FileOperationArtifacts::TouchGuardReceipt* const receiptOut,
                                                                       const std::optional<FolderWindow::Pane> anchorPane) noexcept
{
    Debug::Perf::Scope perf(L"fileops.artifact.touch.external.us");
    if (receiptOut != nullptr)
    {
        *receiptOut = {};
    }
    if (! fileSystem || pluginId.empty() || providerPaths.empty())
    {
        perf.SetHr(E_INVALIDARG);
        return E_INVALIDARG;
    }

    const auto collect = [&](std::vector<FileOperationArtifacts::Candidate>& candidates) -> HRESULT
    {
        candidates.clear();
        candidates.reserve(providerPaths.size());
        for (const std::filesystem::path& path : providerPaths)
        {
            FileOperationArtifacts::Candidate candidate{};
            const HRESULT captureHr =
                FileOperationArtifacts::CaptureProviderObjectCandidate(fileSystem.get(), path.native(), pluginId, instanceContext, candidate);
            if (FAILED(captureHr))
            {
                return captureHr;
            }
            if (captureHr != S_OK)
            {
                continue;
            }
            const auto duplicate = std::ranges::find_if(candidates,
                                                        [&](const FileOperationArtifacts::Candidate& existing) noexcept
            {
                return SameArtifactEndpoint(existing.endpoint, candidate.endpoint) &&
                       EquivalentPath(candidate.pathIdentity, existing.path.native(), candidate.path.native());
            });
            if (duplicate == candidates.end())
            {
                candidates.push_back(std::move(candidate));
            }
        }
        return candidates.empty() ? S_FALSE : S_OK;
    };

    FileOperationArtifacts::TouchGuardReceipt receipt;
    for (;;)
    {
        std::vector<FileOperationArtifacts::Candidate> candidates;
        const HRESULT collectHr = collect(candidates);
        if (FAILED(collectHr))
        {
            perf.SetHr(collectHr);
            return collectHr;
        }
        if (collectHr == S_FALSE)
        {
            perf.SetValue0(0u);
            perf.SetHr(S_OK);
            return S_OK;
        }

        FileOperationArtifacts::TouchGuardRequest request;
        const HRESULT requestHr = FileOperationArtifacts::BuildTouchGuardRequest(candidates, request);
        if (FAILED(requestHr))
        {
            perf.SetHr(requestHr);
            return requestHr;
        }
        if (requestHr == S_FALSE)
        {
            perf.SetValue0(0u);
            perf.SetHr(S_OK);
            return S_OK;
        }

        // I35 R8.10: on the pane the request came from when the caller knows it, else on the Folder window.
        const HWND promptAnchor = anchorPane.has_value() ? IngressPromptAnchor(anchorPane.value()) : _owner.GetHwnd();
        const HRESULT confirmHr = ConfirmArtifactTouch(promptAnchor, request, receipt, true);
        if (confirmHr == HRESULT_FROM_WIN32(ERROR_RETRY) && ! _completionShutdown.load(std::memory_order_acquire))
        {
            continue; // R7.1: capture the candidates again; an item that could not be opened may open now
        }
        if (! request.allItemsRevalidatable && confirmHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED))
        {
            // The blocked prompt already said the operation was not started; callers show nothing more.
            perf.SetValue0(static_cast<uint64_t>(request.items.size()));
            perf.SetHr(FolderWindow::kFileOperationRefusedAlreadyReported);
            return FolderWindow::kFileOperationRefusedAlreadyReported;
        }
        if (confirmHr != S_OK)
        {
            perf.SetValue0(static_cast<uint64_t>(request.items.size()));
            perf.SetHr(confirmHr == HRESULT_FROM_WIN32(ERROR_RETRY) ? HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS) : confirmHr);
            return confirmHr == HRESULT_FROM_WIN32(ERROR_RETRY) ? HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS) : confirmHr;
        }
        break;
    }

    std::vector<FileOperationArtifacts::Candidate> currentCandidates;
    const HRESULT currentCollectHr = collect(currentCandidates);
    if (currentCollectHr != S_OK)
    {
        const HRESULT status = FAILED(currentCollectHr) ? currentCollectHr : HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        perf.SetHr(status);
        return status;
    }
    const HRESULT revalidateHr = FileOperationArtifacts::RevalidateTouchGuard(receipt, currentCandidates);
    perf.SetValue0(static_cast<uint64_t>(receipt.items.size()));
    perf.SetHr(revalidateHr);
    if (revalidateHr == S_OK && receiptOut != nullptr)
    {
        *receiptOut = std::move(receipt);
    }
    return revalidateHr;
}

bool FolderWindow::FileOperationState::HasActiveOperations() noexcept
{
    {
        std::scoped_lock lock(_mutex);
        if (! _tasks.empty())
        {
            return true;
        }
    }

    // Defensive fallback: active operations are expected to always have a task object.
    std::scoped_lock lock(_queueMutex);
    return _activeOperations > 0 || ! _queue.empty();
}

FolderWindow::FileOperationState::PopupTelemetryWork FolderWindow::FileOperationState::GetPopupTelemetryWork() noexcept
{
    std::vector<Task*> tasks;
    {
        std::scoped_lock lock(_mutex);
        if (std::ranges::any_of(_informationalTasks, [](const auto& task) noexcept { return ! task.finished; }))
            return PopupTelemetryWork::Active;
        if (_tasks.empty())
            return PopupTelemetryWork::None;
        tasks.reserve(_tasks.size());
        for (const auto& task : _tasks)
            tasks.push_back(task.get());
    }

    // x-perf#5: a task the user paused, a queued or overlap-waiting task and a task waiting on an
    // unanswered prompt change nothing on screen; they need the waiting cadence only. Like
    // BuildSnapshot, borrow the tasks outside _mutex: only this owner thread removes them, and the
    // prompt lock is never taken under _mutex.
    for (Task* task : tasks)
    {
        // An owned terminal worker awaits its owner-thread completion and a stopping task may roll
        // back or clean up; both stay telemetry work.
        if (! task || task->_taskFinished.load(std::memory_order_acquire) || task->CancelRequestedDurable())
            return PopupTelemetryWork::Active;
        // IsWaitingInQueue is set only while the worker is held before EnterOperation admits it;
        // leaving that wait clears it and posts a presentation change.
        if (task->IsPaused() || task->IsQueuePaused() || task->IsWaitingInQueue())
            continue;
        std::scoped_lock promptLock(task->_conflictArbiter.mutex);
        const auto& prompt = task->_conflictArbiter.prompt;
        if (! prompt.active || prompt.metadataLoading)
            return PopupTelemetryWork::Active;
    }
    return PopupTelemetryWork::Waiting;
}

bool FolderWindow::FileOperationState::ShouldQueueNewTask() noexcept
{
    if (! _queueNewTasks.load(std::memory_order_acquire))
    {
        return false;
    }

    return HasActiveOperations();
}

void FolderWindow::FileOperationState::SetQueueNewTasks(bool queue) noexcept
{
    _queueNewTasks.store(queue, std::memory_order_release);
}

bool FolderWindow::FileOperationState::GetQueueNewTasks() const noexcept
{
    return _queueNewTasks.load(std::memory_order_acquire);
}

void FolderWindow::FileOperationState::ApplyQueueMode(bool queue) noexcept
{
    _queueNewTasks.store(queue, std::memory_order_release);

    std::vector<Task*> tasks;
    CollectTasks(tasks);

    for (auto* task : tasks)
    {
        if (! task)
        {
            continue;
        }

        // Inline F2 bypasses global scheduling, including live mode changes.
        // Its mutation interlock and explicit user pause remain authoritative.
        if (! queue || task->IsInlineRenameTask())
        {
            task->SetWaitForOthers(false);
            continue;
        }

        if (! task->HasEnteredOperation())
        {
            task->SetWaitForOthers(true);
            continue;
        }
    }

    UpdateQueuePausedTasks();
    NotifyQueueChanged();
}

bool FolderWindow::FileOperationState::RunQueuedTaskNow(uint64_t taskId) noexcept
{
    if (taskId == 0)
    {
        return false;
    }

    Task* task = FindTask(taskId);
    if (! task || task->CancelRequestedDurable() || task->HasEnteredOperation() || (! task->IsWaitingForOthers() && ! task->IsWaitingInQueue()))
    {
        return false;
    }

    task->SetWaitForOthers(false);
    task->SetWaitingInQueue(false);
    UpdateQueuePausedTasks();
    NotifyQueueChanged();
    return true;
}

uint64_t FolderWindow::FileOperationState::AllocateLateQueueOrderKey() noexcept
{
    std::scoped_lock lock(_mutex);
    return _nextTaskId++;
}

bool FolderWindow::FileOperationState::MoveQueuedTask(uint64_t taskId, bool moveUp) noexcept
{
    const uint64_t startedUs = PerfNowUs();
    std::vector<Task*> tasks;
    CollectTasks(tasks);

    std::vector<Task*> waiting;
    waiting.reserve(tasks.size());
    for (Task* task : tasks)
    {
        if (task && ! task->CancelRequestedDurable() && ! task->HasEnteredOperation() && (task->IsWaitingForOthers() || task->IsWaitingInQueue()))
        {
            waiting.push_back(task);
        }
    }
    std::ranges::sort(waiting, {}, &Task::GetQueueOrderKey);

    const auto current = std::ranges::find(waiting, taskId, &Task::GetId);
    if (current == waiting.end())
    {
        return false;
    }

    const size_t index = static_cast<size_t>(std::distance(waiting.begin(), current));
    if ((moveUp && index == 0u) || (! moveUp && index + 1u >= waiting.size()))
    {
        return false;
    }

    const size_t adjacentIndex = moveUp ? index - 1u : index + 1u;
    Task* const adjacent       = waiting[adjacentIndex];
    const uint64_t currentKey  = (*current)->GetQueueOrderKey();
    const uint64_t adjacentKey = adjacent->GetQueueOrderKey();

    {
        std::scoped_lock lock(_queueMutex);
        if ((*current)->HasEnteredOperation() || adjacent->HasEnteredOperation() || (! (*current)->IsWaitingForOthers() && ! (*current)->IsWaitingInQueue()) ||
            (! adjacent->IsWaitingForOthers() && ! adjacent->IsWaitingInQueue()))
        {
            return false;
        }
        (*current)->_queueOrderKey.store(adjacentKey, std::memory_order_release);
        adjacent->_queueOrderKey.store(currentKey, std::memory_order_release);

        std::unordered_map<uint64_t, uint64_t> orderByTaskId;
        orderByTaskId.reserve(waiting.size());
        for (Task* waitingTask : waiting)
        {
            orderByTaskId.emplace(waitingTask->GetId(), waitingTask->GetQueueOrderKey());
        }
        std::ranges::sort(_queue,
                          [&](uint64_t left, uint64_t right)
        {
            const auto leftIt       = orderByTaskId.find(left);
            const auto rightIt      = orderByTaskId.find(right);
            const uint64_t leftKey  = leftIt != orderByTaskId.end() ? leftIt->second : left;
            const uint64_t rightKey = rightIt != orderByTaskId.end() ? rightIt->second : right;
            return leftKey < rightKey;
        });
    }

    NotifyQueueChanged();
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.Queue.ReorderUs", moveUp ? L"up" : L"down", PerfElapsedUs(startedUs), waiting.size(), index, S_OK);
    }
    return true;
}

void FolderWindow::FileOperationState::SetAllRunningTasksPaused(bool paused) noexcept
{
    std::vector<Task*> tasks;
    CollectTasks(tasks);

    for (Task* task : tasks)
    {
        if (! task || ! task->HasStarted() || task->CancelRequestedDurable() || task->_taskFinished.load(std::memory_order_acquire))
        {
            continue;
        }

        task->SetPaused(paused);
    }
}

bool FolderWindow::FileOperationState::HasCancelableOperations() noexcept
{
    // Popup availability needs engine truth, including work awaiting its first
    // visible card. Inspect atomics under the task ownership lock without copying
    // a vector of borrowed task pointers on every layout/failure-surface tick.
    std::scoped_lock lock(_mutex);
    return std::ranges::any_of(
        _tasks, [](const auto& task) noexcept { return task && ! task->CancelRequestedDurable() && ! task->_taskFinished.load(std::memory_order_acquire); });
}

std::vector<uint64_t> FolderWindow::FileOperationState::CancelAll() noexcept
{
    std::vector<Task*> tasks;
    {
        std::scoped_lock lock(_mutex);
        tasks.reserve(_tasks.size());
        for (const auto& task : _tasks)
        {
            if (task)
            {
                tasks.push_back(task.get());
            }
        }
    }

    std::vector<uint64_t> canceled;
    canceled.reserve(tasks.size());
    for (Task* task : tasks)
    {
        if (task && ! task->CancelRequestedDurable() && ! task->_taskFinished.load(std::memory_order_acquire))
        {
            task->RequestCancel();
            canceled.push_back(task->_taskId);
        }
    }
    if (GetWindowThreadProcessId(_owner.GetHwnd(), nullptr) == GetCurrentThreadId() && _pendingConflictInspection)
        CancelConflictInspectionForTask(_pendingConflictInspection->taskId);
    return canceled;
}

void FolderWindow::FileOperationState::CollectTasks(std::vector<Task*>& outTasks) noexcept
{
    std::scoped_lock lock(_mutex);
    outTasks.clear();
    outTasks.reserve(_tasks.size());
    for (const auto& task : _tasks)
    {
        if (task)
        {
            outTasks.push_back(task.get());
        }
    }
}

void FolderWindow::FileOperationState::CollectInformationalTasks(std::vector<FolderWindow::InformationalTaskUpdate>& outTasks) noexcept
{
    const uint64_t waitStartedUs = PerfNowUs();
    std::scoped_lock lock(_mutex);
    const uint64_t waitUs        = PerfElapsedUs(waitStartedUs);
    const uint64_t copyStartedUs = PerfNowUs();
    outTasks.clear();
    outTasks.reserve(_informationalTasks.size());
    for (const auto& task : _informationalTasks)
    {
        outTasks.push_back(task);
    }
    if (Debug::Perf::IsCaptureEnabled())
    {
        Debug::Perf::Emit(L"FileOps.InfoTask.Collect.LockWaitUs", L"", waitUs, static_cast<uint64_t>(_informationalTasks.size()), 0u, S_OK);
        Debug::Perf::Emit(L"FileOps.InfoTask.Collect.CopyUs", L"", PerfElapsedUs(copyStartedUs), static_cast<uint64_t>(_informationalTasks.size()), 0u, S_OK);
    }
}

void FolderWindow::FileOperationState::CollectCompletedTasks(std::vector<CompletedTaskSummary>& outTasks) noexcept
{
    std::scoped_lock lock(_mutex);
    outTasks.clear();
    outTasks.reserve(_completedTasks.size());
    for (const auto& summary : _completedTasks)
    {
        outTasks.push_back(summary);
    }
}

std::shared_ptr<const std::vector<FolderWindow::FileOperationState::CompletedTaskPresentation>> FolderWindow::FileOperationState::
    CollectCompletedTaskPresentations() noexcept
{
    // Every _completedTasks change bumps _issuesRevision under _mutex, so an unchanged revision
    // proves the shared projection is current. Readers only ever see an immutable vector.
    std::scoped_lock lock(_mutex);
    const uint64_t revision         = _issuesRevision.load(std::memory_order_relaxed);
    const uint64_t languageRevision = Localization::GetLanguagePreferenceRevision();
    if (_completedPresentation && _completedPresentationRevision == revision && _completedPresentationLanguageRevision == languageRevision)
        return _completedPresentation;

    auto presentation = std::make_shared<std::vector<CompletedTaskPresentation>>();
    presentation->reserve(_completedTasks.size());
    for (auto& summary : _completedTasks)
    {
        // Q9: a Last note with a typed source follows the live language; it is formatted again
        // only after a language change.
        if (summary.lastNoteLanguageRevision != languageRevision &&
            (summary.lastProblem.has_value() || summary.lastNoteText.messageId != 0u || summary.lastInfoNote.has_value()))
        {
            summary.lastDiagnosticMessage    = FormatCompletedLastNote(summary);
            summary.lastInfoNoteMessage      = FormatCompletedLastInfoNote(summary);
            summary.lastNoteLanguageRevision = languageRevision;
        }
        CompletedTaskPresentation& item   = presentation->emplace_back();
        item.facts                        = static_cast<const CompletedTaskFacts&>(summary);
        item.hasSourceActionPaths         = ! summary.retainedSourcePaths.empty() || ! summary.unknownSourcePaths.empty();
        item.exactRetainedSourceItemCount = summary.exactRetainedSourceItems.size();
        item.hasRetainedDiagnostics       = ! summary.issueDiagnostics.empty();
    }
    _completedPresentation                 = std::move(presentation);
    _completedPresentationRevision         = revision;
    _completedPresentationLanguageRevision = languageRevision;
#ifdef ENABLE_TESTS
    _debugCompletedPresentationBuilds.fetch_add(1u, std::memory_order_relaxed);
#endif
    return _completedPresentation;
}

void FolderWindow::FileOperationState::CollectDiagnostics(std::vector<TaskDiagnosticEntry>& outEntries) noexcept
{
    std::scoped_lock lock(_diagnosticsMutex);
    outEntries.clear();
    outEntries.reserve(_diagnosticsInMemory.size());
    for (const auto& entry : _diagnosticsInMemory)
    {
        outEntries.push_back(entry);
    }
}

bool FolderWindow::FileOperationState::CollectIssuesSnapshotIfChanged(uint64_t lastRevision,
                                                                      std::vector<TaskDiagnosticEntry>& outDiagnostics,
                                                                      std::vector<TaskDiagnosticEntry>& outCompletedIssues,
                                                                      uint64_t& outCapturedRevision) noexcept
{
    // Writers increment the shared revision under their reservoir's lock. Taking
    // both locks yields one coherent, bounded snapshot and skips all copies on idle.
    std::scoped_lock lock(_mutex, _diagnosticsMutex);
    const uint64_t revision = _issuesRevision.load(std::memory_order_relaxed);
    if (lastRevision == revision)
        return false;
    outDiagnostics.assign(_diagnosticsInMemory.begin(), _diagnosticsInMemory.end());
    // model#3: the Issues projection reads only the retained issue entries, never the per-root paths.
    size_t completedIssueCount = 0u;
    for (const CompletedTaskSummary& summary : _completedTasks)
        completedIssueCount += summary.issueDiagnostics.size();
    outCompletedIssues.clear();
    outCompletedIssues.reserve(completedIssueCount);
    for (const CompletedTaskSummary& summary : _completedTasks)
        outCompletedIssues.insert(outCompletedIssues.end(), summary.issueDiagnostics.begin(), summary.issueDiagnostics.end());
    outCapturedRevision = revision;
    return true;
}

void FolderWindow::FileOperationState::GetTaskProblemCounts(const uint64_t taskId, unsigned long& warningCount, unsigned long& errorCount) noexcept
{
    warningCount = 0u;
    errorCount   = 0u;
    std::scoped_lock lock(_diagnosticsMutex);
    if (const auto countsIt = _taskDiagnosticCounts.find(taskId); countsIt != _taskDiagnosticCounts.end())
    {
        warningCount = countsIt->second.first;
        errorCount   = countsIt->second.second;
    }
}

void FolderWindow::FileOperationState::CollectTaskDiagnosticSnapshot(uint64_t taskId,
                                                                     unsigned long& warningCount,
                                                                     unsigned long& errorCount,
                                                                     std::wstring& lastDiagnosticMessage,
                                                                     unsigned long* noteCount,
                                                                     std::wstring* lastInfoNoteMessage) noexcept
{
    warningCount = 0;
    errorCount   = 0;
    lastDiagnosticMessage.clear();
    if (noteCount != nullptr)
        *noteCount = 0u;
    if (lastInfoNoteMessage != nullptr)
        lastInfoNoteMessage->clear();

    std::optional<TaskDiagnosticEntry> lastProblem;
    std::optional<TaskDiagnosticEntry> lastNote;
    {
        std::scoped_lock lock(_diagnosticsMutex);
        const auto countsIt = _taskDiagnosticCounts.find(taskId);
        if (countsIt != _taskDiagnosticCounts.end())
        {
            warningCount = countsIt->second.first;
            errorCount   = countsIt->second.second;
        }

        const auto problemIt = _taskLastProblem.find(taskId);
        if (problemIt != _taskLastProblem.end())
        {
            lastProblem = problemIt->second;
        }
        if (noteCount != nullptr)
        {
            if (const auto notesIt = _taskNoteCounts.find(taskId); notesIt != _taskNoteCounts.end())
                *noteCount = notesIt->second;
        }
        if (lastInfoNoteMessage != nullptr)
        {
            if (const auto noteIt = _taskLastNote.find(taskId); noteIt != _taskLastNote.end())
                lastNote = noteIt->second;
        }
    }
    // Q9: the typed text is formatted by the reader, outside the lock, in the live language.
    if (lastProblem.has_value())
    {
        lastDiagnosticMessage = FormatIssueProblemText(lastProblem.value());
    }
    if (lastNote.has_value() && lastInfoNoteMessage != nullptr)
    {
        *lastInfoNoteMessage = FormatIssueProblemText(lastNote.value());
    }
}

void FolderWindow::FileOperationState::DismissCompletedTask(const uint64_t sessionTaskId) noexcept
{
    wil::unique_hwnd popupToClose;
    std::vector<std::filesystem::path> acknowledgedBreadcrumbs;

    {
        std::scoped_lock lock(_mutex);
        for (const CompletedTaskSummary& summary : _completedTasks)
        {
            if (summary.taskId == sessionTaskId && ! summary.interruptedMoveBreadcrumbPath.empty())
            {
                acknowledgedBreadcrumbs.push_back(summary.interruptedMoveBreadcrumbPath);
            }
        }
        _completedTasks.erase(std::remove_if(_completedTasks.begin(),
                                             _completedTasks.end(),
                                             [&](const CompletedTaskSummary& summary) noexcept { return summary.taskId == sessionTaskId; }),
                              _completedTasks.end());
        _issuesRevision.fetch_add(1u, std::memory_order_relaxed);

        if (_tasks.empty() && _completedTasks.empty() && _informationalTasks.empty())
        {
            popupToClose = std::move(_popup);
        }
    }
    for (const std::filesystem::path& breadcrumb : acknowledgedBreadcrumbs)
    {
        const HRESULT acknowledgeHr = FileOperationMoveBreadcrumb::Acknowledge(breadcrumb);
        if (FAILED(acknowledgeHr))
        {
            Debug::Warning(L"File Operations could not acknowledge an interrupted-Move breadcrumb (hr=0x{:08X}, path='{}').",
                           static_cast<unsigned long>(acknowledgeHr),
                           breadcrumb.native());
        }
    }
}

void FolderWindow::FileOperationState::ClearCompletedTasks(const std::span<const uint64_t> sessionTaskIds) noexcept
{
    if (sessionTaskIds.empty())
    {
        return;
    }

    wil::unique_hwnd popupToClose;
    {
        std::scoped_lock lock(_mutex);
        const auto removed = std::ranges::remove_if(_completedTasks,
                                                    [&](const CompletedTaskSummary& summary) noexcept
        {
            return summary.interruptedMoveBreadcrumbPath.empty() && ! summary.interruptedMoveNotice &&
                   std::ranges::find(sessionTaskIds, summary.taskId) != sessionTaskIds.end();
        });
        if (removed.empty())
        {
            return;
        }
        _completedTasks.erase(removed.begin(), removed.end());
        _issuesRevision.fetch_add(1u, std::memory_order_relaxed);

        if (_tasks.empty() && _completedTasks.empty() && _informationalTasks.empty())
        {
            popupToClose = std::move(_popup);
        }
    }
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::DebugAppendCompletedTaskForSelfTest(CompletedTaskSummary summary) noexcept
{
    std::scoped_lock lock(_mutex);
    _completedTasks.push_back(std::move(summary));
    _issuesRevision.fetch_add(1u, std::memory_order_relaxed);
}

FolderWindow::FileOperationState::CompletedTaskSummary FolderWindow::FileOperationState::DebugRecordCompletedTaskForSelfTest(Task& task) noexcept
{
    return RecordCompletedTask(task);
}
#endif

ULONGLONG FolderWindow::FileOperationState::TaskPresentationNowTick() const noexcept
{
#ifdef ENABLE_TESTS
    const ULONGLONG testTick = _debugTaskPresentationNowTick.load(std::memory_order_acquire);
    if (testTick != kTaskPresentationLiveClock)
    {
        return testTick;
    }
#endif
    return GetTickCount64();
}

void FolderWindow::FileOperationState::RequestTaskPresentationRefresh() noexcept
{
#ifdef ENABLE_TESTS
    if (_debugSuppressTaskPresentationPosts.load(std::memory_order_acquire))
        return;
#endif
    const HWND owner = _owner.GetHwnd();
    if (owner)
    {
        static_cast<void>(PostMessageW(owner, WndMsg::kFileOperationPresentationChanged, 0, 0));
    }
}

HRESULT FolderWindow::FileOperationState::ArmCompletionRecoveryTimer() noexcept
{
    const HWND ownerHwnd = _owner.GetHwnd();
    if (! ownerHwnd || GetWindowThreadProcessId(ownerHwnd, nullptr) != GetCurrentThreadId())
        return HRESULT_FROM_WIN32(ERROR_INVALID_THREAD_ID);
#ifdef ENABLE_TESTS
    // Owner-thread only: let the configured number of arms pass, then fail the next one once.
    const int32_t passingArms = _debugRecoveryTimerArmPassesBeforeFailure.load(std::memory_order_acquire);
    if (passingArms == 0)
    {
        _debugRecoveryTimerArmPassesBeforeFailure.store(-1, std::memory_order_release);
        return HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA);
    }
    if (passingArms > 0)
        _debugRecoveryTimerArmPassesBeforeFailure.store(passingArms - 1, std::memory_order_release);
#endif
    if (SetTimer(ownerHwnd, FolderWindow::kFileOperationPresentationTimerId, kCompletionRecoveryHeartbeatMs, nullptr) != 0)
        return S_OK;
    const DWORD error = GetLastError();
    return HRESULT_FROM_WIN32(error != ERROR_SUCCESS ? error : ERROR_NOT_ENOUGH_MEMORY);
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::DebugFailRecoveryTimerArmForSelfTest(const int32_t passingArms) noexcept
{
    _debugRecoveryTimerArmPassesBeforeFailure.store(passingArms < 0 ? -1 : passingArms, std::memory_order_release);
}
#endif

void FolderWindow::FileOperationState::RequestActionablePromptPresentation() noexcept
{
    _actionablePromptPresentationPending.store(true, std::memory_order_release);
    RequestTaskPresentationRefresh();
}

std::optional<UINT> FolderWindow::FileOperationState::RefreshDeferredTaskPresentation() noexcept
{
    if (_actionablePromptPresentationPending.exchange(false, std::memory_order_acq_rel))
    {
        EnsurePopupVisible();
    }
    if (_interruptedMoveSummariesPendingPresentation.exchange(false, std::memory_order_acq_rel))
    {
        EnsurePopupVisible();
    }
    const ULONGLONG nowTick = TaskPresentationNowTick();
    std::optional<ULONGLONG> nextDeadline;
    bool showPopup = false;

    {
        std::scoped_lock lock(_mutex);
        for (const auto& task : _tasks)
        {
            if (! task)
            {
                continue;
            }

            Task::TaskPresentationState presentation = task->_presentationState.load(std::memory_order_acquire);
            if (presentation == Task::TaskPresentationState::RevealRequested)
            {
                task->_presentationState.store(Task::TaskPresentationState::Presented, std::memory_order_release);
                showPopup = true;
#ifdef ENABLE_TESTS
                _debugTaskPresentedCount.fetch_add(1u, std::memory_order_relaxed);
#endif
                continue;
            }
            if (presentation != Task::TaskPresentationState::Hidden || task->_taskFinished.load(std::memory_order_acquire))
            {
                continue;
            }

            if (nowTick >= task->_presentationDeadlineTick)
            {
                Task::TaskPresentationState expected = Task::TaskPresentationState::Hidden;
                if (task->_presentationState.compare_exchange_strong(
                        expected, Task::TaskPresentationState::Presented, std::memory_order_acq_rel, std::memory_order_acquire))
                {
                    showPopup                    = true;
                    const ULONGLONG admittedTick = task->_presentationDeadlineTick >= FileOperations::kTaskCardRevealDelayMs
                                                       ? task->_presentationDeadlineTick - FileOperations::kTaskCardRevealDelayMs
                                                       : 0u;
                    Debug::Perf::Emit(L"FileOps.TaskPresentation.RevealMs",
                                      L"deadline",
                                      nowTick >= admittedTick ? nowTick - admittedTick : 0u,
                                      FileOperations::kTaskCardRevealDelayMs,
                                      0u,
                                      S_OK);
#ifdef ENABLE_TESTS
                    _debugTaskPresentedCount.fetch_add(1u, std::memory_order_relaxed);
#endif
                }
                else if (expected == Task::TaskPresentationState::RevealRequested)
                {
                    task->_presentationState.store(Task::TaskPresentationState::Presented, std::memory_order_release);
                    showPopup                    = true;
                    const ULONGLONG admittedTick = task->_presentationDeadlineTick >= FileOperations::kTaskCardRevealDelayMs
                                                       ? task->_presentationDeadlineTick - FileOperations::kTaskCardRevealDelayMs
                                                       : 0u;
                    Debug::Perf::Emit(L"FileOps.TaskPresentation.RevealMs",
                                      L"non-clean-race",
                                      nowTick >= admittedTick ? nowTick - admittedTick : 0u,
                                      FileOperations::kTaskCardRevealDelayMs,
                                      0u,
                                      S_OK);
#ifdef ENABLE_TESTS
                    _debugTaskPresentedCount.fetch_add(1u, std::memory_order_relaxed);
#endif
                }
                continue;
            }

            if (! nextDeadline.has_value() || task->_presentationDeadlineTick < nextDeadline.value())
            {
                nextDeadline = task->_presentationDeadlineTick;
            }
        }
    }

    if (showPopup)
    {
        EnsurePopupVisible();
    }
    FileOperationsPopup::RefreshTelemetry(_popup.get());
    // x-lifetime#7: the completion-recovery safety net stays armed for as long as work exists, but
    // it wakes every 100 ms only while a fallback completion waits; otherwise once a second. A
    // payload-less completion wake re-arms the fast cadence (FolderWindow::OnFileOperationCompleted).
    const UINT heartbeatMs = HasPendingFallbackCompletions() ? kCompletionRecoveryHeartbeatMs : kCompletionRecoveryIdleHeartbeatMs;
    std::optional<UINT> next;
    if (nextDeadline.has_value())
    {
        const ULONGLONG remaining = nextDeadline.value() > nowTick ? nextDeadline.value() - nowTick : 1u;
        next                      = static_cast<UINT>((std::min<ULONGLONG>)(remaining, heartbeatMs));
    }
    else if (HasActiveOperations())
        next = heartbeatMs;
#ifdef ENABLE_TESTS
    _debugLastPresentationHeartbeatMs.store(next.value_or(0u), std::memory_order_release);
#endif
    return next;
}

void FolderWindow::FileOperationState::RevealPendingTasksImmediately() noexcept
{
    bool showPopup = false;
    {
        std::scoped_lock lock(_mutex);
        for (const auto& task : _tasks)
        {
            if (! task || task->_taskFinished.load(std::memory_order_acquire))
            {
                continue;
            }
            const Task::TaskPresentationState previous = task->_presentationState.exchange(Task::TaskPresentationState::Presented, std::memory_order_acq_rel);
            if (previous == Task::TaskPresentationState::Hidden || previous == Task::TaskPresentationState::RevealRequested)
            {
                showPopup = true;
#ifdef ENABLE_TESTS
                _debugTaskPresentedCount.fetch_add(1u, std::memory_order_relaxed);
#endif
            }
            else if (previous != Task::TaskPresentationState::Presented)
            {
                task->_presentationState.store(previous, std::memory_order_release);
            }
        }
    }
    if (showPopup)
    {
        EnsurePopupVisible();
    }
}

#ifdef ENABLE_TESTS
void FolderWindow::FileOperationState::DebugSetTaskPresentationNowTickForSelfTest(std::optional<ULONGLONG> nowTick) noexcept
{
    _debugTaskPresentationNowTick.store(nowTick.value_or(kTaskPresentationLiveClock), std::memory_order_release);
}

uint64_t FolderWindow::FileOperationState::DebugInlineRenameSilentCleanCompletionCount() const noexcept
{
    return _debugInlineRenameSilentCleanCompletionCount.load(std::memory_order_acquire);
}

uint64_t FolderWindow::FileOperationState::DebugTaskPresentedCount() const noexcept
{
    return _debugTaskPresentedCount.load(std::memory_order_acquire);
}

bool FolderWindow::FileOperationState::DebugValidateRetainedSourceActionForSelfTest(uint64_t taskId, IFileSystem* fileSystem) noexcept
{
    CompletedTaskSummary summary{};
    {
        std::scoped_lock lock(_mutex);
        const auto found = std::ranges::find_if(_completedTasks, [taskId](const CompletedTaskSummary& value) noexcept { return value.taskId == taskId; });
        if (found == _completedTasks.end())
        {
            return false;
        }
        summary = *found;
    }
    return summary.clipboardMoveAdmission && summary.operation == FILESYSTEM_MOVE && summary.retainedSourceCount > 0u && summary.unknownSourceCount == 0u &&
           summary.exactRetainedSourceItems.size() == summary.retainedSourceCount &&
           ValidateRetainedSourceActionItems(fileSystem, summary.exactRetainedSourceItems, nullptr);
}
#endif

uint64_t FolderWindow::FileOperationState::CreateOrUpdateInformationalTask(const FolderWindow::InformationalTaskUpdate& update) noexcept
{
    bool createdNew    = false;
    bool autoDismissed = false;
    bool needShowPopup = false;
    HWND popup         = nullptr;
    wil::unique_hwnd popupToClose;

    uint64_t taskId      = update.taskId;
    bool updatedExisting = false;

    const uint64_t lockWaitStartedUs = PerfNowUs();
    {
        std::scoped_lock lock(_mutex);
        const uint64_t lockWaitUs  = PerfElapsedUs(lockWaitStartedUs);
        const uint64_t lockHeldUs0 = PerfNowUs();
        if (taskId != 0)
        {
            for (auto& existing : _informationalTasks)
            {
                if (existing.taskId == taskId)
                {
                    existing        = update;
                    existing.taskId = taskId;
                    updatedExisting = true;
                    break;
                }
            }
        }

        if (taskId == 0 || ! updatedExisting)
        {
            taskId                                       = _nextTaskId++;
            FolderWindow::InformationalTaskUpdate stored = update;
            stored.taskId                                = taskId;
            _informationalTasks.push_back(std::move(stored));

            createdNew = true;
        }

        const bool autoDismissSuccess = _owner._settings ? GetAutoDismissSuccessFromSettings(*_owner._settings) : false;
        if (autoDismissSuccess && update.finished && IsAutoDismissableFileOperationCompletion(update.resultHr, 0, 0))
        {
            _informationalTasks.erase(std::remove_if(_informationalTasks.begin(),
                                                     _informationalTasks.end(),
                                                     [&](const FolderWindow::InformationalTaskUpdate& task) noexcept { return task.taskId == taskId; }),
                                      _informationalTasks.end());

            autoDismissed = true;
            if (_tasks.empty() && _completedTasks.empty() && _informationalTasks.empty())
            {
                popupToClose = std::move(_popup);
            }
        }

        popup = _popup.get();
        if (! autoDismissed && ! update.finished)
        {
            // Informational tasks should surface progress similarly to file operations: when the task starts, ensure the popup is shown.
            needShowPopup = createdNew || popup == nullptr || IsWindowVisible(popup) == FALSE || IsIconic(popup) != FALSE;
        }
        else
        {
            needShowPopup = createdNew && (popup == nullptr) && ! autoDismissed;
        }

        if (Debug::Perf::IsCaptureEnabled())
        {
            Debug::Perf::Emit(L"FileOps.InfoTask.Update.LockWaitUs", L"", lockWaitUs, static_cast<uint64_t>(_informationalTasks.size()), 0u, S_OK);
            Debug::Perf::Emit(
                L"FileOps.InfoTask.Update.LockHoldUs", L"", PerfElapsedUs(lockHeldUs0), static_cast<uint64_t>(_informationalTasks.size()), 0u, S_OK);
        }
    }

    if (needShowPopup)
    {
        EnsurePopupVisible();
    }
    else if (popup)
    {
        InvalidateRect(popup, nullptr, FALSE);
    }

    FileOperationsPopup::RefreshTelemetry(_popup.get());
    return taskId;
}

void FolderWindow::FileOperationState::DismissInformationalTask(uint64_t taskId) noexcept
{
    if (taskId == 0)
    {
        return;
    }

    wil::unique_hwnd popupToClose;
    HWND popup = nullptr;
    {
        std::scoped_lock lock(_mutex);
        _informationalTasks.erase(std::remove_if(_informationalTasks.begin(),
                                                 _informationalTasks.end(),
                                                 [&](const FolderWindow::InformationalTaskUpdate& task) noexcept { return task.taskId == taskId; }),
                                  _informationalTasks.end());

        if (_tasks.empty() && _completedTasks.empty() && _informationalTasks.empty())
        {
            popupToClose = std::move(_popup);
        }
        else
        {
            popup = _popup.get();
        }
    }

    if (popup)
    {
        InvalidateRect(popup, nullptr, FALSE);
    }
    FileOperationsPopup::RefreshTelemetry(popup);
}

bool FolderWindow::FileOperationState::GetAutoDismissSuccess() const noexcept
{
    if (! _owner._settings)
    {
        return false;
    }

    return GetAutoDismissSuccessFromSettings(*_owner._settings);
}

void FolderWindow::FileOperationState::SetAutoDismissSuccess(bool enabled) noexcept
{
    if (! _owner._settings)
    {
        return;
    }

    const bool previous = GetAutoDismissSuccessFromSettings(*_owner._settings);
    SetAutoDismissSuccessInSettings(*_owner._settings, enabled);

    if (previous != enabled)
    {
        QueueSettingsSave(L"auto-dismiss");
    }
}

bool FolderWindow::FileOperationState::GetPopupFooterOnly() const noexcept
{
    if (! _owner._settings)
    {
        return false;
    }

    return GetPopupFooterOnlyFromSettings(*_owner._settings);
}

void FolderWindow::FileOperationState::SetPopupFooterOnly(bool footerOnly) noexcept
{
    if (! _owner._settings)
    {
        return;
    }

    const bool previous = GetPopupFooterOnlyFromSettings(*_owner._settings);
    if (previous == footerOnly)
    {
        return;
    }

    SetPopupFooterOnlyInSettings(*_owner._settings, footerOnly);
    QueueSettingsSave(L"popup footer-only");
}

bool FolderWindow::FileOperationState::GetPopupCompactDensity() const noexcept
{
    if (! _owner._settings)
    {
        return false;
    }

    return GetPopupCompactDensityFromSettings(*_owner._settings);
}

bool FolderWindow::FileOperationState::GetPopupThroughputExpanded() const noexcept
{
    return ! _owner._settings || ! _owner._settings->fileOperations.has_value() || _owner._settings->fileOperations->popupThroughputExpanded;
}

void FolderWindow::FileOperationState::SetPopupThroughputExpanded(bool expanded) noexcept
{
    if (! _owner._settings || GetPopupThroughputExpanded() == expanded)
        return;
    auto& settings = *_owner._settings;
    if (! settings.fileOperations.has_value())
        settings.fileOperations.emplace();
    settings.fileOperations->popupThroughputExpanded = expanded;
    PruneFileOperationsSettingsIfDefault(settings);
    QueueSettingsSave(L"popup throughput disclosure");
}

void FolderWindow::FileOperationState::SetPopupCompactDensity(bool compactDensity) noexcept
{
    if (! _owner._settings)
    {
        return;
    }

    const bool previous = GetPopupCompactDensityFromSettings(*_owner._settings);
    SetPopupCompactDensityInSettings(*_owner._settings, compactDensity);
    if (previous == compactDensity)
    {
        return;
    }

    QueueSettingsSave(L"popup density");
}
