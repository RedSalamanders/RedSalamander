#include "CompareDirectoriesEngine.SelfTest.h"

#ifdef ENABLE_TESTS

#include "Framework.h"

#include "ConnectionProfileUtils.h"
#include "LocalSearchIndexCore.h"
#include "SearchFallbackEngine.h"
#include "SearchServiceBroker.h"
#include "SearchTextHelpers.h"
#include "SqliteIndexStore.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <format>
#include <iterator>
#include <limits>
#include <mutex>
#include <new>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <sqlite3.h>
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "CompareDirectoriesEngine.h"
#include "ConnectionSecrets.h"
#include "CrashHandler.h"
#include "CrashQuarantine.h"
#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "HostServices.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/Factory.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "SelfTestCommon.h"
#include "SessionState.h"
#include "SettingsStore.h"
#include "WindowsHello.h"

extern Common::Settings::Settings g_settings;

namespace
{
constexpr std::wstring_view kBuiltinLocalFileSystemId            = L"builtin/file-system";
constexpr std::wstring_view kBuiltinDummyFileSystemId            = L"builtin/file-system-dummy";
constexpr std::wstring_view kBuiltin7zFileSystemId               = L"builtin/file-system-7z";
constexpr std::wstring_view kBuiltinFtpFileSystemId              = L"builtin/file-system-ftp";
constexpr std::wstring_view kBuiltinGoogleDriveFileSystemId      = L"builtin/file-system-gdrive";
constexpr std::wstring_view kBuiltinS3FileSystemId               = L"builtin/file-system-s3";
constexpr std::wstring_view kBuiltinOneDrivePersonalFileSystemId = L"builtin/file-system-onedrive-personal";
constexpr std::wstring_view kBuiltinOneDriveBusinessFileSystemId = L"builtin/file-system-onedrive-business";
constexpr std::wstring_view kBuiltinSharePointFileSystemId       = L"builtin/file-system-sharepoint";

constexpr std::wstring_view kSelfTestEnvConnFtp              = L"REDSALAMANDER_SELFTEST_CONN_FTP";
constexpr std::wstring_view kSelfTestEnvConnS3               = L"REDSALAMANDER_SELFTEST_CONN_S3";
constexpr std::wstring_view kSelfTestEnvConnOneDrivePersonal = L"REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL";
constexpr std::wstring_view kSelfTestEnvConnOneDriveBusiness = L"REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_BUSINESS";
constexpr std::wstring_view kSelfTestEnvConnSharePoint       = L"REDSALAMANDER_SELFTEST_CONN_SHAREPOINT";

constexpr std::wstring_view kSelfTestDefaultConnFtp              = L"FileOpsSelfTest FTP";
constexpr std::wstring_view kSelfTestDefaultConnS3               = L"FileOpsSelfTest S3";
constexpr std::wstring_view kSelfTestDefaultConnOneDrivePersonal = L"FileOpsSelfTest OneDrive Personal";
constexpr std::wstring_view kSelfTestDefaultConnOneDriveBusiness = L"FileOpsSelfTest OneDrive Business";
constexpr std::wstring_view kSelfTestDefaultConnSharePoint       = L"FileOpsSelfTest SharePoint";
constexpr uint32_t kWarmIndexedFirstBatchBudgetMs                = 100u;

constexpr std::wstring_view kCompareCaseNames[] = {
    L"local_search_qi_and_capabilities",
    L"local_search_callback_contract",
    L"local_search_backend_preferences_roundtrip",
    L"local_search_scan_wide_tree_parallel_walk_name_only",
    L"local_search_name_wildcard_recursive",
    L"local_search_name_windows_filesystem_case_parity",
    L"local_search_content_literal",
    L"local_search_name_and_content_and_semantics",
    L"local_search_invalid_query_rejected",
    L"local_index_core_snapshot_reload",
    L"local_index_core_refs_probe_and_query_if_available",
    L"local_index_core_journal_replay_rename_delete_create",
    L"local_index_core_snapshot_corruption_rebuild",
    L"sqlite_index_store_bootstrap_creates_schema",
    L"sqlite_index_store_manual_compaction_reclaims_space",
    L"sqlite_index_store_automatic_checkpoint_truncates_wal",
    L"sqlite_index_store_automatic_compaction_is_bounded",
    L"sqlite_index_store_upgrade_paths",
    L"sqlite_index_store_load_and_apply_journal_delta",
    L"local_index_core_sqlite_option_keeps_snapshot_runtime_store",
    L"local_index_core_sqlite_cold_start_bypasses_snapshot_runtime_store",
    L"local_index_core_sqlite_authoritative_replays_without_snapshot_runtime_store",
    L"local_index_core_sqlite_authoritative_reseeds_after_store_loss_during_replay",
    L"local_index_core_sqlite_cold_start_stale_root_refreshes_before_query",
    L"local_index_core_sqlite_direct_query_freshness_policy",
    L"local_index_core_sqlite_cutover_blocked_by_pending_legacy_import",
    L"local_index_core_sqlite_sidecar_imports_legacy_snapshot",
    L"local_index_core_sqlite_prefilter_classifies_name_patterns",
    L"local_search_native_matches_host_fallback",
    L"local_search_native_unicode_long_path_matches_host_fallback",
    L"local_search_scan_follow_symlink_loop_guard",
    L"search_service_help_lists_cli_options",
    L"search_service_compact_cli_runs_manual_sqlite_maintenance",
    L"search_service_compact_request_roundtrip",
    L"search_service_sqlite_status_reports_maintenance_history",
    L"search_service_sqlite_idle_maintenance_queue_and_completion",
    L"search_service_sqlite_delete_burst_maintenance_preserves_query_parity",
    L"search_service_sqlite_bootstrap_status_roundtrip",
    L"search_service_sqlite_cold_start_bypasses_snapshot_runtime_store",
    L"search_service_sqlite_ntfs_traversal_seed_stays_degraded",
    L"search_service_sqlite_cold_start_stale_root_refreshes_before_query",
    L"search_service_sqlite_invalid_store_falls_back_live_scan",
    L"search_service_sqlite_query_failure_falls_back_live_scan",
    L"search_service_sqlite_midquery_failure_restarts_live_scan_without_duplicates",
    L"search_service_sqlite_prefilter_roundtrip",
    L"search_service_sqlite_status_reports_pending_legacy_import",
    L"search_service_sqlite_startup_warmup_failure_status_roundtrip",
    L"search_service_binary_uses_console_subsystem",
    L"search_service_default_identity_matches_build",
    L"search_service_sqlite_default_store_uses_build_specific_programdata_root",
    L"search_service_sqlite_seeded_default_store_reuses_build_specific_programdata_root",
    L"search_service_discovers_fixed_local_roots_on_start",
    L"search_service_sqlite_startup_warms_overridden_roots",
    L"search_service_foreground_rejects_second_instance",
    L"search_service_foreground_logs_request_status",
    L"search_service_status_and_query_roundtrip",
    L"search_service_query_reports_live_progress",
    L"local_search_service_indexed_name_latency_and_parity",
    L"local_search_service_matches_host_fallback",
    L"local_search_service_single_request_uses_query",
    L"local_search_service_content_early_stop_stats",
    L"local_search_service_protocol_mismatch_falls_back_local_index",
    L"local_search_service_disconnect_falls_back_local_index",
    L"search_service_multi_client_and_rebuild_control",
    L"search_text_helpers_decoding_and_binary",
    L"search_text_helpers_chunk_overlap_literal_and_regex",
    L"host_fallback_search_local_plugin_path_root",
    L"host_fallback_search_content_degraded_without_io",
    L"host_fallback_search_access_denied_warning",
    L"host_fallback_search_short_read_and_cancel",
    L"host_fallback_search_dummy_name_only",
    L"host_fallback_search_7z_name_only",
    L"host_fallback_search_remote_ftp_name_only",
    L"windows_hello_cache",
    L"oauth_refresh_token_storage",
    L"oauth_authmode_roundtrip",
    L"google_drive_plugin_contract",
    L"google_drive_navigation_menu_callback_clear_drains",
    L"google_drive_cleared_client_id_requires_configuration",
    L"google_drive_connection_requires_refresh_token",
    L"onedrive_personal_cleared_client_id_requires_configuration",
    L"unique",
    L"typemismatch",
    L"size",
    L"time",
    L"attributes",
    L"content",
    L"content_dual_io",
    L"content_no_io_disables_compareContent",
    L"content_size_mismatch_no_pending",
    L"zero_vs_nonzero_content",
    L"unicode_filenames",
    L"content short reads",
    L"subdir pending",
    L"subdirs",
    L"no_sync_deep_scan",
    L"subdirattrs",
    L"missing folder",
    L"reparse",
    L"dummy_content",
    L"deep_tree",
    L"invalidate",
    L"concurrent_get_or_compute_decision",
    L"empty_directories",
    L"ignore",
    L"ignore_multiple_patterns",
    L"ignore_pattern_length_cap",
    L"ignore_pattern_count_cap",
    L"ignore_wildcard_pathology_runtime_bound",
    L"crash_quarantine_synthetic_marker",
    L"showIdentical",
    L"content_pending_elided",
    L"setCompareEnabled",
    L"invalidateForPath",
    L"decisionUpdatedCallback",
    L"uiVersion",
    L"accessors",
    L"plugin_path_math",
    L"connection_display_url",
    L"try_make_relative_outside_root",
    L"baseInterfaces",
    L"contentCacheHit",
    L"root_decision_empty_directories",
    L"zeroByteContent",
    L"setSettingsInvalidates",
    L"dircache_not_polluted_by_compare_scan",
    L"content_queue_bounded_hi_lo",
    L"decision_cache_eviction_budget_pins_visible",
    L"cancel_completes_bounded",
    L"invalid_directory_entry_buffer",
    L"scan_inflight_stamp_guards_restart",
    L"content_inflight_stamp_guards_restart",
    L"directory_size_local_callback_contract",
    L"directory_size_dummy_callback_contract",
    L"directory_size_7z_callback_contract",
    L"remote_file_s3",
    L"remote_s3_directory_size_callback_contract",
    L"remote_s3_pagination",
    L"remote_file_onedrive_personal",
    L"remote_onedrive_personal_directory_size_callback_contract",
    L"remote_file_onedrive_business",
    L"remote_onedrive_business_directory_size_callback_contract",
    L"remote_file_sharepoint",
    L"remote_sharepoint_directory_size_callback_contract",
    L"remote_file_ftp",
    L"remote_ftp_directory_size_callback_contract",
    L"remote_ftp_continue_on_error_partial",
    L"remote_s3_metadata_smoke",
    L"remote_s3_delete_missing",
};

std::atomic<uint32_t> g_windowsHelloVerifierCalls{0};

HRESULT TestWindowsHelloVerifier(HWND /*ownerWindow*/, std::wstring_view /*message*/) noexcept
{
    g_windowsHelloVerifierCalls.fetch_add(1u, std::memory_order_relaxed);
    return S_OK;
}

void Trace(std::wstring_view message) noexcept
{
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::CompareDirectories, message);
    SelfTest::AppendSelfTestTrace(message);
}

void AppendCompareSelfTestTraceLine(std::wstring_view message) noexcept;

void AppendCaseResult(SelfTest::SelfTestSuiteResult& suite,
                      std::wstring_view name,
                      SelfTest::SelfTestCaseResult::Status status,
                      std::wstring_view reason = {}) noexcept
{
    std::wstring caseLine;
    caseLine.reserve(6 + name.size());
    caseLine.append(L"Case: ");
    caseLine.append(name);
    Trace(caseLine);

    SelfTest::AppendCaseResult(suite, name, status, reason, 0);
}

[[nodiscard]] bool EqualsIgnoreCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        if (std::towlower(a[i]) != std::towlower(b[i]))
        {
            return false;
        }
    }

    return true;
}

void AppendSkippedCompareCasesForSetupFailure(const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult& suite, std::wstring_view reason) noexcept
{
    for (const std::wstring_view name : kCompareCaseNames)
    {
        if (! SelfTest::CaseFilterMatches(options.caseFilter, name))
        {
            continue;
        }

        AppendCaseResult(suite, name, SelfTest::SelfTestCaseResult::Status::skipped, reason);
    }
}

[[nodiscard]] bool ContainsIgnoreCase(std::wstring_view text, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }

    if (text.size() < needle.size())
    {
        return false;
    }

    for (size_t i = 0; i + needle.size() <= text.size(); ++i)
    {
        bool match = true;
        for (size_t j = 0; j < needle.size(); ++j)
        {
            if (std::towlower(text[i + j]) != std::towlower(needle[j]))
            {
                match = false;
                break;
            }
        }

        if (match)
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::wstring TrimWhitespace(std::wstring_view text) noexcept
{
    size_t start = 0;
    while (start < text.size())
    {
        const wchar_t ch = text[start];
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }
        ++start;
    }

    size_t end = text.size();
    while (end > start)
    {
        const wchar_t ch = text[end - 1u];
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }
        --end;
    }

    return std::wstring(text.substr(start, end - start));
}

[[nodiscard]] std::wstring GetEnvVarTrimmed(std::wstring_view name) noexcept
{
    if (name.empty())
    {
        return {};
    }

    std::wstring key(name);
    const DWORD required = GetEnvironmentVariableW(key.c_str(), nullptr, 0);
    if (required == 0)
    {
        return {};
    }

    std::wstring value;
    value.resize(required);
    const DWORD written = GetEnvironmentVariableW(key.c_str(), value.data(), required);
    if (written == 0 || written >= required)
    {
        return {};
    }

    value.resize(written);
    return TrimWhitespace(value);
}

void SecureClearAndFreeSecret(wil::unique_cotaskmem_string& secret) noexcept
{
    if (wchar_t* text = secret.get())
    {
        const size_t len = wcslen(text);
        if (len > 0)
        {
            SecureZeroMemory(text, len * sizeof(wchar_t));
        }
    }
    secret.reset();
}

struct PhaseCheckResult
{
    SelfTest::SelfTestCaseResult::Status status = SelfTest::SelfTestCaseResult::Status::skipped;
    std::wstring reason;
};

struct ResolvedRemoteProfile
{
    const Common::Settings::ConnectionProfile* profile = nullptr;
    std::wstring profileName;
    bool usedFallback = false;
};

[[nodiscard]] ResolvedRemoteProfile ResolveRemoteConnectionProfile(std::wstring_view envVarName,
                                                                   std::wstring_view defaultProfileName,
                                                                   std::wstring_view expectedPluginId) noexcept
{
    ResolvedRemoteProfile resolved{};
    const std::wstring overrideName = GetEnvVarTrimmed(envVarName);
    resolved.profileName            = ! overrideName.empty() ? overrideName : std::wstring(defaultProfileName);
    resolved.profile                = ConnectionProfileUtils::FindConnectionProfileByName(&g_settings, resolved.profileName);
    if (resolved.profile || ! overrideName.empty() || ! g_settings.connections)
    {
        return resolved;
    }

    for (const Common::Settings::ConnectionProfile& profile : g_settings.connections->items)
    {
        if (profile.name.empty() || profile.pluginId.empty() || ! EqualsIgnoreCase(profile.pluginId, expectedPluginId))
        {
            continue;
        }

        bool selfTestNamed = false;
        for (size_t i = 0; i + 8u <= profile.name.size(); ++i)
        {
            if (std::towlower(profile.name[i + 0u]) == L's' && std::towlower(profile.name[i + 1u]) == L'e' && std::towlower(profile.name[i + 2u]) == L'l' &&
                std::towlower(profile.name[i + 3u]) == L'f' && std::towlower(profile.name[i + 4u]) == L't' && std::towlower(profile.name[i + 5u]) == L'e' &&
                std::towlower(profile.name[i + 6u]) == L's' && std::towlower(profile.name[i + 7u]) == L't')
            {
                selfTestNamed = true;
                break;
            }
        }

        if (! selfTestNamed)
        {
            continue;
        }

        resolved.profile      = &profile;
        resolved.profileName  = profile.name;
        resolved.usedFallback = true;
        break;
    }

    return resolved;
}

[[nodiscard]] PhaseCheckResult CheckRemoteConnectionSecret(std::wstring_view protocolLabel,
                                                           std::wstring_view envVarName,
                                                           std::wstring_view defaultProfileName,
                                                           std::wstring_view expectedPluginId) noexcept
{
    const ResolvedRemoteProfile resolved               = ResolveRemoteConnectionProfile(envVarName, defaultProfileName, expectedPluginId);
    const std::wstring& profileName                    = resolved.profileName;
    const Common::Settings::ConnectionProfile* profile = resolved.profile;
    if (! profile)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: connection profile not found (set {} or create '{}').", protocolLabel, envVarName, defaultProfileName)};
    }

    if (profile->pluginId.empty() || ! EqualsIgnoreCase(profile->pluginId, expectedPluginId))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile targets a different plugin.", protocolLabel)};
    }

    if (profile->authMode == Common::Settings::ConnectionAuthMode::Anonymous)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: authMode=anonymous (no secret needed).", protocolLabel)};
    }

    if (profile->id.empty())
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile is missing a stable id.", protocolLabel)};
    }

    const bool bypassHello = g_settings.connections ? g_settings.connections->bypassWindowsHello : false;
    if (profile->requireWindowsHello && ! bypassHello)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: requireWindowsHello=true (enable bypassWindowsHello or disable the profile flag for automation).", protocolLabel)};
    }

    if (! profile->savePassword)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: savePassword=false (secret is not persisted).", protocolLabel)};
    }

    HostConnectionSecretKind kind = HOST_CONNECTION_SECRET_PASSWORD;
    if (profile->authMode == Common::Settings::ConnectionAuthMode::SshKey)
    {
        kind = HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE;
    }
    else if (profile->authMode == Common::Settings::ConnectionAuthMode::OAuth2Pkce)
    {
        kind = HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN;
    }

    wil::com_ptr<IHostConnections> hostConnections;
    const HRESULT hrQI = GetHostServices()->QueryInterface(IID_PPV_ARGS(hostConnections.addressof()));
    if (FAILED(hrQI) || ! hostConnections)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::failed,
                .reason = std::format(L"{}: missing IHostConnections. hr=0x{:08X}", protocolLabel, static_cast<unsigned long>(hrQI))};
    }

    wil::unique_cotaskmem_string secret;
    const HRESULT hrSecret = hostConnections->GetConnectionSecret(profileName.c_str(), kind, nullptr, secret.put());
    if (hrSecret == HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
    {
        SecureClearAndFreeSecret(secret);
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: secret not found.", protocolLabel)};
    }

    if (FAILED(hrSecret))
    {
        SecureClearAndFreeSecret(secret);
        return {.status = SelfTest::SelfTestCaseResult::Status::failed,
                .reason = std::format(L"{}: GetConnectionSecret failed. hr=0x{:08X}", protocolLabel, static_cast<unsigned long>(hrSecret))};
    }

    SecureClearAndFreeSecret(secret);
    return {.status = SelfTest::SelfTestCaseResult::Status::passed};
}

[[nodiscard]] std::wstring NormalizePluginPathForSelfTest(std::wstring_view rawPath) noexcept
{
    std::wstring path = TrimWhitespace(rawPath);
    for (wchar_t& ch : path)
    {
        if (ch == L'\\')
        {
            ch = L'/';
        }
    }

    while (path.size() > 1u && path.back() == L'/')
    {
        path.pop_back();
    }

    return path;
}

[[nodiscard]] PhaseCheckResult CheckRemoteConnectionSandbox(std::wstring_view protocolLabel,
                                                            std::wstring_view envVarName,
                                                            std::wstring_view defaultProfileName,
                                                            std::wstring_view expectedPluginId) noexcept
{
    const ResolvedRemoteProfile resolved               = ResolveRemoteConnectionProfile(envVarName, defaultProfileName, expectedPluginId);
    const Common::Settings::ConnectionProfile* profile = resolved.profile;
    if (! profile)
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: connection profile not found (set {} or create '{}').", protocolLabel, envVarName, defaultProfileName)};
    }

    if (profile->pluginId.empty() || ! EqualsIgnoreCase(profile->pluginId, expectedPluginId))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped, .reason = std::format(L"{}: profile targets a different plugin.", protocolLabel)};
    }

    const std::wstring initialPath = NormalizePluginPathForSelfTest(profile->initialPath);
    if (initialPath.empty() || initialPath[0] != L'/')
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must be an absolute plugin path (starting with '/').", protocolLabel)};
    }

    if (initialPath == L"/")
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must point to a dedicated selftest folder/prefix (not '/').", protocolLabel)};
    }

    if (ContainsIgnoreCase(initialPath, L"/@conn:"))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must not include the host-reserved '/@conn:' prefix.", protocolLabel)};
    }

    size_t segmentCount = 0;
    std::wstring_view soleSegment;
    for (size_t i = 0; i < initialPath.size();)
    {
        while (i < initialPath.size() && initialPath[i] == L'/')
        {
            ++i;
        }

        const size_t start = i;
        while (i < initialPath.size() && initialPath[i] != L'/')
        {
            ++i;
        }

        if (i == start)
        {
            continue;
        }

        const std::wstring_view segment = std::wstring_view(initialPath).substr(start, i - start);
        if (segment == L"." || segment == L"..")
        {
            return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                    .reason = std::format(L"{}: HARD REQUIREMENT: initialPath must not contain '.' or '..' segments.", protocolLabel)};
        }

        ++segmentCount;
        if (segmentCount == 1u)
        {
            soleSegment = segment;
        }
    }

    const bool isS3 = EqualsIgnoreCase(expectedPluginId, kBuiltinS3FileSystemId);
    if (isS3 && segmentCount < 2)
    {
        const bool bucketRootIsSelfTestOnly = segmentCount == 1u && ContainsIgnoreCase(soleSegment, L"selftest");
        if (bucketRootIsSelfTestOnly)
        {
            return {.status = SelfTest::SelfTestCaseResult::Status::passed};
        }

        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason = std::format(
                    L"{}: HARD REQUIREMENT: initialPath must include a bucket and a dedicated selftest prefix (e.g. '/bucket/red-salamander-selftest'), "
                    L"or use a dedicated selftest bucket root whose bucket name includes 'selftest'.",
                    protocolLabel)};
    }

    if (! ContainsIgnoreCase(initialPath, L"selftest"))
    {
        return {.status = SelfTest::SelfTestCaseResult::Status::skipped,
                .reason =
                    std::format(L"{}: HARD REQUIREMENT: initialPath must include 'selftest' (case-insensitive) to prove it is test-only.", protocolLabel)};
    }

    return {.status = SelfTest::SelfTestCaseResult::Status::passed};
}

using CreateFactoryFunc = HRESULT(__stdcall*)(REFIID, const FactoryOptions*, IHost*, const wchar_t*, void**);

struct CreatedFileSystemInstance
{
    wil::unique_hmodule module;
    wil::com_ptr<IFileSystem> fileSystem;

    CreatedFileSystemInstance()                                            = default;
    CreatedFileSystemInstance(const CreatedFileSystemInstance&)            = delete;
    CreatedFileSystemInstance& operator=(const CreatedFileSystemInstance&) = delete;
    CreatedFileSystemInstance(CreatedFileSystemInstance&&)                 = default;
    CreatedFileSystemInstance& operator=(CreatedFileSystemInstance&&)      = default;
};

[[nodiscard]] const FileSystemPluginManager::PluginEntry* FindFileSystemPluginById(std::wstring_view pluginId) noexcept
{
    if (pluginId.empty())
    {
        return nullptr;
    }

    const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
    for (const FileSystemPluginManager::PluginEntry& entry : plugins)
    {
        if (entry.id.empty())
        {
            continue;
        }

        if (CompareStringOrdinal(entry.id.c_str(), -1, pluginId.data(), static_cast<int>(pluginId.size()), TRUE) == CSTR_EQUAL)
        {
            return &entry;
        }
    }

    return nullptr;
}

[[nodiscard]] HRESULT TryCreateFileSystemInstance(std::wstring_view pluginId, std::wstring_view instanceContext, CreatedFileSystemInstance& out) noexcept
{
    out = {};

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(pluginId);
    if (! entry || entry->id.empty() || entry->disabled || ! entry->loadable || entry->path.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    wil::unique_hmodule module(LoadLibraryExW(entry->path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
    if (! module)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(module.get(), "RedSalamanderCreate"));
#pragma warning(pop)
    if (! createFactory)
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

    wil::com_ptr<IFileSystem> fileSystem;
    const std::wstring requestedPluginId = entry->factoryPluginId.empty() ? entry->id : entry->factoryPluginId;
    if (requestedPluginId.empty())
    {
        return E_INVALIDARG;
    }
    const HRESULT createHr = createFactory(__uuidof(IFileSystem), &options, GetHostServices(), requestedPluginId.c_str(), fileSystem.put_void());

    if (FAILED(createHr) || ! fileSystem)
    {
        return createHr;
    }

    wil::com_ptr<IInformations> informations;
    const HRESULT qiInfos = fileSystem->QueryInterface(__uuidof(IInformations), informations.put_void());
    if (FAILED(qiInfos) || ! informations)
    {
        return qiInfos;
    }

    if (entry->informations)
    {
        const char* configuration = nullptr;
        static_cast<void>(entry->informations->GetConfiguration(&configuration));
        if (configuration && configuration[0] != '\0')
        {
            static_cast<void>(informations->SetConfiguration(configuration));
        }
    }

    if (! instanceContext.empty())
    {
        wil::com_ptr<IFileSystemInitialize> initializer;
        const HRESULT qiInit = fileSystem->QueryInterface(__uuidof(IFileSystemInitialize), initializer.put_void());
        if (FAILED(qiInit) || ! initializer)
        {
            return qiInit;
        }

        std::wstring contextText(instanceContext);
        const HRESULT initHr = initializer->Initialize(contextText.c_str(), nullptr);
        if (FAILED(initHr))
        {
            return initHr;
        }
    }

    out.module     = std::move(module);
    out.fileSystem = std::move(fileSystem);
    return S_OK;
}

[[nodiscard]] HRESULT TryCreateFileSystemInstanceWithHost(std::wstring_view pluginId,
                                                          IHost* host,
                                                          std::wstring_view instanceContext,
                                                          CreatedFileSystemInstance& out) noexcept
{
    out = {};

    const FileSystemPluginManager::PluginEntry* entry = FindFileSystemPluginById(pluginId);
    if (! entry || entry->id.empty() || entry->disabled || ! entry->loadable || entry->path.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    wil::unique_hmodule module(LoadLibraryExW(entry->path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH));
    if (! module)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createFactory = reinterpret_cast<CreateFactoryFunc>(GetProcAddress(module.get(), "RedSalamanderCreate"));
#pragma warning(pop)
    if (! createFactory)
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    FactoryOptions options{};
    options.debugLevel = DEBUG_LEVEL_NONE;

    wil::com_ptr<IFileSystem> fileSystem;
    const std::wstring requestedPluginId = entry->factoryPluginId.empty() ? entry->id : entry->factoryPluginId;
    if (requestedPluginId.empty())
    {
        return E_INVALIDARG;
    }
    const HRESULT createHr = createFactory(__uuidof(IFileSystem), &options, host, requestedPluginId.c_str(), fileSystem.put_void());

    if (FAILED(createHr) || ! fileSystem)
    {
        return createHr;
    }

    wil::com_ptr<IInformations> informations;
    const HRESULT qiInfos = fileSystem->QueryInterface(__uuidof(IInformations), informations.put_void());
    if (FAILED(qiInfos) || ! informations)
    {
        return qiInfos;
    }

    if (entry->informations)
    {
        const char* configuration = nullptr;
        static_cast<void>(entry->informations->GetConfiguration(&configuration));
        if (configuration && configuration[0] != '\0')
        {
            static_cast<void>(informations->SetConfiguration(configuration));
        }
    }

    if (! instanceContext.empty())
    {
        wil::com_ptr<IFileSystemInitialize> initializer;
        const HRESULT qiInit = fileSystem->QueryInterface(__uuidof(IFileSystemInitialize), initializer.put_void());
        if (FAILED(qiInit) || ! initializer)
        {
            return qiInit;
        }

        std::wstring contextText(instanceContext);
        const HRESULT initHr = initializer->Initialize(contextText.c_str(), nullptr);
        if (FAILED(initHr))
        {
            return initHr;
        }
    }

    out.module     = std::move(module);
    out.fileSystem = std::move(fileSystem);
    return S_OK;
}

class BlockingConnectionManagerHost final : public IHost, public IHostConnections
{
public:
    BlockingConnectionManagerHost()                                                = default;
    BlockingConnectionManagerHost(const BlockingConnectionManagerHost&)            = delete;
    BlockingConnectionManagerHost(BlockingConnectionManagerHost&&)                 = delete;
    BlockingConnectionManagerHost& operator=(const BlockingConnectionManagerHost&) = delete;
    BlockingConnectionManagerHost& operator=(BlockingConnectionManagerHost&&)      = delete;
    ~BlockingConnectionManagerHost()                                               = default;

    void PrepareShowConnectionManagerResult(std::wstring connectionName, HRESULT result = S_OK) noexcept
    {
        std::lock_guard lock(_mutex);
        _connectionName = std::move(connectionName);
        _result         = result;
        _entered        = false;
        _releaseShow    = false;
    }

    [[nodiscard]] bool WaitForShowConnectionManagerCall(std::chrono::milliseconds timeout) noexcept
    {
        std::unique_lock lock(_mutex);
        return _cv.wait_for(lock, timeout, [this]() noexcept { return _entered; });
    }

    void ReleaseShowConnectionManager() noexcept
    {
        {
            std::lock_guard lock(_mutex);
            _releaseShow = true;
        }
        _cv.notify_all();
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IHost))
        {
            *ppvObject = static_cast<IHost*>(this);
        }
        else if (riid == __uuidof(IHostConnections))
        {
            *ppvObject = static_cast<IHostConnections*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return static_cast<ULONG>(++_refCount);
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG refCount = static_cast<ULONG>(--_refCount);
        if (refCount == 0)
        {
            delete this;
        }
        return refCount;
    }

    HRESULT STDMETHODCALLTYPE ShowConnectionManager(const HostConnectionManagerRequest* request, HostConnectionManagerResult* result) noexcept override
    {
        if (! request || ! result)
        {
            return E_POINTER;
        }

        result->version        = 1;
        result->sizeBytes      = sizeof(HostConnectionManagerResult);
        result->connectionName = nullptr;

        std::wstring connectionName;
        HRESULT showResult = E_FAIL;
        {
            std::unique_lock lock(_mutex);
            _entered = true;
            _cv.notify_all();
            _cv.wait(lock, [this]() noexcept { return _releaseShow; });
            _releaseShow   = false;
            connectionName = _connectionName;
            showResult     = _result;
        }

        if (showResult != S_OK)
        {
            return showResult;
        }

        const size_t bytes = (connectionName.size() + 1u) * sizeof(wchar_t);
        auto* allocated    = static_cast<wchar_t*>(CoTaskMemAlloc(bytes));
        if (! allocated)
        {
            return E_OUTOFMEMORY;
        }

        memcpy(allocated, connectionName.c_str(), bytes);
        result->connectionName = allocated;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetConnectionJsonUtf8(const wchar_t*, char**) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    HRESULT STDMETHODCALLTYPE GetConnectionSecret(const wchar_t*, HostConnectionSecretKind, HWND, wchar_t**) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    HRESULT STDMETHODCALLTYPE PromptForConnectionSecret(const wchar_t*, HostConnectionSecretKind, HWND, wchar_t**) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE ClearCachedConnectionSecret(const wchar_t*, HostConnectionSecretKind) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE UpgradeFtpAnonymousToPassword(const wchar_t*, HWND) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE SetConnectionSecret(const wchar_t*, HostConnectionSecretKind, const wchar_t*, BOOL) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    HRESULT STDMETHODCALLTYPE DeleteConnectionSecret(const wchar_t*, HostConnectionSecretKind, BOOL) noexcept override
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

private:
    std::atomic_ulong _refCount{1};
    std::mutex _mutex;
    std::condition_variable _cv;
    std::wstring _connectionName;
    HRESULT _result   = S_OK;
    bool _entered     = false;
    bool _releaseShow = false;
};

struct NavigationMenuCallbackProbe final : INavigationMenuCallback
{
    NavigationMenuCallbackProbe()                                              = default;
    NavigationMenuCallbackProbe(const NavigationMenuCallbackProbe&)            = delete;
    NavigationMenuCallbackProbe(NavigationMenuCallbackProbe&&)                 = delete;
    NavigationMenuCallbackProbe& operator=(const NavigationMenuCallbackProbe&) = delete;
    NavigationMenuCallbackProbe& operator=(NavigationMenuCallbackProbe&&)      = delete;
    ~NavigationMenuCallbackProbe()                                             = default;

    HRESULT STDMETHODCALLTYPE NavigationMenuRequestNavigate(const wchar_t* path, void* cookie) noexcept override
    {
        std::lock_guard lock(mutex);
        lastPath   = path ? path : L"";
        lastCookie = cookie;
        ++callCount;
        return S_OK;
    }

    [[nodiscard]] unsigned int GetCallCount() const noexcept
    {
        std::lock_guard lock(mutex);
        return callCount;
    }

    [[nodiscard]] std::wstring GetLastPath() const
    {
        std::lock_guard lock(mutex);
        return lastPath;
    }

    [[nodiscard]] void* GetLastCookie() const noexcept
    {
        std::lock_guard lock(mutex);
        return lastCookie;
    }

    void Reset() noexcept
    {
        std::lock_guard lock(mutex);
        callCount = 0;
        lastPath.clear();
        lastCookie = nullptr;
    }

    mutable std::mutex mutex;
    unsigned int callCount = 0;
    std::wstring lastPath;
    void* lastCookie = nullptr;
};

[[nodiscard]] std::wstring MakeGuidText() noexcept
{
    GUID guid{};
    if (FAILED(::CoCreateGuid(&guid)))
    {
        return {};
    }

    wchar_t buffer[64]{};
    if (::StringFromGUID2(guid, buffer, static_cast<int>(std::size(buffer))) <= 0)
    {
        return {};
    }

    std::wstring text(buffer);
    text.erase(std::remove_if(text.begin(), text.end(), [](wchar_t ch) noexcept { return ch == L'{' || ch == L'}'; }), text.end());
    return text;
}

[[nodiscard]] bool SetFileLastWriteTime(const std::filesystem::path& path, const FILETIME& lastWriteTime) noexcept
{
    wil::unique_handle file(::CreateFileW(
        path.c_str(), FILE_WRITE_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return false;
    }

    return ::SetFileTime(file.get(), nullptr, nullptr, &lastWriteTime) != 0;
}

struct ObservedFileMetadata final
{
    unsigned long attributes = 0u;
    int64_t creationTime     = 0;
    int64_t lastAccessTime   = 0;
    int64_t lastWriteTime    = 0;
    int64_t changeTime       = 0;
    int64_t endOfFile        = 0;
    int64_t allocationSize   = 0;
};

[[nodiscard]] bool TryReadObservedFileMetadata(const std::filesystem::path& path, ObservedFileMetadata& outMetadata) noexcept
{
    outMetadata = {};

    wil::unique_handle file(::CreateFileW(path.c_str(),
                                          FILE_READ_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_FLAG_BACKUP_SEMANTICS,
                                          nullptr));
    if (! file)
    {
        return false;
    }

    FILE_BASIC_INFO basic{};
    if (::GetFileInformationByHandleEx(file.get(), FileBasicInfo, &basic, sizeof(basic)) == 0)
    {
        return false;
    }

    FILE_STANDARD_INFO standard{};
    if (::GetFileInformationByHandleEx(file.get(), FileStandardInfo, &standard, sizeof(standard)) == 0)
    {
        return false;
    }

    outMetadata.attributes     = (basic.FileAttributes != 0u) ? basic.FileAttributes : GetFileAttributesW(path.c_str());
    outMetadata.creationTime   = basic.CreationTime.QuadPart;
    outMetadata.lastAccessTime = basic.LastAccessTime.QuadPart;
    outMetadata.lastWriteTime  = basic.LastWriteTime.QuadPart;
    outMetadata.changeTime     = basic.ChangeTime.QuadPart;
    outMetadata.endOfFile      = standard.EndOfFile.QuadPart;
    outMetadata.allocationSize = standard.AllocationSize.QuadPart;
    return true;
}

[[nodiscard]] std::vector<SqliteIndexStore::ImportedEntry> BuildSyntheticSqliteEntries(const std::filesystem::path& rootPath, const size_t entryCount) noexcept
{
    constexpr uint64_t kBaseTimestamp = 133838640000000000ull;

    std::vector<SqliteIndexStore::ImportedEntry> entries;
    entries.reserve(entryCount);
    for (size_t index = 0; index < entryCount; ++index)
    {
        const std::wstring name              = std::format(L"entry_{:05}.txt", index);
        const std::filesystem::path fullPath = rootPath / name;
        const uint64_t timestamp             = kBaseTimestamp + static_cast<uint64_t>(index) * 10'000ull;
        entries.push_back({
            .fileIdLow           = static_cast<uint64_t>(index + 1u),
            .fileIdHigh          = 0u,
            .parentIdLow         = 0u,
            .parentIdHigh        = 0u,
            .fullPath            = fullPath.wstring(),
            .name                = name,
            .attributes          = FILE_ATTRIBUTE_ARCHIVE,
            .sizeBytes           = 4096u + static_cast<uint64_t>(index % 97u),
            .writeTime100ns      = timestamp,
            .creationTime100ns   = timestamp,
            .lastAccessTime100ns = timestamp,
            .changeTime100ns     = timestamp,
            .allocationSize      = 4096u,
        });
    }

    return entries;
}

[[nodiscard]] bool PrepareSqliteMaintenanceStore(const std::filesystem::path& caseRoot,
                                                 const std::filesystem::path& databasePath,
                                                 SqliteIndexStore::StoreInfo& outInfo,
                                                 std::wstring& outError,
                                                 const size_t largeEntryCount         = 2048u,
                                                 const size_t compactedEntryCount     = 24u,
                                                 const bool materializeCompactedFiles = false) noexcept
{
    outInfo = {};
    outError.clear();

    SqliteIndexStore::StoreInfo bootstrapInfo{};
    HRESULT hr = SqliteIndexStore::EnsureBootstrap(databasePath.wstring(), &bootstrapInfo);
    if (FAILED(hr))
    {
        outError = std::format(L"SqliteIndexStore::EnsureBootstrap failed. hr=0x{:08X}", static_cast<unsigned long>(hr));
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest largeRequest{};
    largeRequest.rootPath       = caseRoot.wstring();
    largeRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    largeRequest.journalId      = 11u;
    largeRequest.nextUsn        = 100u;
    largeRequest.state          = SqliteIndexStore::kVolumeStateReady;
    largeRequest.entries        = BuildSyntheticSqliteEntries(caseRoot, largeEntryCount);

    SqliteIndexStore::ReplaceVolumeResult replaceResult{};
    hr = SqliteIndexStore::ReplaceVolume(databasePath.wstring(), largeRequest, &replaceResult);
    if (FAILED(hr))
    {
        outError = std::format(L"Initial ReplaceVolume failed. hr=0x{:08X}", static_cast<unsigned long>(hr));
        return false;
    }

    SqliteIndexStore::ReplaceVolumeRequest compactedRequest{};
    compactedRequest.rootPath       = caseRoot.wstring();
    compactedRequest.fileSystemKind = LocalSearchIndexCore::FileSystemKind::Ntfs;
    compactedRequest.journalId      = 12u;
    compactedRequest.nextUsn        = 200u;
    compactedRequest.state          = SqliteIndexStore::kVolumeStateReady;
    compactedRequest.entries        = BuildSyntheticSqliteEntries(caseRoot, compactedEntryCount);

    hr = SqliteIndexStore::ReplaceVolume(databasePath.wstring(), compactedRequest, &replaceResult);
    if (FAILED(hr))
    {
        outError = std::format(L"Compacted ReplaceVolume failed. hr=0x{:08X}", static_cast<unsigned long>(hr));
        return false;
    }

    if (materializeCompactedFiles)
    {
        for (const auto& entry : compactedRequest.entries)
        {
            if (! SelfTest::WriteTextFile(entry.fullPath, "sqlite maintenance parity"))
            {
                outError = std::format(L"Failed to materialize '{}'.", entry.fullPath);
                return false;
            }
        }
    }

    hr = SqliteIndexStore::InspectStore(databasePath.wstring(), outInfo);
    if (FAILED(hr))
    {
        outError = std::format(L"SqliteIndexStore::InspectStore failed. hr=0x{:08X}", static_cast<unsigned long>(hr));
        return false;
    }

    if (outInfo.freelistPageCount == 0u)
    {
        outError = std::format(L"Expected a fragmented SQLite store before maintenance, but freelistPageCount was {}.", outInfo.freelistPageCount);
        return false;
    }

    return true;
}

[[nodiscard]] wil::com_ptr<IFileSystem> GetSelfTestFileSystem(std::wstring_view pluginId) noexcept
{
    wil::com_ptr<IFileSystem> fileSystem = SelfTest::GetFileSystem(pluginId);
    if (fileSystem)
    {
        return fileSystem;
    }

    FileSystemPluginManager& pluginManager = FileSystemPluginManager::GetInstance();
    if (FAILED(pluginManager.EnablePlugin(pluginId, g_settings)))
    {
        return {};
    }

    return SelfTest::GetFileSystem(pluginId);
}

[[nodiscard]] wil::com_ptr<IFileSystem> GetLocalFileSystem() noexcept
{
    return GetSelfTestFileSystem(kBuiltinLocalFileSystemId);
}

class ShortReadFileReader final : public IFileReader
{
public:
    ShortReadFileReader(wil::com_ptr<IFileReader> inner, unsigned long maxBytesPerRead, DWORD delayMs) noexcept
        : _inner(std::move(inner)),
          _maxBytesPerRead(std::max<unsigned long>(maxBytesPerRead, 1u)),
          _delayMs(delayMs)
    {
    }

    ShortReadFileReader(const ShortReadFileReader&)            = delete;
    ShortReadFileReader(ShortReadFileReader&&)                 = delete;
    ShortReadFileReader& operator=(const ShortReadFileReader&) = delete;
    ShortReadFileReader& operator=(ShortReadFileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (! _inner)
        {
            return E_FAIL;
        }
        return _inner->GetSize(sizeBytes);
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (! _inner)
        {
            return E_FAIL;
        }
        return _inner->Seek(offset, origin, newPosition);
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (bytesRead == nullptr)
        {
            return E_POINTER;
        }

        *bytesRead = 0;

        if (bytesToRead == 0)
        {
            return S_OK;
        }

        if (buffer == nullptr)
        {
            return E_POINTER;
        }

        if (! _inner)
        {
            return E_FAIL;
        }

        const unsigned long capped = std::min(bytesToRead, _maxBytesPerRead);
        if (_delayMs != 0)
        {
            ::Sleep(_delayMs);
        }
        return _inner->Read(buffer, capped, bytesRead);
    }

private:
    ~ShortReadFileReader() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileReader> _inner;
    unsigned long _maxBytesPerRead = 1;
    DWORD _delayMs                 = 0;
};

// ShortReadFileSystem wraps a real IFileSystem/IFileSystemIO and limits every Read()
// call to at most maxBytesPerRead bytes.  Used as a regression guard to verify that
// the content-comparison engine handles partial reads correctly (i.e. it never assumes
// a single Read() returns the full file).
class ShortReadFileSystem final : public IFileSystem, public IFileSystemIO
{
public:
    ShortReadFileSystem(wil::com_ptr<IFileSystem> base, std::filesystem::path shortReadRoot, unsigned long maxBytesPerRead, DWORD delayMs) noexcept
        : _base(std::move(base)),
          _shortReadRoot(std::move(shortReadRoot)),
          _maxBytesPerRead(std::max<unsigned long>(maxBytesPerRead, 1u)),
          _delayMs(delayMs)
    {
        if (_base)
        {
            static_cast<void>(_base->QueryInterface(__uuidof(IFileSystemIO), _baseIo.put_void()));
        }
    }

    ShortReadFileSystem(const ShortReadFileSystem&)            = delete;
    ShortReadFileSystem(ShortReadFileSystem&&)                 = delete;
    ShortReadFileSystem& operator=(const ShortReadFileSystem&) = delete;
    ShortReadFileSystem& operator=(ShortReadFileSystem&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        if (riid == __uuidof(IFileSystemIO))
        {
            *ppvObject = static_cast<IFileSystemIO*>(this);
            AddRef();
            return S_OK;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->ReadDirectoryInfo(path, ppFilesInformation);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->CopyItem(sourcePath, destinationPath, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->MoveItem(sourcePath, destinationPath, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->DeleteItem(path, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->RenameItem(sourcePath, destinationPath, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->CopyItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->MoveItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->DeleteItems(paths, count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->RenameItems(items, count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->GetCapabilities(jsonUtf8);
    }

    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                               FileSystemOperation operationType,
                                               FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->GetTransferHints(path, operationType, endpoint, hints);
    }

    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->GetStorageCharacteristics(path, characteristics);
    }

    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetAttributes(path, fileAttributes);
    }

    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override
    {
        if (reader == nullptr)
        {
            return E_POINTER;
        }
        *reader = nullptr;

        if (! _baseIo)
        {
            return E_POINTER;
        }

        wil::com_ptr<IFileReader> inner;
        const HRESULT hr = _baseIo->CreateFileReader(path, inner.put());
        if (FAILED(hr) || ! inner)
        {
            return FAILED(hr) ? hr : E_FAIL;
        }

        const std::wstring_view pathText(path ? path : L"");
        const std::wstring rootText = _shortReadRoot.wstring();
        const bool shouldShortRead  = ! rootText.empty() && OrdinalString::StartsWithNoCase(pathText, rootText);
        if (! shouldShortRead)
        {
            *reader = inner.detach();
            return S_OK;
        }

        auto* wrapper = new (std::nothrow) ShortReadFileReader(std::move(inner), _maxBytesPerRead, _delayMs);
        if (! wrapper)
        {
            return E_OUTOFMEMORY;
        }

        *reader = wrapper;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->CreateFileWriter(path, flags, writer);
    }

    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetFileBasicInformation(path, info);
    }

    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->SetFileBasicInformation(path, info);
    }

    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetItemProperties(path, jsonUtf8);
    }

private:
    ~ShortReadFileSystem() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    wil::com_ptr<IFileSystemIO> _baseIo;
    std::filesystem::path _shortReadRoot;
    unsigned long _maxBytesPerRead = 1;
    DWORD _delayMs                 = 0;
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreateShortReadFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                  const std::filesystem::path& shortReadRoot,
                                                                  unsigned long maxBytesPerRead,
                                                                  DWORD delayMs) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) ShortReadFileSystem(base, shortReadRoot, maxBytesPerRead, delayMs);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

struct ReadDirectoryTestBehavior
{
    std::filesystem::path targetPath;
    DWORD delayMs           = 0;
    bool returnMalformed    = false;
    unsigned long fakeCount = 1;
    HRESULT forcedHr        = S_OK;
};

[[nodiscard]] std::wstring NormalizePathForNoCaseCompare(std::filesystem::path value) noexcept
{
    value = value.lexically_normal();
    while (! value.empty() && ! value.has_filename() && value != value.root_path())
    {
        value = value.parent_path();
    }

    std::wstring text = value.wstring();
    std::replace(text.begin(), text.end(), L'/', L'\\');

    if (text.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        text.erase(0, 8);
        text.insert(0, L"\\\\");
    }
    else if (text.rfind(L"\\\\?\\", 0) == 0)
    {
        text.erase(0, 4);
    }

    if (! text.empty())
    {
        ::CharLowerBuffW(text.data(), static_cast<DWORD>(text.size()));
    }

    return text;
}

class RawFilesInformation final : public IFilesInformation
{
public:
    RawFilesInformation(std::vector<unsigned char> buffer, unsigned long count) noexcept : _buffer(std::move(buffer)), _count(count)
    {
    }

    RawFilesInformation(const RawFilesInformation&)            = delete;
    RawFilesInformation(RawFilesInformation&&)                 = delete;
    RawFilesInformation& operator=(const RawFilesInformation&) = delete;
    RawFilesInformation& operator=(RawFilesInformation&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFilesInformation))
        {
            *ppvObject = static_cast<IFilesInformation*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** ppFileInfo) noexcept override
    {
        if (ppFileInfo == nullptr)
        {
            return E_POINTER;
        }

        if (_buffer.empty())
        {
            *ppFileInfo = nullptr;
            return S_OK;
        }

        *ppFileInfo = reinterpret_cast<FileInfo*>(_buffer.data());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept override
    {
        if (pSize == nullptr)
        {
            return E_POINTER;
        }

        *pSize = static_cast<unsigned long>(_buffer.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept override
    {
        if (pSize == nullptr)
        {
            return E_POINTER;
        }

        *pSize = static_cast<unsigned long>(_buffer.capacity());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept override
    {
        if (pCount == nullptr)
        {
            return E_POINTER;
        }

        *pCount = _count;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept override
    {
        if (ppEntry == nullptr)
        {
            return E_POINTER;
        }

        *ppEntry = nullptr;
        if (index != 0 || _buffer.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_INDEX);
        }

        *ppEntry = reinterpret_cast<FileInfo*>(_buffer.data());
        return S_OK;
    }

private:
    ~RawFilesInformation() = default;

    std::atomic_ulong _refCount{1};
    std::vector<unsigned char> _buffer;
    unsigned long _count = 0;
};

[[nodiscard]] std::vector<unsigned char> BuildMalformedDirectoryBuffer(std::wstring_view name) noexcept
{
    const size_t nameBytes = name.size() * sizeof(wchar_t);
    const size_t totalSize = offsetof(FileInfo, FileName) + nameBytes;

    std::vector<unsigned char> buffer(totalSize);
    if (buffer.empty())
    {
        return buffer;
    }

    auto* entry            = reinterpret_cast<FileInfo*>(buffer.data());
    entry->NextEntryOffset = static_cast<unsigned long>(sizeof(FileInfo) + 16u); // Intentionally larger than the backing buffer.
    entry->FileIndex       = 1;
    entry->CreationTime    = 0;
    entry->LastAccessTime  = 0;
    entry->LastWriteTime   = 0;
    entry->ChangeTime      = 0;
    entry->EndOfFile       = 1;
    entry->AllocationSize  = 1;
    entry->FileAttributes  = FILE_ATTRIBUTE_ARCHIVE;
    entry->FileNameSize    = static_cast<unsigned long>(nameBytes);
    entry->EaSize          = 0;

    if (nameBytes != 0)
    {
        std::memcpy(entry->FileName, name.data(), nameBytes);
    }

    return buffer;
}

class ReadDirectoryBehaviorFileSystem final : public IFileSystem
{
public:
    ReadDirectoryBehaviorFileSystem(wil::com_ptr<IFileSystem> base, ReadDirectoryTestBehavior behavior) noexcept
        : _base(std::move(base)),
          _behavior(std::move(behavior))
    {
        _normalizedTargetPath = NormalizePathForNoCaseCompare(_behavior.targetPath);
    }

    ReadDirectoryBehaviorFileSystem(const ReadDirectoryBehaviorFileSystem&)            = delete;
    ReadDirectoryBehaviorFileSystem(ReadDirectoryBehaviorFileSystem&&)                 = delete;
    ReadDirectoryBehaviorFileSystem& operator=(const ReadDirectoryBehaviorFileSystem&) = delete;
    ReadDirectoryBehaviorFileSystem& operator=(ReadDirectoryBehaviorFileSystem&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (ppFilesInformation == nullptr)
        {
            return E_POINTER;
        }

        *ppFilesInformation = nullptr;
        if (! _base)
        {
            return E_POINTER;
        }

        const std::filesystem::path requestedPath(path ? path : L"");
        const std::wstring requested = NormalizePathForNoCaseCompare(requestedPath);
        const bool intercept         = _normalizedTargetPath.empty() || requested == _normalizedTargetPath;
        if (! intercept)
        {
            return _base->ReadDirectoryInfo(path, ppFilesInformation);
        }

        if (_behavior.delayMs != 0u)
        {
            ::Sleep(_behavior.delayMs);
        }

        if (FAILED(_behavior.forcedHr))
        {
            return _behavior.forcedHr;
        }

        if (! _behavior.returnMalformed)
        {
            return _base->ReadDirectoryInfo(path, ppFilesInformation);
        }

        auto* info = new (std::nothrow) RawFilesInformation(BuildMalformedDirectoryBuffer(L"corrupt-entry.txt"), _behavior.fakeCount);
        if (! info)
        {
            return E_OUTOFMEMORY;
        }

        *ppFilesInformation = info;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->CopyItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->MoveItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* itemPath, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        return _base ? _base->DeleteItem(itemPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _base ? _base->RenameItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _base ? _base->CopyItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _base ? _base->MoveItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _base ? _base->DeleteItems(paths, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _base ? _base->RenameItems(items, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        return _base ? _base->GetCapabilities(jsonUtf8) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                               FileSystemOperation operationType,
                                               FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override
    {
        return _base ? _base->GetTransferHints(path, operationType, endpoint, hints) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept override
    {
        return _base ? _base->GetStorageCharacteristics(path, characteristics) : E_POINTER;
    }

private:
    ~ReadDirectoryBehaviorFileSystem() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    ReadDirectoryTestBehavior _behavior;
    std::wstring _normalizedTargetPath;
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreateReadDirectoryBehaviorFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                              ReadDirectoryTestBehavior behavior) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) ReadDirectoryBehaviorFileSystem(base, std::move(behavior));
    if (! wrapper)
    {
        return {};
    }

    wrapped.attach(wrapper);
    return wrapped;
}

class PluginPathMappedRootFileSystem final : public IFileSystem, public IFileSystemIO
{
public:
    PluginPathMappedRootFileSystem(wil::com_ptr<IFileSystem> base, std::filesystem::path windowsRoot) noexcept
        : _base(std::move(base)),
          _windowsRoot(std::move(windowsRoot))
    {
        if (_base)
        {
            static_cast<void>(_base->QueryInterface(__uuidof(IFileSystemIO), _baseIo.put_void()));
        }
    }

    PluginPathMappedRootFileSystem(const PluginPathMappedRootFileSystem&)            = delete;
    PluginPathMappedRootFileSystem(PluginPathMappedRootFileSystem&&)                 = delete;
    PluginPathMappedRootFileSystem& operator=(const PluginPathMappedRootFileSystem&) = delete;
    PluginPathMappedRootFileSystem& operator=(PluginPathMappedRootFileSystem&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        // Do not expose IInformations: DirectoryInfoCache uses it to decide path semantics and would
        // otherwise treat this wrapper as the file plugin, breaking plugin-path mapping.
        if (riid == __uuidof(IInformations))
        {
            *ppvObject = nullptr;
            return E_NOINTERFACE;
        }

        if (riid == __uuidof(IFileSystemIO))
        {
            if (! _baseIo)
            {
                *ppvObject = nullptr;
                return E_NOINTERFACE;
            }

            *ppvObject = static_cast<IFileSystemIO*>(this);
            AddRef();
            return S_OK;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    // IFileSystem
    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->ReadDirectoryInfo(MapPath(path).c_str(), ppFilesInformation);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->CopyItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->MoveItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        return _base ? _base->DeleteItem(MapPath(path).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _base ? _base->RenameItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(sourcePaths ? sourcePaths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        const std::filesystem::path mappedDestination = MapPath(destinationFolder);
        return _base->CopyItems(mappedPtrs.data(), count, mappedDestination.c_str(), flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(sourcePaths ? sourcePaths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        const std::filesystem::path mappedDestination = MapPath(destinationFolder);
        return _base->MoveItems(mappedPtrs.data(), count, mappedDestination.c_str(), flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(paths ? paths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        return _base->DeleteItems(mappedPtrs.data(), count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mappedSources;
        std::vector<FileSystemRenamePair> mappedPairs;
        mappedSources.reserve(count);
        mappedPairs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            const FileSystemRenamePair& src = items ? items[i] : FileSystemRenamePair{};
            mappedSources.emplace_back(MapPath(src.sourcePath).wstring());

            FileSystemRenamePair pair{};
            pair.sizeBytes  = sizeof(FileSystemRenamePair);
            pair.sourcePath = mappedSources.back().c_str();
            pair.newName    = src.newName;
            mappedPairs.emplace_back(pair);
        }

        return _base->RenameItems(mappedPairs.data(), count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        return _base ? _base->GetCapabilities(jsonUtf8) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                               FileSystemOperation operationType,
                                               FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override
    {
        return _base ? _base->GetTransferHints(MapPath(path).c_str(), operationType, endpoint, hints) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept override
    {
        return _base ? _base->GetStorageCharacteristics(MapPath(path).c_str(), characteristics) : E_POINTER;
    }

    // IFileSystemIO
    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetAttributes(MapPath(path).c_str(), fileAttributes);
    }

    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->CreateFileReader(MapPath(path).c_str(), reader);
    }

    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->CreateFileWriter(MapPath(path).c_str(), flags, writer);
    }

    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetFileBasicInformation(MapPath(path).c_str(), info);
    }

    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->SetFileBasicInformation(MapPath(path).c_str(), info);
    }

    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override
    {
        if (! _baseIo)
        {
            return E_POINTER;
        }
        return _baseIo->GetItemProperties(MapPath(path).c_str(), jsonUtf8);
    }

private:
    ~PluginPathMappedRootFileSystem() = default;

    [[nodiscard]] std::filesystem::path MapPath(const wchar_t* path) const noexcept
    {
        if (! path || path[0] == L'\0')
        {
            return _windowsRoot;
        }

        std::wstring text(path);
        std::replace(text.begin(), text.end(), L'\\', L'/');

        const size_t start = text.find_first_not_of(L"/");
        if (start == std::wstring::npos)
        {
            return _windowsRoot;
        }
        text.erase(0, start);
        if (text.empty())
        {
            return _windowsRoot;
        }

        std::replace(text.begin(), text.end(), L'/', L'\\');
        return (_windowsRoot / std::filesystem::path(text)).lexically_normal();
    }

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    wil::com_ptr<IFileSystemIO> _baseIo;
    std::filesystem::path _windowsRoot;
};

class PluginPathMappedRootFileSystemNoIO final : public IFileSystem
{
public:
    PluginPathMappedRootFileSystemNoIO(wil::com_ptr<IFileSystem> base, std::filesystem::path windowsRoot) noexcept
        : _base(std::move(base)),
          _windowsRoot(std::move(windowsRoot))
    {
    }

    PluginPathMappedRootFileSystemNoIO(const PluginPathMappedRootFileSystemNoIO&)            = delete;
    PluginPathMappedRootFileSystemNoIO(PluginPathMappedRootFileSystemNoIO&&)                 = delete;
    PluginPathMappedRootFileSystemNoIO& operator=(const PluginPathMappedRootFileSystemNoIO&) = delete;
    PluginPathMappedRootFileSystemNoIO& operator=(PluginPathMappedRootFileSystemNoIO&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        // Do not expose IInformations: DirectoryInfoCache uses it to decide path semantics and would
        // otherwise treat this wrapper as the file plugin, breaking plugin-path mapping.
        if (riid == __uuidof(IInformations))
        {
            *ppvObject = nullptr;
            return E_NOINTERFACE;
        }

        if (riid == __uuidof(IFileSystemIO))
        {
            *ppvObject = nullptr;
            return E_NOINTERFACE;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    // IFileSystem
    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }
        return _base->ReadDirectoryInfo(MapPath(path).c_str(), ppFilesInformation);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->CopyItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->MoveItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        return _base ? _base->DeleteItem(MapPath(path).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _base ? _base->RenameItem(MapPath(sourcePath).c_str(), MapPath(destinationPath).c_str(), flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(sourcePaths ? sourcePaths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        const std::filesystem::path mappedDestination = MapPath(destinationFolder);
        return _base->CopyItems(mappedPtrs.data(), count, mappedDestination.c_str(), flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(sourcePaths ? sourcePaths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        const std::filesystem::path mappedDestination = MapPath(destinationFolder);
        return _base->MoveItems(mappedPtrs.data(), count, mappedDestination.c_str(), flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mapped;
        std::vector<const wchar_t*> mappedPtrs;
        mapped.reserve(count);
        mappedPtrs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            mapped.emplace_back(MapPath(paths ? paths[i] : nullptr).wstring());
            mappedPtrs.emplace_back(mapped.back().c_str());
        }

        return _base->DeleteItems(mappedPtrs.data(), count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        std::vector<std::wstring> mappedSources;
        std::vector<FileSystemRenamePair> mappedPairs;
        mappedSources.reserve(count);
        mappedPairs.reserve(count);

        for (unsigned long i = 0; i < count; ++i)
        {
            const FileSystemRenamePair& src = items ? items[i] : FileSystemRenamePair{};
            mappedSources.emplace_back(MapPath(src.sourcePath).wstring());

            FileSystemRenamePair pair{};
            pair.sizeBytes  = sizeof(FileSystemRenamePair);
            pair.sourcePath = mappedSources.back().c_str();
            pair.newName    = src.newName;
            mappedPairs.emplace_back(pair);
        }

        return _base->RenameItems(mappedPairs.data(), count, flags, options, callback, cookie);
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        return _base ? _base->GetCapabilities(jsonUtf8) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                               FileSystemOperation operationType,
                                               FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override
    {
        return _base ? _base->GetTransferHints(MapPath(path).c_str(), operationType, endpoint, hints) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept override
    {
        return _base ? _base->GetStorageCharacteristics(MapPath(path).c_str(), characteristics) : E_POINTER;
    }

private:
    ~PluginPathMappedRootFileSystemNoIO() = default;

    [[nodiscard]] std::filesystem::path MapPath(const wchar_t* path) const noexcept
    {
        if (! path || path[0] == L'\0')
        {
            return _windowsRoot;
        }

        std::wstring text(path);
        std::replace(text.begin(), text.end(), L'\\', L'/');

        const size_t start = text.find_first_not_of(L"/");
        if (start == std::wstring::npos)
        {
            return _windowsRoot;
        }
        text.erase(0, start);
        if (text.empty())
        {
            return _windowsRoot;
        }

        std::replace(text.begin(), text.end(), L'/', L'\\');
        return (_windowsRoot / std::filesystem::path(text)).lexically_normal();
    }

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    std::filesystem::path _windowsRoot;
};

class CountingReadDirectoryFileSystem final : public IFileSystem
{
public:
    CountingReadDirectoryFileSystem(wil::com_ptr<IFileSystem> base, std::atomic_uint32_t* counter) noexcept : _base(std::move(base)), _counter(counter)
    {
    }

    CountingReadDirectoryFileSystem(const CountingReadDirectoryFileSystem&)            = delete;
    CountingReadDirectoryFileSystem(CountingReadDirectoryFileSystem&&)                 = delete;
    CountingReadDirectoryFileSystem& operator=(const CountingReadDirectoryFileSystem&) = delete;
    CountingReadDirectoryFileSystem& operator=(CountingReadDirectoryFileSystem&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
        {
            *ppvObject = static_cast<IFileSystem*>(this);
            AddRef();
            return S_OK;
        }

        if (_base)
        {
            return _base->QueryInterface(riid, ppvObject);
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    // IFileSystem
    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override
    {
        if (! _base)
        {
            return E_POINTER;
        }

        if (_counter)
        {
            static_cast<void>(_counter->fetch_add(1u, std::memory_order_relaxed));
        }

        return _base->ReadDirectoryInfo(path, ppFilesInformation);
    }

    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->CopyItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options,
                                       IFileSystemCallback* callback,
                                       void* cookie) noexcept override
    {
        return _base ? _base->MoveItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE
    DeleteItem(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept override
    {
        return _base ? _base->DeleteItem(path, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options,
                                         IFileSystemCallback* callback,
                                         void* cookie) noexcept override
    {
        return _base ? _base->RenameItem(sourcePath, destinationPath, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _base ? _base->CopyItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie) noexcept override
    {
        return _base ? _base->MoveItems(sourcePaths, count, destinationFolder, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _base ? _base->DeleteItems(paths, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options,
                                          IFileSystemCallback* callback,
                                          void* cookie) noexcept override
    {
        return _base ? _base->RenameItems(items, count, flags, options, callback, cookie) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override
    {
        return _base ? _base->GetCapabilities(jsonUtf8) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                               FileSystemOperation operationType,
                                               FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override
    {
        return _base ? _base->GetTransferHints(path, operationType, endpoint, hints) : E_POINTER;
    }

    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept override
    {
        return _base ? _base->GetStorageCharacteristics(path, characteristics) : E_POINTER;
    }

private:
    ~CountingReadDirectoryFileSystem() = default;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<IFileSystem> _base;
    std::atomic_uint32_t* _counter = nullptr;
};

[[nodiscard]] wil::com_ptr<IFileSystem> CreatePluginPathMappedRootFileSystem(const wil::com_ptr<IFileSystem>& base,
                                                                             const std::filesystem::path& windowsRoot) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) PluginPathMappedRootFileSystem(base, windowsRoot);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreatePluginPathMappedRootFileSystemNoIO(const wil::com_ptr<IFileSystem>& base,
                                                                                 const std::filesystem::path& windowsRoot) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) PluginPathMappedRootFileSystemNoIO(base, windowsRoot);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> CreateCountingReadDirectoryFileSystem(const wil::com_ptr<IFileSystem>& base, std::atomic_uint32_t* counter) noexcept
{
    wil::com_ptr<IFileSystem> wrapped;
    auto* wrapper = new (std::nothrow) CountingReadDirectoryFileSystem(base, counter);
    if (! wrapper)
    {
        return {};
    }
    wrapped.attach(wrapper);
    return wrapped;
}

[[nodiscard]] wil::com_ptr<IFileSystem> GetDummyFileSystem() noexcept
{
    return GetSelfTestFileSystem(kBuiltinDummyFileSystemId);
}

[[nodiscard]] bool CreateFileSystemIo(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IFileSystemIO>& outIo) noexcept
{
    outIo.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemIO), outIo.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outIo);
}

[[nodiscard]] bool CreateFileSystemSearch(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IFileSystemSearch>& outSearch) noexcept
{
    outSearch.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemSearch), outSearch.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outSearch);
}

[[nodiscard]] std::wstring CopySearchString(const wchar_t* text, unsigned long sizeBytes) noexcept
{
    if (text == nullptr)
    {
        return {};
    }

    if (sizeBytes != 0)
    {
        return std::wstring(text, sizeBytes / sizeof(wchar_t));
    }

    return std::wstring(text);
}

struct RecordedSearchMatch final
{
    std::wstring fullPath;
    std::wstring relativePath;
    std::wstring displayName;
    std::wstring previewText;
    int64_t creationTime     = 0;
    int64_t lastAccessTime   = 0;
    int64_t lastWriteTime    = 0;
    int64_t changeTime       = 0;
    int64_t endOfFile        = 0;
    int64_t allocationSize   = 0;
    uint32_t matchedBy       = 0;
    unsigned long attributes = 0;
};

struct RecordedSearchProgress final
{
    FileSystemSearchPhase phase                                 = FILESYSTEM_SEARCH_PHASE_INITIALIZING;
    FileSystemSearchBackend backend                             = FILESYSTEM_SEARCH_BACKEND_UNKNOWN;
    uint32_t warningFlags                                       = FILESYSTEM_SEARCH_WARNING_NONE;
    HRESULT statusHint                                          = S_OK;
    uint64_t scannedDirectories                                 = 0u;
    uint64_t scannedFiles                                       = 0u;
    uint64_t candidateFiles                                     = 0u;
    uint64_t matchedEntries                                     = 0u;
    LocalSearchIndexCore::StoreState storeState                 = LocalSearchIndexCore::StoreState::Unknown;
    LocalSearchIndexCore::SyncPhase syncPhase                   = LocalSearchIndexCore::SyncPhase::Idle;
    LocalSearchIndexCore::QueryExecutionMode queryExecutionMode = LocalSearchIndexCore::QueryExecutionMode::Unknown;
    LocalSearchIndexCore::FallbackReason fallbackReason         = LocalSearchIndexCore::FallbackReason::None;
    uint64_t completedRoots                                     = 0u;
    uint64_t totalRoots                                         = 0u;
    std::wstring activeRoot;
    std::wstring currentPath;
};

class BrokerProgressRecorder final
{
public:
    BrokerProgressRecorder()                                         = default;
    BrokerProgressRecorder(const BrokerProgressRecorder&)            = delete;
    BrokerProgressRecorder(BrokerProgressRecorder&&)                 = delete;
    BrokerProgressRecorder& operator=(const BrokerProgressRecorder&) = delete;
    BrokerProgressRecorder& operator=(BrokerProgressRecorder&&)      = delete;

    static HRESULT STDMETHODCALLTYPE ProgressThunk(const SearchServiceBroker::QueryProgress* progress, void* cookie) noexcept
    {
        if (progress == nullptr || cookie == nullptr)
        {
            return E_POINTER;
        }

        auto& recorder = *static_cast<BrokerProgressRecorder*>(cookie);
        RecordedSearchProgress copy{};
        copy.phase              = progress->phase;
        copy.backend            = FILESYSTEM_SEARCH_BACKEND_SERVICE;
        copy.warningFlags       = progress->warningFlags;
        copy.statusHint         = progress->statusHint;
        copy.scannedDirectories = progress->scannedDirectories;
        copy.scannedFiles       = progress->scannedFiles;
        copy.candidateFiles     = progress->candidateFiles;
        copy.matchedEntries     = progress->matchedEntries;
        copy.storeState         = progress->storeState;
        copy.syncPhase          = progress->syncPhase;
        copy.queryExecutionMode = progress->queryExecutionMode;
        copy.fallbackReason     = progress->fallbackReason;
        copy.completedRoots     = progress->completedRoots;
        copy.totalRoots         = progress->totalRoots;
        copy.activeRoot         = progress->activeRoot;
        copy.currentPath        = progress->currentPath;

        std::lock_guard lock(recorder._mutex);
        recorder._snapshots.push_back(std::move(copy));
        return S_OK;
    }

    [[nodiscard]] std::vector<RecordedSearchProgress> Snapshots() const noexcept
    {
        std::lock_guard lock(_mutex);
        return _snapshots;
    }

private:
    mutable std::mutex _mutex;
    std::vector<RecordedSearchProgress> _snapshots;
};

class BrokerBatchRecorder final
{
public:
    BrokerBatchRecorder() noexcept : _startedAt(std::chrono::steady_clock::now())
    {
    }

    BrokerBatchRecorder(const BrokerBatchRecorder&)            = delete;
    BrokerBatchRecorder(BrokerBatchRecorder&&)                 = delete;
    BrokerBatchRecorder& operator=(const BrokerBatchRecorder&) = delete;
    BrokerBatchRecorder& operator=(BrokerBatchRecorder&&)      = delete;

    static HRESULT STDMETHODCALLTYPE CandidateBatchThunk(LocalSearchIndexCore::Candidate* candidates,
                                                         size_t count,
                                                         size_t* consumedCount,
                                                         void* cookie) noexcept
    {
        if (consumedCount == nullptr || cookie == nullptr)
        {
            return E_POINTER;
        }

        if (count != 0u && candidates == nullptr)
        {
            return E_POINTER;
        }

        auto& recorder = *static_cast<BrokerBatchRecorder*>(cookie);

        std::lock_guard lock(recorder._mutex);
        if (! recorder._firstBatchElapsedMs.has_value())
        {
            const auto elapsed = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - recorder._startedAt).count();
            recorder._firstBatchElapsedMs =
                elapsed <= 0 ? 0u
                             : (elapsed >= static_cast<decltype(elapsed)>(std::numeric_limits<uint32_t>::max()) ? std::numeric_limits<uint32_t>::max()
                                                                                                                : static_cast<uint32_t>(elapsed));
        }

        recorder._batchesSeen += 1u;
        for (size_t index = 0; index < count; ++index)
        {
            recorder._candidates.push_back(candidates[index]);
        }

        *consumedCount = count;
        return S_OK;
    }

    [[nodiscard]] std::optional<uint32_t> FirstBatchElapsedMs() const noexcept
    {
        std::lock_guard lock(_mutex);
        return _firstBatchElapsedMs;
    }

    [[nodiscard]] uint32_t BatchesSeen() const noexcept
    {
        std::lock_guard lock(_mutex);
        return _batchesSeen;
    }

    [[nodiscard]] std::vector<LocalSearchIndexCore::Candidate> Candidates() const noexcept
    {
        std::lock_guard lock(_mutex);
        return _candidates;
    }

private:
    const std::chrono::steady_clock::time_point _startedAt;
    mutable std::mutex _mutex;
    std::optional<uint32_t> _firstBatchElapsedMs;
    uint32_t _batchesSeen = 0u;
    std::vector<LocalSearchIndexCore::Candidate> _candidates;
};

class RecordingSearchCallback final : public IFileSystemSearchCallback
{
public:
    enum class Mode : uint8_t
    {
        Success,
        Cancel,
        FailProgress,
        FailCancel,
        AbortProgress,
        AbortMatch,
    };

    explicit RecordingSearchCallback(Mode mode = Mode::Success, uint32_t cancelAfterChecks = 0u) noexcept : _mode(mode), _cancelAfterChecks(cancelAfterChecks)
    {
    }

    RecordingSearchCallback(const RecordingSearchCallback&)            = delete;
    RecordingSearchCallback(RecordingSearchCallback&&)                 = delete;
    RecordingSearchCallback& operator=(const RecordingSearchCallback&) = delete;
    RecordingSearchCallback& operator=(RecordingSearchCallback&&)      = delete;

    HRESULT STDMETHODCALLTYPE FileSystemSearchMatch(const ::FileSystemSearchMatch* searchMatch, [[maybe_unused]] void* cookie) noexcept override
    {
        if (searchMatch == nullptr)
        {
            return E_POINTER;
        }

        if (searchMatch->sizeBytes != sizeof(::FileSystemSearchMatch))
        {
            return E_INVALIDARG;
        }

        if (_mode == Mode::AbortMatch)
        {
            return E_ABORT;
        }

        RecordedSearchMatch copy{};
        copy.fullPath       = CopySearchString(searchMatch->fullPath, searchMatch->fullPathSize);
        copy.relativePath   = CopySearchString(searchMatch->relativePath, searchMatch->relativePathSize);
        copy.displayName    = CopySearchString(searchMatch->displayName, searchMatch->displayNameSize);
        copy.previewText    = CopySearchString(searchMatch->previewText, searchMatch->previewTextSize);
        copy.creationTime   = searchMatch->creationTime;
        copy.lastAccessTime = searchMatch->lastAccessTime;
        copy.lastWriteTime  = searchMatch->lastWriteTime;
        copy.changeTime     = searchMatch->changeTime;
        copy.endOfFile      = searchMatch->endOfFile;
        copy.allocationSize = searchMatch->allocationSize;
        copy.matchedBy      = searchMatch->matchedBy;
        copy.attributes     = searchMatch->fileAttributes;

        std::lock_guard lock(_mutex);
        _matches.push_back(std::move(copy));
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchProgress(const ::FileSystemSearchProgress* searchProgress, [[maybe_unused]] void* cookie) noexcept override
    {
        if (searchProgress == nullptr)
        {
            return E_POINTER;
        }

        if (searchProgress->sizeBytes != sizeof(::FileSystemSearchProgress))
        {
            return E_INVALIDARG;
        }

        _progressCalls.fetch_add(1u, std::memory_order_relaxed);
        if (searchProgress->currentPath == nullptr)
        {
            _nullPathProgressCalls.fetch_add(1u, std::memory_order_relaxed);
        }

        RecordedSearchProgress copy{};
        copy.phase              = searchProgress->phase;
        copy.backend            = searchProgress->backend;
        copy.warningFlags       = searchProgress->warningFlags;
        copy.statusHint         = searchProgress->statusHint;
        copy.scannedDirectories = searchProgress->scannedDirectories;
        copy.scannedFiles       = searchProgress->scannedFiles;
        copy.candidateFiles     = searchProgress->candidateFiles;
        copy.matchedEntries     = searchProgress->matchedEntries;
        copy.currentPath        = CopySearchString(searchProgress->currentPath, searchProgress->currentPathSize);
        {
            std::lock_guard lock(_mutex);
            _progress.push_back(std::move(copy));
        }

        if (_mode == Mode::FailProgress)
        {
            return HRESULT_FROM_WIN32(ERROR_RETRY);
        }

        if (_mode == Mode::AbortProgress)
        {
            return E_ABORT;
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemSearchShouldCancel(BOOL* pCancel, [[maybe_unused]] void* cookie) noexcept override
    {
        if (pCancel == nullptr)
        {
            return E_POINTER;
        }

        _cancelCalls.fetch_add(1u, std::memory_order_relaxed);

        if (_mode == Mode::FailCancel)
        {
            return E_ACCESSDENIED;
        }

        const uint32_t cancelCalls = _cancelCalls.load(std::memory_order_acquire);
        *pCancel                   = (_mode == Mode::Cancel || (_cancelAfterChecks != 0u && cancelCalls >= _cancelAfterChecks)) ? TRUE : FALSE;
        return S_OK;
    }

    [[nodiscard]] std::vector<RecordedSearchMatch> Matches() const noexcept
    {
        std::lock_guard lock(_mutex);
        return _matches;
    }

    [[nodiscard]] uint32_t ProgressCalls() const noexcept
    {
        return _progressCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint32_t CancelCalls() const noexcept
    {
        return _cancelCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint32_t NullPathProgressCalls() const noexcept
    {
        return _nullPathProgressCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] std::vector<RecordedSearchProgress> ProgressSnapshots() const noexcept
    {
        std::lock_guard lock(_mutex);
        return _progress;
    }

private:
    Mode _mode                  = Mode::Success;
    uint32_t _cancelAfterChecks = 0u;
    mutable std::mutex _mutex;
    std::vector<RecordedSearchMatch> _matches;
    std::vector<RecordedSearchProgress> _progress;
    std::atomic<uint32_t> _progressCalls{0};
    std::atomic<uint32_t> _cancelCalls{0};
    std::atomic<uint32_t> _nullPathProgressCalls{0};
};

[[nodiscard]] const RecordedSearchMatch* FindRecordedSearchMatch(const std::vector<RecordedSearchMatch>& matches, std::wstring_view displayName) noexcept
{
    const auto it = std::find_if(
        matches.begin(), matches.end(), [&](const RecordedSearchMatch& match) noexcept { return std::wstring_view(match.displayName) == displayName; });
    return it != matches.end() ? &*it : nullptr;
}

[[nodiscard]] const RecordedSearchProgress* FindRecordedSearchProgress(const std::vector<RecordedSearchProgress>& progress,
                                                                       FileSystemSearchPhase phase) noexcept
{
    const auto it = std::find_if(progress.begin(), progress.end(), [&](const RecordedSearchProgress& snapshot) noexcept { return snapshot.phase == phase; });
    return it != progress.end() ? &*it : nullptr;
}

[[nodiscard]] bool PrepareSearchCaseRoot(const std::filesystem::path& root, std::wstring_view name, std::filesystem::path& outPath) noexcept
{
    outPath = root / std::filesystem::path(name);
    std::error_code removeError;
    std::filesystem::remove_all(outPath, removeError);
    return SelfTest::EnsureDirectory(outPath);
}

[[nodiscard]] std::optional<std::filesystem::path> FindFirstFixedVolumeByFileSystem(std::wstring_view fileSystemType) noexcept
{
    wchar_t drives[512] = {};
    const DWORD written = ::GetLogicalDriveStringsW(static_cast<DWORD>(std::size(drives)), drives);
    if (written == 0u || written >= std::size(drives))
    {
        return std::nullopt;
    }

    for (const wchar_t* cursor = drives; *cursor != L'\0'; cursor += std::char_traits<wchar_t>::length(cursor) + 1u)
    {
        if (::GetDriveTypeW(cursor) != DRIVE_FIXED)
        {
            continue;
        }

        wchar_t volumeFileSystem[MAX_PATH] = {};
        if (::GetVolumeInformationW(cursor, nullptr, 0u, nullptr, nullptr, nullptr, volumeFileSystem, static_cast<DWORD>(std::size(volumeFileSystem))) == 0)
        {
            continue;
        }

        if (EqualsIgnoreCase(volumeFileSystem, fileSystemType))
        {
            return std::filesystem::path(cursor);
        }
    }

    return std::nullopt;
}

[[nodiscard]] std::filesystem::path GetSiblingExecutablePath(std::wstring_view fileName) noexcept
{
    std::filesystem::path path;
    wchar_t buffer[MAX_PATH] = {};
    const DWORD written      = ::GetModuleFileNameW(nullptr, buffer, static_cast<DWORD>(std::size(buffer)));
    if (written == 0u || written >= std::size(buffer))
    {
        return path;
    }

    path = std::filesystem::path(buffer).parent_path() / std::filesystem::path(fileName);
    return path;
}

[[nodiscard]] std::wstring MakeUniquePipeName() noexcept
{
    return std::format(LR"(\\.\pipe\RedSalamander.SearchService.Test.{})", MakeGuidText());
}

struct CapturedProcessResult final
{
    DWORD exitCode = 0u;
    std::string output;
};

[[nodiscard]] bool RunProcessAndCaptureOutput(
    std::wstring_view executablePath, std::wstring_view arguments, DWORD timeoutMs, CapturedProcessResult& outResult, std::wstring& outError) noexcept
{
    outResult = {};
    outError.clear();

    SECURITY_ATTRIBUTES securityAttributes{};
    securityAttributes.nLength              = sizeof(securityAttributes);
    securityAttributes.bInheritHandle       = TRUE;
    securityAttributes.lpSecurityDescriptor = nullptr;

    wil::unique_handle readPipe;
    wil::unique_handle writePipe;
    if (::CreatePipe(readPipe.put(), writePipe.put(), &securityAttributes, 0u) == 0)
    {
        outError = std::format(L"CreatePipe failed. error={}", GetLastError());
        return false;
    }

    if (::SetHandleInformation(readPipe.get(), HANDLE_FLAG_INHERIT, 0u) == 0)
    {
        outError = std::format(L"SetHandleInformation failed. error={}", GetLastError());
        return false;
    }

    std::wstring commandLine = std::format(L"\"{}\"", executablePath);
    if (! arguments.empty())
    {
        commandLine.append(L" ");
        commandLine.append(arguments);
    }

    STARTUPINFOW startupInfo{};
    startupInfo.cb         = sizeof(startupInfo);
    startupInfo.dwFlags    = STARTF_USESTDHANDLES;
    startupInfo.hStdInput  = ::GetStdHandle(STD_INPUT_HANDLE);
    startupInfo.hStdOutput = writePipe.get();
    startupInfo.hStdError  = writePipe.get();

    PROCESS_INFORMATION processInfo{};
    if (::CreateProcessW(nullptr, commandLine.data(), nullptr, nullptr, TRUE, 0u, nullptr, nullptr, &startupInfo, &processInfo) == 0)
    {
        outError = std::format(L"CreateProcessW failed. error={}", GetLastError());
        return false;
    }

    wil::unique_handle process(processInfo.hProcess);
    wil::unique_handle thread(processInfo.hThread);
    writePipe.reset();

    const DWORD waitResult = ::WaitForSingleObject(process.get(), timeoutMs);
    if (waitResult == WAIT_TIMEOUT)
    {
        static_cast<void>(::TerminateProcess(process.get(), 1u));
        static_cast<void>(::WaitForSingleObject(process.get(), 2000u));
        outError = L"Timed out waiting for child process completion.";
        return false;
    }

    if (waitResult != WAIT_OBJECT_0)
    {
        outError = std::format(L"WaitForSingleObject failed. error={}", GetLastError());
        return false;
    }

    DWORD exitCode = 0u;
    if (::GetExitCodeProcess(process.get(), &exitCode) == 0)
    {
        outError = std::format(L"GetExitCodeProcess failed. error={}", GetLastError());
        return false;
    }
    outResult.exitCode = exitCode;

    std::array<char, 4096> buffer{};
    for (;;)
    {
        DWORD bytesRead = 0u;
        if (::ReadFile(readPipe.get(), buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr) == 0)
        {
            const DWORD error = ::GetLastError();
            if (error == ERROR_BROKEN_PIPE)
            {
                break;
            }

            outError = std::format(L"ReadFile failed. error={}", error);
            return false;
        }

        if (bytesRead == 0u)
        {
            break;
        }

        try
        {
            outResult.output.append(buffer.data(), buffer.data() + bytesRead);
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::exception&)
        {
            // Mandatory: self-test helper runs under noexcept tests. Fail the helper instead of unwinding.
            outError = L"Process output capture failed with std::exception.";
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool TryReadPortableExecutableSubsystem(std::wstring_view filePath, WORD& outSubsystem, std::wstring& outError) noexcept
{
    outSubsystem = 0u;
    outError.clear();
    const std::wstring pathText(filePath);

    wil::unique_handle file(::CreateFileW(
        pathText.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        outError = std::format(L"CreateFileW failed for '{}'. error={}", pathText, GetLastError());
        return false;
    }

    LARGE_INTEGER size{};
    if (::GetFileSizeEx(file.get(), &size) == 0)
    {
        outError = std::format(L"GetFileSizeEx failed for '{}'. error={}", pathText, GetLastError());
        return false;
    }

    if (size.QuadPart < static_cast<LONGLONG>(sizeof(IMAGE_DOS_HEADER)))
    {
        outError = std::format(L"'{}' is too small to be a PE image.", pathText);
        return false;
    }

    if (size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<DWORD>::max()))
    {
        outError = std::format(L"'{}' is too large for the PE self-test reader.", pathText);
        return false;
    }

    std::vector<std::byte> bytes;
    try
    {
        bytes.resize(static_cast<size_t>(size.QuadPart));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Mandatory: self-test helper runs under noexcept tests. Fail the helper instead of unwinding.
        outError = L"Allocating the PE image buffer failed with std::exception.";
        return false;
    }

    DWORD bytesRead = 0u;
    if (::ReadFile(file.get(), bytes.data(), static_cast<DWORD>(bytes.size()), &bytesRead, nullptr) == 0)
    {
        outError = std::format(L"ReadFile failed for '{}'. error={}", pathText, GetLastError());
        return false;
    }

    if (bytesRead != bytes.size())
    {
        outError = std::format(L"ReadFile returned {} bytes for '{}', expected {}.", bytesRead, pathText, bytes.size());
        return false;
    }

    const auto* dosHeader = reinterpret_cast<const IMAGE_DOS_HEADER*>(bytes.data());
    if (dosHeader->e_magic != IMAGE_DOS_SIGNATURE)
    {
        outError = std::format(L"'{}' is missing the DOS signature.", pathText);
        return false;
    }

    if (dosHeader->e_lfanew < 0)
    {
        outError = std::format(L"'{}' has a negative PE header offset.", pathText);
        return false;
    }

    const size_t ntHeaderOffset = static_cast<size_t>(dosHeader->e_lfanew);
    if (ntHeaderOffset + sizeof(DWORD) + sizeof(IMAGE_FILE_HEADER) + sizeof(WORD) > bytes.size())
    {
        outError = std::format(L"'{}' has a truncated PE header.", pathText);
        return false;
    }

    const auto* signature = reinterpret_cast<const DWORD*>(bytes.data() + ntHeaderOffset);
    if (*signature != IMAGE_NT_SIGNATURE)
    {
        outError = std::format(L"'{}' is missing the NT signature.", pathText);
        return false;
    }

    const size_t optionalHeaderOffset = ntHeaderOffset + sizeof(DWORD) + sizeof(IMAGE_FILE_HEADER);
    const auto* optionalMagic         = reinterpret_cast<const WORD*>(bytes.data() + optionalHeaderOffset);
    if (*optionalMagic == IMAGE_NT_OPTIONAL_HDR32_MAGIC)
    {
        if (ntHeaderOffset + sizeof(IMAGE_NT_HEADERS32) > bytes.size())
        {
            outError = std::format(L"'{}' has a truncated PE32 header.", pathText);
            return false;
        }

        const auto* ntHeaders = reinterpret_cast<const IMAGE_NT_HEADERS32*>(bytes.data() + ntHeaderOffset);
        outSubsystem          = ntHeaders->OptionalHeader.Subsystem;
        return true;
    }

    if (*optionalMagic == IMAGE_NT_OPTIONAL_HDR64_MAGIC)
    {
        if (ntHeaderOffset + sizeof(IMAGE_NT_HEADERS64) > bytes.size())
        {
            outError = std::format(L"'{}' has a truncated PE32+ header.", pathText);
            return false;
        }

        const auto* ntHeaders = reinterpret_cast<const IMAGE_NT_HEADERS64*>(bytes.data() + ntHeaderOffset);
        outSubsystem          = ntHeaders->OptionalHeader.Subsystem;
        return true;
    }

    outError = std::format(L"'{}' has an unknown optional header magic 0x{:04X}.", pathText, *optionalMagic);
    return false;
}

class ForegroundSearchServiceProcess final
{
public:
    ForegroundSearchServiceProcess()                                                 = default;
    ForegroundSearchServiceProcess(const ForegroundSearchServiceProcess&)            = delete;
    ForegroundSearchServiceProcess& operator=(const ForegroundSearchServiceProcess&) = delete;
    ~ForegroundSearchServiceProcess()
    {
        Stop();
    }

    [[nodiscard]] bool Start(std::wstring_view pipeName,
                             uint32_t maxRequestsBeforeExit,
                             uint32_t disconnectAfterBatches,
                             uint32_t protocolVersion,
                             bool waitForReady,
                             std::wstring& outError,
                             bool captureOutput               = false,
                             std::wstring_view extraArguments = {}) noexcept
    {
        outError.clear();
        _capturedOutput.clear();

        const std::filesystem::path servicePath = GetSiblingExecutablePath(L"RedSalamanderSearchService.exe");
        std::error_code existsEc;
        const bool serviceExists = ! servicePath.empty() && std::filesystem::exists(servicePath, existsEc);
        if (! serviceExists)
        {
            outError = std::format(L"Service executable not found: {}", servicePath.wstring());
            return false;
        }

        _pipeName = std::wstring(pipeName);

        std::wstring commandLine = std::format(L"\"{}\" --run-foreground --pipe-name=\"{}\" --max-requests={} --protocol-version={}",
                                               servicePath.wstring(),
                                               _pipeName,
                                               maxRequestsBeforeExit,
                                               protocolVersion);
        if (disconnectAfterBatches != 0u)
        {
            commandLine.append(std::format(L" --disconnect-after-batches={}", disconnectAfterBatches));
        }
        if (! extraArguments.empty())
        {
            commandLine.push_back(L' ');
            commandLine.append(extraArguments);
        }

        SECURITY_ATTRIBUTES securityAttributes{};
        securityAttributes.nLength              = sizeof(securityAttributes);
        securityAttributes.bInheritHandle       = TRUE;
        securityAttributes.lpSecurityDescriptor = nullptr;

        if (captureOutput)
        {
            if (! CreateCaptureFile(securityAttributes, outError))
            {
                return false;
            }
        }

        STARTUPINFOW startupInfo{};
        startupInfo.cb = sizeof(startupInfo);
        if (captureOutput)
        {
            startupInfo.dwFlags    = STARTF_USESTDHANDLES;
            startupInfo.hStdInput  = ::GetStdHandle(STD_INPUT_HANDLE);
            startupInfo.hStdOutput = _stdoutCaptureFile.get();
            startupInfo.hStdError  = _stdoutCaptureFile.get();
        }

        PROCESS_INFORMATION processInfo{};
        if (::CreateProcessW(nullptr, commandLine.data(), nullptr, nullptr, captureOutput ? TRUE : FALSE, 0u, nullptr, nullptr, &startupInfo, &processInfo) ==
            0)
        {
            outError = std::format(L"CreateProcessW failed. error={}", GetLastError());
            return false;
        }

        _process.reset(processInfo.hProcess);
        _thread.reset(processInfo.hThread);
        if (! waitForReady)
        {
            return true;
        }

        return WaitUntilReady(outError);
    }

    void Stop() noexcept
    {
        if (_process)
        {
            static_cast<void>(::TerminateProcess(_process.get(), 0u));
            static_cast<void>(::WaitForSingleObject(_process.get(), 2000u));
        }
        _thread.reset();
        _process.reset();
        _stdoutCaptureFile.reset();
        DeleteCapturedOutputFile();
        _capturedOutput.clear();
        _pipeName.clear();
    }

    [[nodiscard]] bool WaitForExitAndCapture(DWORD timeoutMs, std::string& outOutput, std::wstring& outError) noexcept
    {
        outOutput.clear();
        outError.clear();

        if (! _process)
        {
            outError = L"Foreground search service process is not running.";
            return false;
        }

        const DWORD waitResult = ::WaitForSingleObject(_process.get(), timeoutMs);
        if (waitResult == WAIT_TIMEOUT)
        {
            outError = L"Timed out waiting for the foreground search service process to exit.";
            return false;
        }
        if (waitResult != WAIT_OBJECT_0)
        {
            outError = std::format(L"WaitForSingleObject failed. error={}", GetLastError());
            return false;
        }

        if (! DrainCapturedOutput(outError))
        {
            return false;
        }

        outOutput = _capturedOutput;
        return true;
    }

private:
    [[nodiscard]] bool WaitForPipeReady(std::wstring& outError, bool allowExitedProcess) noexcept
    {
        if (_pipeName.empty())
        {
            outError = L"Foreground search service pipe name is empty.";
            return false;
        }

        for (int attempt = 0; attempt < 50; ++attempt)
        {
            if (_process && ::WaitForSingleObject(_process.get(), 0u) == WAIT_OBJECT_0)
            {
                DWORD exitCode = 0u;
                static_cast<void>(::GetExitCodeProcess(_process.get(), &exitCode));
                if (allowExitedProcess && exitCode == 0u)
                {
                    return true;
                }

                outError = std::format(L"Service process exited before pipe readiness. exitCode={}", exitCode);
                return false;
            }

            if (::WaitNamedPipeW(_pipeName.c_str(), 50u) != 0)
            {
                return true;
            }

            const DWORD error = ::GetLastError();
            if (error != ERROR_FILE_NOT_FOUND && error != ERROR_SEM_TIMEOUT && error != ERROR_PIPE_BUSY)
            {
                outError = std::format(L"WaitNamedPipeW failed while waiting for the foreground search service pipe. error={}", error);
                return false;
            }

            std::this_thread::sleep_for(std::chrono::milliseconds(25));
        }

        outError = L"Timed out waiting for the foreground search service pipe.";
        return false;
    }

    [[nodiscard]] bool WaitUntilReady(std::wstring& outError) noexcept
    {
        if (! WaitForPipeReady(outError, false))
        {
            return false;
        }

        for (int attempt = 0; attempt < 50; ++attempt)
        {
            SearchServiceBroker::ServiceStatus status{};
            if (SUCCEEDED(SearchServiceBroker::GetStatus(status)))
            {
                return WaitForPipeReady(outError, true);
            }

            if (_process && ::WaitForSingleObject(_process.get(), 0u) == WAIT_OBJECT_0)
            {
                DWORD exitCode = 0u;
                static_cast<void>(::GetExitCodeProcess(_process.get(), &exitCode));
                outError = std::format(L"Service process exited before readiness. exitCode={}", exitCode);
                return false;
            }

            std::this_thread::sleep_for(std::chrono::milliseconds(50));
        }

        outError = L"Timed out waiting for the search service foreground process.";
        return false;
    }

    [[nodiscard]] bool DrainCapturedOutput(std::wstring& outError) noexcept
    {
        if (! _stdoutCaptureFile)
        {
            return true;
        }

        LARGE_INTEGER begin{};
        if (::SetFilePointerEx(_stdoutCaptureFile.get(), begin, nullptr, FILE_BEGIN) == 0)
        {
            outError = std::format(L"SetFilePointerEx failed. error={}", GetLastError());
            return false;
        }

        std::array<char, 4096> buffer{};
        for (;;)
        {
            DWORD bytesRead = 0u;
            if (::ReadFile(_stdoutCaptureFile.get(), buffer.data(), static_cast<DWORD>(buffer.size()), &bytesRead, nullptr) == 0)
            {
                const DWORD error = ::GetLastError();
                outError          = std::format(L"ReadFile failed. error={}", error);
                return false;
            }

            if (bytesRead == 0u)
            {
                break;
            }

            try
            {
                _capturedOutput.append(buffer.data(), buffer.data() + bytesRead);
            }
            catch (const std::bad_alloc&)
            {
                std::terminate();
            }
            catch (const std::exception&)
            {
                // Mandatory: self-test helper is used from noexcept tests. Convert output failures into a test error.
                outError = L"Foreground search service output capture failed with std::exception.";
                return false;
            }
        }

        _stdoutCaptureFile.reset();
        DeleteCapturedOutputFile();
        return true;
    }

    [[nodiscard]] bool CreateCaptureFile(const SECURITY_ATTRIBUTES& securityAttributes, std::wstring& outError) noexcept
    {
        DeleteCapturedOutputFile();

        wchar_t tempPath[MAX_PATH] = {};
        const DWORD pathLength     = ::GetTempPathW(static_cast<DWORD>(std::size(tempPath)), tempPath);
        if (pathLength == 0u || pathLength >= std::size(tempPath))
        {
            outError = std::format(L"GetTempPathW failed. error={}", GetLastError());
            return false;
        }

        wchar_t tempFile[MAX_PATH] = {};
        if (::GetTempFileNameW(tempPath, L"RSS", 0u, tempFile) == 0)
        {
            outError = std::format(L"GetTempFileNameW failed. error={}", GetLastError());
            return false;
        }

        _capturedOutputPath = tempFile;
        _stdoutCaptureFile.reset(::CreateFileW(_capturedOutputPath.c_str(),
                                               GENERIC_READ | GENERIC_WRITE,
                                               FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                               const_cast<SECURITY_ATTRIBUTES*>(&securityAttributes),
                                               CREATE_ALWAYS,
                                               FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_SEQUENTIAL_SCAN,
                                               nullptr));
        if (! _stdoutCaptureFile)
        {
            outError = std::format(L"CreateFileW failed. error={}", GetLastError());
            DeleteCapturedOutputFile();
            return false;
        }

        return true;
    }

    void DeleteCapturedOutputFile() noexcept
    {
        if (_capturedOutputPath.empty())
        {
            return;
        }

        std::error_code ec;
        std::filesystem::remove(_capturedOutputPath, ec);
        _capturedOutputPath.clear();
    }

    wil::unique_handle _process;
    wil::unique_handle _thread;
    wil::unique_handle _stdoutCaptureFile;
    std::string _capturedOutput;
    std::wstring _capturedOutputPath;
    std::wstring _pipeName;
};

[[nodiscard]] std::vector<std::wstring> CollectIndexedCandidateNames(const std::vector<LocalSearchIndexCore::Candidate>& candidates) noexcept
{
    std::vector<std::wstring> names;
    names.reserve(candidates.size());
    for (const auto& candidate : candidates)
    {
        names.push_back(candidate.displayName);
    }
    std::sort(names.begin(), names.end());
    return names;
}

[[nodiscard]] std::vector<std::filesystem::path> CollectDirectoryFilesByExtension(const std::filesystem::path& directory, std::wstring_view extension) noexcept
{
    std::vector<std::filesystem::path> files;
    std::error_code ec;
    for (std::filesystem::directory_iterator it(directory, ec), end; ! ec && it != end; it.increment(ec))
    {
        std::error_code statusEc;
        if (! it->is_regular_file(statusEc) || statusEc)
        {
            continue;
        }

        const std::filesystem::path path = it->path();
        if (EqualsIgnoreCase(path.extension().wstring(), extension))
        {
            files.push_back(path);
        }
    }

    std::sort(files.begin(), files.end());
    return files;
}

[[nodiscard]] bool CanProveDirectSqliteCurrentness(const LocalSearchIndexCore::QueryStats& stats) noexcept
{
    return stats.journalAvailable && ! stats.usedTraversalSeed;
}

[[nodiscard]] std::wstring_view DirectSqliteCurrentnessSkipReason(const LocalSearchIndexCore::QueryStats& stats) noexcept
{
    if (! stats.journalAvailable)
    {
        return L"Live journal cursor unavailable; direct SQLite query cutover cannot prove currentness.";
    }
    if (stats.usedTraversalSeed)
    {
        return L"NTFS MFT enumeration unavailable; traversal-seeded SQLite roots remain currentness-unproven until a later journal-backed run.";
    }

    return {};
}

[[nodiscard]] HRESULT RunIndexedNameQuery(LocalSearchIndexCore::Repository& repository,
                                          std::wstring_view rootPath,
                                          std::wstring_view namePattern,
                                          FileSystemSearchNameMode nameMode,
                                          LocalSearchIndexCore::QueryStats& outStats,
                                          std::vector<LocalSearchIndexCore::Candidate>& outCandidates) noexcept
{
    LocalSearchIndexCore::QueryPlan plan{};
    plan.rootPath           = std::wstring(rootPath);
    plan.namePattern        = std::wstring(namePattern);
    plan.nameMode           = nameMode;
    plan.matchCaseName      = false;
    plan.recursive          = true;
    plan.includeFiles       = true;
    plan.includeDirectories = false;
    plan.maxResults         = 0u;

    return repository.Query(plan, nullptr, nullptr, outCandidates, &outStats);
}

[[nodiscard]] HRESULT RunIndexedNameQuery(LocalSearchIndexCore::Repository& repository,
                                          std::wstring_view rootPath,
                                          std::wstring_view namePattern,
                                          LocalSearchIndexCore::QueryStats& outStats,
                                          std::vector<LocalSearchIndexCore::Candidate>& outCandidates) noexcept
{
    return RunIndexedNameQuery(repository, rootPath, namePattern, FILESYSTEM_SEARCH_NAME_WILDCARD, outStats, outCandidates);
}

[[nodiscard]] bool TryPrepareDirectSqliteCursorForSelfTest(
    SelfTest::CaseState& state, const std::filesystem::path& caseRoot, std::wstring_view label, uint64_t& outJournalId, uint64_t& outNextUsn) noexcept
{
    outJournalId = 0u;
    outNextUsn   = 0u;

    try
    {
        const std::filesystem::path cursorStoreRoot = caseRoot.parent_path() / (caseRoot.filename().wstring() + L"-cursor-probe-store");
        LocalSearchIndexCore::Repository cursorRepository({
            .snapshotRootDirectory = cursorStoreRoot.wstring(),
        });

        LocalSearchIndexCore::QueryStats cursorStats{};
        std::vector<LocalSearchIndexCore::Candidate> cursorCandidates;
        const HRESULT cursorHr = RunIndexedNameQuery(cursorRepository, caseRoot.wstring(), L"*", cursorStats, cursorCandidates);
        state.Require(SUCCEEDED(cursorHr), std::format(L"{} cursor probe query failed. hr=0x{:08X}", label, static_cast<unsigned long>(cursorHr)));
        if (FAILED(cursorHr))
        {
            return false;
        }

        if (! CanProveDirectSqliteCurrentness(cursorStats))
        {
            state.Skip(DirectSqliteCurrentnessSkipReason(cursorStats));
            return false;
        }

        outJournalId = cursorStats.journalId;
        outNextUsn   = cursorStats.nextUsn;
        return true;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"CompareSelfTest: direct SQLite cursor probe failed with an unexpected std::exception.");
        state.Require(false, std::format(L"{} cursor probe failed with an unexpected std::exception.", label));
        return false;
    }
}

[[nodiscard]] bool ExecuteSqliteScript(const std::filesystem::path& databasePath, std::string_view script, std::wstring& outError) noexcept
{
    outError.clear();

    sqlite3* rawDb               = nullptr;
    const std::u8string utf8Path = databasePath.u8string();
    if (sqlite3_open_v2(reinterpret_cast<const char*>(utf8Path.c_str()), &rawDb, SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE | SQLITE_OPEN_NOMUTEX, nullptr) !=
        SQLITE_OK)
    {
        const auto* message = (rawDb != nullptr) ? static_cast<const wchar_t*>(sqlite3_errmsg16(rawDb)) : nullptr;
        outError            = std::format(L"sqlite3_open_v2 failed for '{}': {}", databasePath.wstring(), (message != nullptr) ? message : L"<unknown>");
        if (rawDb != nullptr)
        {
            static_cast<void>(sqlite3_close(rawDb));
        }
        return false;
    }

    const auto closeDb = wil::scope_exit([&]() noexcept { static_cast<void>(sqlite3_close(rawDb)); });

    char* errorText        = nullptr;
    const int sqliteResult = sqlite3_exec(rawDb, script.data(), nullptr, nullptr, &errorText);
    const auto freeError   = wil::scope_exit([&]() noexcept
    {
        if (errorText != nullptr)
        {
            sqlite3_free(errorText);
        }
    });

    if (sqliteResult != SQLITE_OK)
    {
        std::wstring message;
        if (errorText != nullptr)
        {
            const int required = ::MultiByteToWideChar(CP_UTF8, 0, errorText, -1, nullptr, 0);
            if (required > 0)
            {
                std::wstring buffer(static_cast<size_t>(required), L'\0');
                static_cast<void>(::MultiByteToWideChar(CP_UTF8, 0, errorText, -1, buffer.data(), required));
                if (! buffer.empty() && buffer.back() == L'\0')
                {
                    buffer.pop_back();
                }
                message = std::move(buffer);
            }
        }

        if (message.empty())
        {
            const auto* fallback = static_cast<const wchar_t*>(sqlite3_errmsg16(rawDb));
            message              = (fallback != nullptr) ? std::wstring(fallback) : std::wstring(L"<unknown>");
        }

        outError = std::format(L"sqlite3_exec failed for '{}': code={} message='{}'", databasePath.wstring(), sqliteResult, message);
        return false;
    }

    return true;
}

[[nodiscard]] bool QuerySqliteSingleInt64(const std::filesystem::path& databasePath, std::string_view sql, uint64_t& outValue, std::wstring& outError) noexcept
{
    outError.clear();
    outValue = 0u;

    sqlite3* rawDb               = nullptr;
    const std::u8string utf8Path = databasePath.u8string();
    if (sqlite3_open_v2(reinterpret_cast<const char*>(utf8Path.c_str()), &rawDb, SQLITE_OPEN_READONLY | SQLITE_OPEN_NOMUTEX, nullptr) != SQLITE_OK)
    {
        const auto* message = (rawDb != nullptr) ? static_cast<const wchar_t*>(sqlite3_errmsg16(rawDb)) : nullptr;
        outError            = std::format(L"sqlite3_open_v2 failed for '{}': {}", databasePath.wstring(), (message != nullptr) ? message : L"<unknown>");
        if (rawDb != nullptr)
        {
            static_cast<void>(sqlite3_close(rawDb));
        }
        return false;
    }

    const auto closeDb = wil::scope_exit([&]() noexcept { static_cast<void>(sqlite3_close(rawDb)); });

    sqlite3_stmt* rawStatement = nullptr;
    const int prepareResult    = sqlite3_prepare_v2(rawDb, sql.data(), static_cast<int>(sql.size()), &rawStatement, nullptr);
    if (prepareResult != SQLITE_OK)
    {
        const auto* message = static_cast<const wchar_t*>(sqlite3_errmsg16(rawDb));
        outError            = std::format(
            L"sqlite3_prepare_v2 failed for '{}': code={} message='{}'", databasePath.wstring(), prepareResult, (message != nullptr) ? message : L"<unknown>");
        return false;
    }

    const auto finalizeStatement = wil::scope_exit([&]() noexcept
    {
        if (rawStatement != nullptr)
        {
            static_cast<void>(sqlite3_finalize(rawStatement));
        }
    });

    const int stepResult = sqlite3_step(rawStatement);
    if (stepResult != SQLITE_ROW)
    {
        const auto* message = static_cast<const wchar_t*>(sqlite3_errmsg16(rawDb));
        outError            = std::format(
            L"sqlite3_step failed for '{}': code={} message='{}'", databasePath.wstring(), stepResult, (message != nullptr) ? message : L"<unknown>");
        return false;
    }

    outValue = static_cast<uint64_t>(sqlite3_column_int64(rawStatement, 0));
    return true;
}

[[nodiscard]] bool CreateInformations(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IInformations>& outInfo) noexcept
{
    outInfo.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IInformations), outInfo.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outInfo);
}

[[nodiscard]] bool CreateFileSystemDirectoryOperations(const wil::com_ptr<IFileSystem>& fs, wil::com_ptr<IFileSystemDirectoryOperations>& outOps) noexcept
{
    outOps.reset();
    if (! fs)
    {
        return false;
    }

    const HRESULT hr = fs->QueryInterface(__uuidof(IFileSystemDirectoryOperations), outOps.put_void());
    return SUCCEEDED(hr) && static_cast<bool>(outOps);
}

[[nodiscard]] bool EnsureDirectoryExistsFsOps(const wil::com_ptr<IFileSystemDirectoryOperations>& ops, const std::filesystem::path& path) noexcept
{
    if (! ops)
    {
        return false;
    }

    const std::filesystem::path normalized = path.lexically_normal();
    std::filesystem::path current          = normalized.root_path();
    for (const auto& part : normalized.relative_path())
    {
        current /= part;
        const HRESULT hr = ops->CreateDirectory(current.c_str());
        if (SUCCEEDED(hr) || hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            continue;
        }
        return false;
    }

    return true;
}

[[nodiscard]] bool WriteFileBytesFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, const void* data, size_t sizeBytes) noexcept
{
    if (! io || ! data)
    {
        return false;
    }

    if (sizeBytes > static_cast<size_t>(std::numeric_limits<unsigned long>::max()))
    {
        return false;
    }

    wil::com_ptr<IFileWriter> writer;
    const HRESULT createHr = io->CreateFileWriter(path.c_str(), FILESYSTEM_FLAG_NONE, writer.put());
    if (FAILED(createHr) || ! writer)
    {
        return false;
    }

    unsigned long written = 0;
    const HRESULT writeHr = writer->Write(data, static_cast<unsigned long>(sizeBytes), &written);
    if (FAILED(writeHr) || written != static_cast<unsigned long>(sizeBytes))
    {
        return false;
    }

    return SUCCEEDED(writer->Commit());
}

[[nodiscard]] bool WriteFileTextFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, std::string_view text) noexcept
{
    return WriteFileBytesFsIo(io, path, text.data(), text.size());
}

[[nodiscard]] std::wstring ToPluginPathText(const std::filesystem::path& path) noexcept
{
    return path.generic_wstring();
}

[[nodiscard]] std::wstring JoinPluginPathForSelfTest(std::wstring_view base, std::wstring_view leaf) noexcept
{
    std::wstring result = NormalizePluginPathForSelfTest(base);
    if (result.empty())
    {
        return {};
    }

    if (! leaf.empty())
    {
        if (result.back() != L'/')
        {
            result.push_back(L'/');
        }
        result.append(leaf);
    }

    return NormalizePluginPathForSelfTest(result);
}

[[nodiscard]] std::wstring MakeConnectionPathForSelfTest(std::wstring_view profileName, std::wstring_view pluginPath) noexcept
{
    if (profileName.empty() || pluginPath.empty() || pluginPath.front() != L'/')
    {
        return {};
    }

    return std::format(L"/@conn:{}{}", profileName, pluginPath);
}

[[nodiscard]] bool PathExistsFsIo(const wil::com_ptr<IFileSystemIO>& io, const std::filesystem::path& path, unsigned long* attributes = nullptr) noexcept
{
    if (! io)
    {
        return false;
    }

    unsigned long attrs           = 0;
    const std::wstring pluginPath = ToPluginPathText(path);
    const HRESULT hr              = io->GetAttributes(pluginPath.c_str(), &attrs);
    if (FAILED(hr))
    {
        return false;
    }

    if (attributes)
    {
        *attributes = attrs;
    }
    return true;
}

[[nodiscard]] std::filesystem::path GetWorkspaceRootFromSourcePath() noexcept
{
    std::filesystem::path candidate = std::filesystem::path(__FILE__).parent_path();
    for (int i = 0; i < 10 && ! candidate.empty(); ++i)
    {
        std::error_code ec;
        if (std::filesystem::exists(candidate / L"RedSalamander.sln", ec) && ! ec &&
            std::filesystem::exists(candidate / L"Plugins" / L"FileSystem7z" / L"Tests" / L"Tests.zip", ec) && ! ec)
        {
            return candidate;
        }

        if (! candidate.has_parent_path() || candidate.parent_path() == candidate)
        {
            break;
        }
        candidate = candidate.parent_path();
    }

    return std::filesystem::path(__FILE__).parent_path().parent_path().parent_path().parent_path();
}

[[nodiscard]] std::optional<std::wstring> FindFirstRegularEntryPath(const wil::com_ptr<IFileSystem>& fs, std::wstring_view rootPath) noexcept
{
    if (! fs || rootPath.empty())
    {
        return std::nullopt;
    }

    std::vector<std::wstring> pending{std::wstring(rootPath)};
    while (! pending.empty())
    {
        const std::wstring current = std::move(pending.back());
        pending.pop_back();

        wil::com_ptr<IFilesInformation> info;
        const HRESULT hr = fs->ReadDirectoryInfo(current.c_str(), info.put());
        if (FAILED(hr) || ! info)
        {
            continue;
        }

        FileInfo* head = nullptr;
        if (FAILED(info->GetBuffer(&head)) || head == nullptr)
        {
            continue;
        }

        for (FileInfo* entry = head; entry != nullptr;)
        {
            if (entry->FileNameSize >= sizeof(wchar_t))
            {
                const size_t charCount = entry->FileNameSize / sizeof(wchar_t);
                std::wstring path(current);
                if (! path.empty() && path.back() != L'/')
                {
                    path.push_back(L'/');
                }
                path.append(entry->FileName, entry->FileName + charCount);

                const bool isDirectory = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (! isDirectory)
                {
                    return path;
                }

                pending.push_back(std::move(path));
            }

            if (entry->NextEntryOffset == 0)
            {
                break;
            }

            entry = reinterpret_cast<FileInfo*>(reinterpret_cast<std::byte*>(entry) + entry->NextEntryOffset);
        }
    }

    return std::nullopt;
}

enum class DirectorySizeCallbackMode : uint8_t
{
    Success,
    FailProgress,
    FailCancel,
    Cancel,
};

constexpr HRESULT kDirectorySizeProgressFailureHr = HRESULT_FROM_WIN32(ERROR_RETRY);
constexpr HRESULT kDirectorySizeCancelFailureHr   = E_ACCESSDENIED;

class RecordingDirectorySizeCallback final : public IFileSystemDirectorySizeCallback
{
public:
    explicit RecordingDirectorySizeCallback(DirectorySizeCallbackMode mode) noexcept : _mode(mode)
    {
    }

    RecordingDirectorySizeCallback(const RecordingDirectorySizeCallback&)            = delete;
    RecordingDirectorySizeCallback(RecordingDirectorySizeCallback&&)                 = delete;
    RecordingDirectorySizeCallback& operator=(const RecordingDirectorySizeCallback&) = delete;
    RecordingDirectorySizeCallback& operator=(RecordingDirectorySizeCallback&&)      = delete;

    HRESULT STDMETHODCALLTYPE DirectorySizeProgress(uint64_t scannedEntries,
                                                    uint64_t totalBytes,
                                                    uint64_t fileCount,
                                                    uint64_t directoryCount,
                                                    const wchar_t* currentPath,
                                                    [[maybe_unused]] void* cookie) noexcept override
    {
        _progressCalls.fetch_add(1u, std::memory_order_relaxed);
        _lastScannedEntries.store(scannedEntries, std::memory_order_release);
        _lastTotalBytes.store(totalBytes, std::memory_order_release);
        _lastFileCount.store(fileCount, std::memory_order_release);
        _lastDirectoryCount.store(directoryCount, std::memory_order_release);

        if (currentPath == nullptr)
        {
            _nullPathProgressCalls.fetch_add(1u, std::memory_order_relaxed);
        }

        if (_mode == DirectorySizeCallbackMode::FailProgress)
        {
            return kDirectorySizeProgressFailureHr;
        }

        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DirectorySizeShouldCancel(BOOL* pCancel, [[maybe_unused]] void* cookie) noexcept override
    {
        if (pCancel == nullptr)
        {
            return E_POINTER;
        }

        _cancelCalls.fetch_add(1u, std::memory_order_relaxed);

        if (_mode == DirectorySizeCallbackMode::FailCancel)
        {
            return kDirectorySizeCancelFailureHr;
        }

        *pCancel = (_mode == DirectorySizeCallbackMode::Cancel) ? TRUE : FALSE;
        return S_OK;
    }

    [[nodiscard]] uint32_t ProgressCalls() const noexcept
    {
        return _progressCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint32_t CancelCalls() const noexcept
    {
        return _cancelCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint32_t NullPathProgressCalls() const noexcept
    {
        return _nullPathProgressCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint64_t LastTotalBytes() const noexcept
    {
        return _lastTotalBytes.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint64_t LastFileCount() const noexcept
    {
        return _lastFileCount.load(std::memory_order_acquire);
    }

    [[nodiscard]] uint64_t LastDirectoryCount() const noexcept
    {
        return _lastDirectoryCount.load(std::memory_order_acquire);
    }

private:
    DirectorySizeCallbackMode _mode = DirectorySizeCallbackMode::Success;
    std::atomic<uint32_t> _progressCalls{0};
    std::atomic<uint32_t> _cancelCalls{0};
    std::atomic<uint32_t> _nullPathProgressCalls{0};
    std::atomic<uint64_t> _lastScannedEntries{0};
    std::atomic<uint64_t> _lastTotalBytes{0};
    std::atomic<uint64_t> _lastFileCount{0};
    std::atomic<uint64_t> _lastDirectoryCount{0};
};

[[nodiscard]] bool ValidateDirectorySizeCallbackContract(SelfTest::CaseState& state,
                                                         IFileSystemDirectoryOperations* dirOps,
                                                         const std::wstring& path,
                                                         FileSystemFlags flags = FILESYSTEM_FLAG_NONE) noexcept
{
    if (dirOps == nullptr || path.empty())
    {
        state.Require(false, L"Directory size callback contract: invalid test inputs.");
        return false;
    }

    FileSystemDirectorySizeResult baseline{};
    baseline.sizeBytes       = sizeof(FileSystemDirectorySizeResult);
    const HRESULT baselineHr = dirOps->GetDirectorySize(path.c_str(), flags, nullptr, nullptr, &baseline);
    state.Require(SUCCEEDED(baselineHr) && SUCCEEDED(baseline.status),
                  std::format(L"Directory size callback contract: baseline GetDirectorySize failed. hr=0x{:08X} status=0x{:08X}",
                              static_cast<unsigned long>(baselineHr),
                              static_cast<unsigned long>(baseline.status)));
    if (FAILED(baselineHr) || FAILED(baseline.status))
    {
        return false;
    }

    const auto runCase = [&](DirectorySizeCallbackMode mode, HRESULT expectedHr, std::wstring_view label) noexcept
    {
        RecordingDirectorySizeCallback callback(mode);
        FileSystemDirectorySizeResult result{};
        result.sizeBytes = sizeof(FileSystemDirectorySizeResult);

        const HRESULT hr = dirOps->GetDirectorySize(path.c_str(), flags, &callback, nullptr, &result);
        state.Require(hr == expectedHr && result.status == expectedHr,
                      std::format(L"Directory size callback contract ({}): expected hr/status 0x{:08X}, got hr=0x{:08X} status=0x{:08X}.",
                                  label,
                                  static_cast<unsigned long>(expectedHr),
                                  static_cast<unsigned long>(hr),
                                  static_cast<unsigned long>(result.status)));
        state.Require(callback.ProgressCalls() >= 1u, std::format(L"Directory size callback contract ({}): expected at least one progress callback.", label));
        if (mode != DirectorySizeCallbackMode::FailProgress)
        {
            state.Require(callback.CancelCalls() >= 1u, std::format(L"Directory size callback contract ({}): expected at least one cancel callback.", label));
        }
    };

    runCase(DirectorySizeCallbackMode::FailProgress, kDirectorySizeProgressFailureHr, L"progress failure");
    if (! state.failure.empty())
    {
        return false;
    }

    runCase(DirectorySizeCallbackMode::FailCancel, kDirectorySizeCancelFailureHr, L"cancel failure");
    if (! state.failure.empty())
    {
        return false;
    }

    runCase(DirectorySizeCallbackMode::Cancel, HRESULT_FROM_WIN32(ERROR_CANCELLED), L"cancel");
    if (! state.failure.empty())
    {
        return false;
    }

    RecordingDirectorySizeCallback successCallback(DirectorySizeCallbackMode::Success);
    FileSystemDirectorySizeResult success{};
    success.sizeBytes       = sizeof(FileSystemDirectorySizeResult);
    const HRESULT successHr = dirOps->GetDirectorySize(path.c_str(), flags, &successCallback, nullptr, &success);
    state.Require(SUCCEEDED(successHr) && SUCCEEDED(success.status),
                  std::format(L"Directory size callback contract (success): hr=0x{:08X} status=0x{:08X}.",
                              static_cast<unsigned long>(successHr),
                              static_cast<unsigned long>(success.status)));
    state.Require(
        success.totalBytes == baseline.totalBytes,
        std::format(L"Directory size callback contract (success): totalBytes mismatch. expected={} got={}.", baseline.totalBytes, success.totalBytes));
    state.Require(success.fileCount == baseline.fileCount,
                  std::format(L"Directory size callback contract (success): fileCount mismatch. expected={} got={}.", baseline.fileCount, success.fileCount));
    state.Require(success.directoryCount == baseline.directoryCount,
                  std::format(L"Directory size callback contract (success): directoryCount mismatch. expected={} got={}.",
                              baseline.directoryCount,
                              success.directoryCount));
    state.Require(successCallback.NullPathProgressCalls() >= 1u,
                  L"Directory size callback contract (success): expected a final progress callback with currentPath=nullptr.");

    return state.failure.empty();
}

struct RecordedFileSystemItem final
{
    bool seen      = false;
    HRESULT status = S_OK;
    std::wstring sourcePath;
    std::wstring destinationPath;
};

class RecordingFileSystemCallback final : public IFileSystemCallback
{
public:
    explicit RecordingFileSystemCallback(unsigned long expectedItemCount) : _items(expectedItemCount)
    {
    }

    RecordingFileSystemCallback(const RecordingFileSystemCallback&)            = delete;
    RecordingFileSystemCallback(RecordingFileSystemCallback&&)                 = delete;
    RecordingFileSystemCallback& operator=(const RecordingFileSystemCallback&) = delete;
    RecordingFileSystemCallback& operator=(RecordingFileSystemCallback&&)      = delete;

    HRESULT STDMETHODCALLTYPE FileSystemProgress([[maybe_unused]] FileSystemOperation operationType,
                                                 [[maybe_unused]] unsigned long totalItems,
                                                 [[maybe_unused]] unsigned long completedItems,
                                                 [[maybe_unused]] uint64_t totalBytes,
                                                 [[maybe_unused]] uint64_t completedBytes,
                                                 [[maybe_unused]] const wchar_t* currentSourcePath,
                                                 [[maybe_unused]] const wchar_t* currentDestinationPath,
                                                 [[maybe_unused]] uint64_t currentItemTotalBytes,
                                                 [[maybe_unused]] uint64_t currentItemCompletedBytes,
                                                 [[maybe_unused]] FileSystemOptions* options,
                                                 [[maybe_unused]] uint64_t progressStreamId,
                                                 [[maybe_unused]] void* cookie) noexcept override
    {
        _progressCalls.fetch_add(1u, std::memory_order_relaxed);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemItemCompleted([[maybe_unused]] FileSystemOperation operationType,
                                                      unsigned long itemIndex,
                                                      const wchar_t* sourcePath,
                                                      const wchar_t* destinationPath,
                                                      HRESULT status,
                                                      [[maybe_unused]] FileSystemOptions* options,
                                                      [[maybe_unused]] void* cookie) noexcept override
    {
        {
            std::lock_guard lock(_mutex);
            if (itemIndex < _items.size())
            {
                RecordedFileSystemItem& item = _items[itemIndex];
                item.seen                    = true;
                item.status                  = status;
                item.sourcePath              = sourcePath ? sourcePath : L"";
                item.destinationPath         = destinationPath ? destinationPath : L"";
            }
        }

        _completedCalls.fetch_add(1u, std::memory_order_relaxed);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL* pCancel, [[maybe_unused]] void* cookie) noexcept override
    {
        if (pCancel == nullptr)
        {
            return E_POINTER;
        }

        *pCancel = FALSE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE FileSystemIssue([[maybe_unused]] FileSystemOperation operationType,
                                              [[maybe_unused]] const wchar_t* sourcePath,
                                              [[maybe_unused]] const wchar_t* destinationPath,
                                              [[maybe_unused]] HRESULT status,
                                              FileSystemIssueAction* action,
                                              [[maybe_unused]] FileSystemOptions* options,
                                              [[maybe_unused]] void* cookie) noexcept override
    {
        _unexpectedIssue.store(true, std::memory_order_release);
        if (action != nullptr)
        {
            *action = FileSystemIssueAction::Cancel;
        }
        return E_UNEXPECTED;
    }

    [[nodiscard]] bool TryGetItem(unsigned long itemIndex, RecordedFileSystemItem& outItem) const noexcept
    {
        std::lock_guard lock(_mutex);
        if (itemIndex >= _items.size() || ! _items[itemIndex].seen)
        {
            return false;
        }

        outItem = _items[itemIndex];
        return true;
    }

    [[nodiscard]] unsigned long CompletedCount() const noexcept
    {
        return _completedCalls.load(std::memory_order_acquire);
    }

    [[nodiscard]] bool SawUnexpectedIssue() const noexcept
    {
        return _unexpectedIssue.load(std::memory_order_acquire);
    }

private:
    mutable std::mutex _mutex;
    std::vector<RecordedFileSystemItem> _items;
    std::atomic<unsigned long> _progressCalls{0};
    std::atomic<unsigned long> _completedCalls{0};
    std::atomic<bool> _unexpectedIssue{false};
};

#ifndef SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE
#define SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE 0x2u
#endif // ENABLE_TESTS

[[nodiscard]] bool TryCreateDirectorySymlink(const std::filesystem::path& linkPath, const std::filesystem::path& targetPath) noexcept
{
    const DWORD flags = SYMBOLIC_LINK_FLAG_DIRECTORY;

    if (::CreateSymbolicLinkW(linkPath.c_str(), targetPath.c_str(), flags | SYMBOLIC_LINK_FLAG_ALLOW_UNPRIVILEGED_CREATE) != 0)
    {
        return true;
    }

    if (::CreateSymbolicLinkW(linkPath.c_str(), targetPath.c_str(), flags) != 0)
    {
        return true;
    }

    return false;
}

[[nodiscard]] bool WriteFileFill(const std::filesystem::path& path, char ch, size_t sizeBytes) noexcept
{
    if (sizeBytes == 0)
    {
        return SelfTest::WriteBinaryFile(path, {});
    }

    const std::string text(sizeBytes, ch);
    const std::span<const char> textBytes(text.data(), text.size());
    return SelfTest::WriteBinaryFile(path, std::as_bytes(textBytes));
}

struct CaseFolders
{
    std::filesystem::path left;
    std::filesystem::path right;
};

[[nodiscard]] std::optional<CaseFolders> CreateCaseFolders(const std::filesystem::path& base, std::wstring_view caseName) noexcept
{
    std::filesystem::path caseRoot = base / std::filesystem::path(caseName);
    std::filesystem::path left     = caseRoot / L"left";
    std::filesystem::path right    = caseRoot / L"right";

    SelfTest::EnsureDirectory(left);
    SelfTest::EnsureDirectory(right);
    if (! SelfTest::PathExists(left) || ! SelfTest::PathExists(right))
    {
        return std::nullopt;
    }

    return CaseFolders{std::move(left), std::move(right)};
}

void AppendCompareSelfTestTraceLine(std::wstring_view message) noexcept
{
    Trace(message);
}

[[nodiscard]] std::vector<std::wstring> EnumerateDirectoryNames(const wil::com_ptr<IFileSystem>& fs,
                                                                const std::filesystem::path& folder,
                                                                SelfTest::CaseState& state) noexcept
{
    if (! fs)
    {
        state.Require(false, L"EnumerateDirectoryNames: file system is null.");
        return {};
    }

    wil::com_ptr<IFilesInformation> info;
    const HRESULT hr = fs->ReadDirectoryInfo(folder.c_str(), info.put());
    state.Require(SUCCEEDED(hr), L"EnumerateDirectoryNames: ReadDirectoryInfo failed.");
    if (FAILED(hr) || ! info)
    {
        return {};
    }

    FileInfo* head         = nullptr;
    const HRESULT hrBuffer = info->GetBuffer(&head);
    state.Require(SUCCEEDED(hrBuffer), L"EnumerateDirectoryNames: GetBuffer failed.");
    if (FAILED(hrBuffer) || head == nullptr)
    {
        return {};
    }

    std::vector<std::wstring> result;
    for (FileInfo* entry = head; entry != nullptr;)
    {
        const size_t nameChars = static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t);
        result.emplace_back(entry->FileName, nameChars);

        if (entry->NextEntryOffset == 0)
        {
            break;
        }
        entry = reinterpret_cast<FileInfo*>(reinterpret_cast<unsigned char*>(entry) + entry->NextEntryOffset);
    }

    return result;
}

[[nodiscard]] bool ContainsName(const std::vector<std::wstring>& names, std::wstring_view name) noexcept
{
    return std::any_of(names.begin(), names.end(), [&](const std::wstring& value) noexcept { return value == name; });
}

struct GetDecisionSehContext
{
    CompareDirectoriesSession* session                                   = nullptr;
    std::shared_ptr<const CompareDirectoriesFolderDecision>* outDecision = nullptr;
};

void InvokeGetRootDecision(void* rawContext) noexcept
{
    auto* ctx = static_cast<GetDecisionSehContext*>(rawContext);
    if (! ctx || ! ctx->session || ! ctx->outDecision)
    {
        return;
    }

    *ctx->outDecision = ctx->session->GetOrComputeDecision(std::filesystem::path{});
}

[[nodiscard]] bool TryGetRootDecisionWithSeh(CompareDirectoriesSession& session, std::shared_ptr<const CompareDirectoriesFolderDecision>& outDecision) noexcept
{
    GetDecisionSehContext ctx{};
    ctx.session     = &session;
    ctx.outDecision = &outDecision;

    __try
    {
        InvokeGetRootDecision(&ctx);
        return true;
    }
    __except (CrashHandler::WriteDumpForException(GetExceptionInformation()))
    {
        return false;
    }
}

[[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> ComputeRootDecision(wil::com_ptr<IFileSystem> baseFs,
                                                                                          const CaseFolders& folders,
                                                                                          Common::Settings::CompareDirectoriesSettings settings,
                                                                                          SelfTest::CaseState& state) noexcept
{
    if (! baseFs)
    {
        state.Require(false, L"Base file system is null.");
        return {};
    }

    auto session = std::make_shared<CompareDirectoriesSession>(baseFs, baseFs, folders.left, folders.right, std::move(settings));
    std::shared_ptr<const CompareDirectoriesFolderDecision> decision;
    if (! TryGetRootDecisionWithSeh(*session, decision))
    {
        state.Require(false, L"GetOrComputeDecision crashed.");
        return {};
    }
    state.Require(static_cast<bool>(decision), L"GetOrComputeDecision returned null.");
    if (! decision)
    {
        return {};
    }

    state.Require(SUCCEEDED(decision->hr), L"Decision hr is failure.");
    return decision;
}

[[nodiscard]] bool StartScanAndWaitForIdle(const std::shared_ptr<CompareDirectoriesSession>& session, std::chrono::milliseconds timeout) noexcept
{
    if (! session)
    {
        return false;
    }

    std::mutex mutex;
    std::condition_variable cv;
    bool started = false;
    bool done    = false;

    session->SetScanProgressCallback([&](const std::filesystem::path&, std::wstring_view, uint64_t, uint64_t, uint32_t activeScans, uint64_t, uint64_t) noexcept
    {
        std::lock_guard lock(mutex);
        if (activeScans != 0u)
        {
            started = true;
        }
        if (started && activeScans == 0u)
        {
            done = true;
            cv.notify_all();
        }
    });

    session->StartScan();

    {
        std::unique_lock lock(mutex);
        static_cast<void>(cv.wait_for(lock, timeout, [&] { return done; }));
    }

    session->SetScanProgressCallback({});
    return done;
}

[[nodiscard]] bool DrainPendingSubdirUpdates(const std::shared_ptr<CompareDirectoriesSession>& session, size_t maxIterations) noexcept
{
    if (! session)
    {
        return false;
    }

    for (size_t i = 0; i < maxIterations; ++i)
    {
        if (! session->FlushPendingSubdirUpdatesBudgeted(64))
        {
            return true;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    return false;
}

[[nodiscard]] const CompareDirectoriesItemDecision* FindItem(const CompareDirectoriesFolderDecision& decision, std::wstring_view name) noexcept
{
    const auto it = decision.items.find(name);
    if (it == decision.items.end())
    {
        return nullptr;
    }
    return &it->second;
}

[[nodiscard]] std::shared_ptr<const CompareDirectoriesFolderDecision> WaitForContentCompare(const std::shared_ptr<CompareDirectoriesSession>& session,
                                                                                            const std::filesystem::path& relativeFolder,
                                                                                            std::wstring_view itemName,
                                                                                            SelfTest::CaseState& state) noexcept
{
    if (! session)
    {
        state.Require(false, L"WaitForContentCompare: session is null.");
        return {};
    }

    const auto deadline =
        std::chrono::steady_clock::now() + std::chrono::milliseconds{static_cast<std::chrono::milliseconds::rep>(SelfTest::ScaleTimeout(5000))};
    while (std::chrono::steady_clock::now() < deadline)
    {
        auto decision = session->GetOrComputeDecision(relativeFolder);
        state.Require(static_cast<bool>(decision), L"WaitForContentCompare: decision is null.");
        if (! decision)
        {
            return {};
        }

        const auto* item = FindItem(*decision, itemName);
        // In differences-only mode, pending content placeholders may be elided to keep memory bounded.
        // Allow the item to appear later once content compare determines it is actually different.
        if (! item)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
            continue;
        }

        if (! HasFlag(item->differenceMask, CompareDirectoriesDiffBit::ContentPending))
        {
            return decision;
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    }

    state.Require(false, std::format(L"Timed out waiting for content compare: {}.", itemName));
    return session->GetOrComputeDecision(relativeFolder);
}
} // namespace

std::vector<std::wstring> CompareDirectoriesSelfTest::ListCases(const SelfTest::SelfTestOptions& options) noexcept
{
    std::vector<std::wstring> names;
    names.reserve(std::size(kCompareCaseNames));
    for (const std::wstring_view name : kCompareCaseNames)
    {
        if (SelfTest::CaseFilterMatches(options.caseFilter, name))
        {
            names.emplace_back(name);
        }
    }
    return names;
}

bool CompareDirectoriesSelfTest::Run(const SelfTest::SelfTestOptions& options, SelfTest::SelfTestSuiteResult* outResult) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    Debug::Info(L"CompareSelfTest: begin");
    AppendCompareSelfTestTraceLine(L"Run: begin");

    SelfTest::SelfTestSuiteResult suite{};
    suite.suite = SelfTest::SelfTestSuite::CompareDirectories;

    std::wstring fatalSetupFailure;

    wil::com_ptr<IFileSystem> baseFs = GetLocalFileSystem();
    if (! baseFs)
    {
        fatalSetupFailure = L"CompareSelfTest: local file system plugin not available.";
    }

    const std::filesystem::path suiteRoot = SelfTest::GetTempRoot(SelfTest::SelfTestSuite::CompareDirectories);
    if (fatalSetupFailure.empty() && suiteRoot.empty())
    {
        fatalSetupFailure = L"CompareSelfTest: suite artifact root not available.";
    }

    const std::filesystem::path root = suiteRoot / L"work";
    if (fatalSetupFailure.empty() && ! SelfTest::EnsureDirectory(root))
    {
        fatalSetupFailure = L"CompareSelfTest: failed to create work root folder.";
    }

    if (! fatalSetupFailure.empty())
    {
        AppendCompareSelfTestTraceLine(L"Run: aborting due to setup failure");
        AppendCaseResult(suite, L"setup", SelfTest::SelfTestCaseResult::Status::failed, fatalSetupFailure);
        AppendSkippedCompareCasesForSetupFailure(options, suite, std::format(L"not executed due to suite setup failure: {}", fatalSetupFailure));
    }
    else
    {
        AppendCompareSelfTestTraceLine(L"Run: root created");
    }

    std::wstring guid = MakeGuidText();
    if (guid.empty())
    {
        guid = L"0";
    }

    wil::com_ptr<IFileSystem> dummyFs = GetDummyFileSystem();
    wil::com_ptr<IInformations> dummyInfo;
    wil::com_ptr<IFileSystemIO> dummyIo;
    wil::com_ptr<IFileSystemDirectoryOperations> dummyOps;

    if (fatalSetupFailure.empty())
    {
        std::wstring setupFailure;
        if (! dummyFs)
        {
            setupFailure = L"CompareSelfTest: FileSystemDummy plugin not available.";
        }
        else
        {
            AppendCompareSelfTestTraceLine(L"Run: dummy plugin setup");
            if (! CreateInformations(dummyFs, dummyInfo))
            {
                setupFailure = L"CompareSelfTest: FileSystemDummy missing IInformations.";
            }
            else
            {
                const HRESULT setHr =
                    dummyInfo->SetConfiguration("{\"maxChildrenPerDirectory\":0,\"maxDepth\":0,\"seed\":1,\"latencyMs\":0,\"virtualSpeedLimit\":\"0\"}");
                if (FAILED(setHr))
                {
                    setupFailure = L"CompareSelfTest: FileSystemDummy SetConfiguration failed.";
                }
            }

            if (setupFailure.empty() && ! CreateFileSystemIo(dummyFs, dummyIo))
            {
                setupFailure = L"CompareSelfTest: FileSystemDummy missing IFileSystemIO.";
            }
            if (setupFailure.empty() && ! CreateFileSystemDirectoryOperations(dummyFs, dummyOps))
            {
                setupFailure = L"CompareSelfTest: FileSystemDummy missing IFileSystemDirectoryOperations.";
            }
        }

        if (! setupFailure.empty())
        {
            AppendCaseResult(suite, L"setup", SelfTest::SelfTestCaseResult::Status::failed, setupFailure);
        }
    }

    if (fatalSetupFailure.empty())
    {
#include "CompareDirectoriesEngine.SelfTest.Cases.CoreDiffs.cpp"
#include "CompareDirectoriesEngine.SelfTest.Cases.RuntimeAndRemote.cpp"
#include "CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp"

        AppendCompareSelfTestTraceLine(L"Run: finalizing");

        const auto endedAt = std::chrono::steady_clock::now();
        suite.durationMs   = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::milliseconds>(endedAt - startedAt).count());

        if (outResult)
        {
            *outResult = suite;
        }

        if (options.writeJsonSummary)
        {
            const std::filesystem::path jsonPath = SelfTest::GetSuiteArtifactPath(SelfTest::SelfTestSuite::CompareDirectories, L"results.json");
            SelfTest::WriteSuiteJson(suite, jsonPath);
        }

        if (suite.failed != 0)
        {
            AppendCompareSelfTestTraceLine(L"Run: failed");
            Debug::Error(L"CompareSelfTest: failed.");
            return false;
        }

        AppendCompareSelfTestTraceLine(L"Run: passed");
        Debug::Info(L"CompareSelfTest: passed.");
        return true;
    }

#endif
