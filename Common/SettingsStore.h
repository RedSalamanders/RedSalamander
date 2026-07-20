#pragma once

#include "ThemeExpression.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <initializer_list>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <variant>
#include <vector>

#pragma warning(push)
// Windows headers: C4710 (not inlined), C4711 (auto inline), C4514 (unreferenced inline)
#pragma warning(disable : 4710 4711 4514)
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>
#pragma warning(pop)

#include "SettingsThumbnail.h"

#ifndef COMMON_API
#ifdef COMMON_EXPORTS
#define COMMON_API __declspec(dllexport)
#else
#define COMMON_API __declspec(dllimport)
#endif
#endif

namespace Common::Settings
{
struct WindowBounds
{
    int x      = 0;
    int y      = 0;
    int width  = 0;
    int height = 0;
};

enum class WindowState : uint8_t
{
    Normal,
    Maximized,
};

struct WindowPlacement
{
    WindowState state = WindowState::Normal;
    WindowBounds bounds{};
    std::optional<unsigned int> dpi;
};

enum class FolderDisplayMode : uint8_t
{
    Brief,
    Detailed,
    ExtraDetailed,
    Thumbnails,
};

enum class FolderSortBy : uint8_t
{
    Name,
    Extension,
    Time,
    Size,
    Attributes,
    None,
};

enum class FolderSortDirection : uint8_t
{
    Ascending,
    Descending,
};

struct FolderViewSettings
{
    FolderDisplayMode display         = FolderDisplayMode::Brief;
    FolderSortBy sortBy               = FolderSortBy::Name;
    FolderSortDirection sortDirection = FolderSortDirection::Ascending;
    bool fileExtensionsVisible        = true;
    uint32_t thumbnailSizeDip         = Thumbnail::kDefaultSizeDip;
    bool thumbnailsVisible            = false;
    bool navigationBarVisible         = true;
    bool filterBarVisible             = false;
    bool statusBarVisible             = true;
};

struct FolderPane
{
    std::wstring slot;
    std::filesystem::path current;
    FolderViewSettings view;
};

struct FolderLayoutSettings
{
    float splitRatio = 0.5f;
    std::optional<std::wstring> zoomedPane;
    std::optional<float> zoomRestoreSplitRatio;
};

struct FolderHistoryFilterState
{
    std::filesystem::path path;
    bool enabled = false;
    std::wstring text;
};

struct FoldersSettings
{
    std::wstring active;
    FolderLayoutSettings layout;
    bool showHiddenFiles = true;
    bool showSystemFiles = true;
    uint32_t historyMax  = 20u;
    std::vector<std::filesystem::path> history;
    std::vector<FolderHistoryFilterState> historyFilters;
    std::vector<FolderPane> items;
};

struct MonitorMenuState
{
    bool toolbarVisible     = true;
    bool lineNumbersVisible = true;
    bool alwaysOnTop        = false;
    bool showIds            = true;
    bool autoScroll         = true;
};

struct MainMenuState
{
    bool menuBarVisible     = true;
    bool functionBarVisible = true;
};

struct StartupSettings
{
    bool showSplash = true;
};

enum class ReducedMotionMode : uint8_t
{
    System,
    On,
    Off,
};

enum class WindowBackdropMode : uint8_t
{
    Default,
    None,
    Mica,
    MicaAlt,
    Acrylic,
};

struct UiSettings
{
    bool compactMode                  = true;
    ReducedMotionMode reducedMotion   = ReducedMotionMode::System;
    WindowBackdropMode windowBackdrop = WindowBackdropMode::Default;
    std::wstring language             = L"system";

    bool operator==(const UiSettings&) const noexcept = default;
};

enum class MonitorFilterPreset : uint8_t
{
    Custom,
    ErrorsOnly,
    ErrorsWarnings,
    AllTypes,
};

struct MonitorFilterState
{
    uint32_t mask              = 63u; // 0..63
    MonitorFilterPreset preset = MonitorFilterPreset::Custom;
};

struct MonitorRetentionSettings
{
    uint32_t maxQueuedEvents       = 4'096u;
    uint32_t maxRetainedLines      = 100'000u;
    uint64_t maxRetainedTextBytes  = 64u * 1024u * 1024u;
    uint32_t maxSearchMatches      = 100'000u;
};

struct MonitorSettings
{
    MonitorMenuState menu;
    MonitorFilterState filter;
    MonitorRetentionSettings retention;
};

struct DirectoryInfoCacheSettings
{
    std::optional<uint64_t> maxBytes;
    std::optional<uint32_t> maxWatchers;
    std::optional<uint32_t> mruWatched;
};

struct CacheSettings
{
    DirectoryInfoCacheSettings directoryInfo;
};

struct JsonArray;
struct JsonObject;

struct JsonValue
{
    using ArrayPtr  = std::shared_ptr<JsonArray>;
    using ObjectPtr = std::shared_ptr<JsonObject>;

    // Holds a JSON value:
    // - std::monostate: null
    // - bool
    // - int64_t / uint64_t / double
    // - std::string: string (UTF-8)
    // - JsonArray / JsonObject
    std::variant<std::monostate, bool, int64_t, uint64_t, double, std::string, ArrayPtr, ObjectPtr> value;
};

struct JsonArray
{
    std::vector<JsonValue> items;
};

struct JsonObject
{
    // Members are stored as UTF-8 to match JSON's encoding.
    std::vector<std::pair<std::string, JsonValue>> members;
};

struct ThemeDefinition
{
    uint32_t formatVersion = 2u;
    std::wstring id;
    std::wstring name;
    std::wstring baseThemeId; // builtin/* or an inline forward-compatible fallback id
    std::unordered_map<std::wstring, ThemeColorSource> palette;
    std::unordered_map<std::wstring, ThemeColorSource> colors;
};

struct ThemeSettings
{
    struct OpaqueEntry
    {
        size_t originalIndex = 0u;
        JsonValue value;
    };

    std::wstring currentThemeId = L"builtin/system";
    std::vector<ThemeDefinition> themes;
    // Structurally unusable and duplicate inline entries retain their JSON value and array position for repair-safe round trips.
    std::vector<OpaqueEntry> opaqueThemeEntries;
};

struct PluginsSettings
{
    // The active IFileSystem plugin (by PluginMetaData.id, long id).
    // Example: "builtin/file-system"
    std::wstring currentFileSystemPluginId = L"builtin/file-system";

    // Absolute paths to custom plugins (outside the application folder).
    std::vector<std::filesystem::path> customPluginPaths;

    // Plugins disabled by the user (by PluginMetaData.id).
    std::vector<std::wstring> disabledPluginIds;

    // Per-plugin configuration payloads as JSON values.
    // Key: PluginMetaData.id
    std::unordered_map<std::wstring, JsonValue> configurationByPluginId;
};

enum class ConnectionAuthMode : uint8_t
{
    Anonymous,
    Password,
    SshKey,
    OAuth2Pkce,
    OAuth2Interactive = OAuth2Pkce,
};

struct ConnectionProfile
{
    std::wstring id;       // stable internal GUID (used for WinCred storage; not used for /@conn navigation)
    std::wstring name;     // user-visible name (unique, case-insensitive)
    std::wstring pluginId; // PluginMetaData.id (long id)
    std::wstring host;
    uint32_t port            = 0;    // 0 = protocol default
    std::wstring initialPath = L"/"; // plugin path, typically '/'
    std::wstring userName;
    ConnectionAuthMode authMode = ConnectionAuthMode::Password;
    bool savePassword           = false;
    bool requireWindowsHello    = true;
    JsonValue extra; // plugin-specific non-secret fields (object recommended)
};

struct ConnectionsSettings
{
    std::vector<ConnectionProfile> items;
    bool bypassWindowsHello                  = false;
    bool allowInsecureTlsInAutomation        = false;
    uint32_t windowsHelloReauthTimeoutMinute = 10;
};

// Persisted connection IDs use one canonical representation so settings, WinCred
// targets, and in-memory authorization state cannot disagree about identity.
inline constexpr std::wstring_view kQuickConnectConnectionId = L"00000000-0000-0000-0000-000000000001";

COMMON_API HRESULT CreateConnectionProfileId(std::wstring& idOut) noexcept;
COMMON_API HRESULT NormalizeConnectionProfileId(std::wstring_view id, std::wstring& canonicalIdOut) noexcept;
COMMON_API HRESULT ValidateConnectionProfileIds(const ConnectionsSettings& settings) noexcept;

struct GridColumnLayoutEntry
{
    std::wstring columnId;
    uint32_t displayIndex = 0u;
    float widthDip        = 0.0f;
};

struct FileOperationsSettings
{
    bool autoDismissSuccess                      = false;
    bool popupFooterOnly                         = false;
    bool popupCompactDensity                     = false;
    bool preCalcEnabled                          = true;
    uint32_t preCalcMaxWorkers                   = 4;
    uint32_t crossFsBridgeBufferSizeKB           = 4096;
    uint64_t defaultBandwidthLimitBytesPerSecond = 0;
    uint32_t maxDiagnosticsLogFiles              = 14;
    // Diagnostics verbosity: by default, Debug builds keep more context while Release builds stay lean.
#if defined(_DEBUG) || defined(DEBUG)
    bool diagnosticsInfoEnabled  = true;
    bool diagnosticsDebugEnabled = true;
#else
    bool diagnosticsInfoEnabled  = false;
    bool diagnosticsDebugEnabled = false;
#endif
    std::optional<uint32_t> maxIssueReportFiles;
    std::optional<uint32_t> maxDiagnosticsInMemory;
    std::optional<uint32_t> maxDiagnosticsPerFlush;
    std::optional<uint32_t> diagnosticsFlushIntervalMs;
    std::optional<uint32_t> diagnosticsCleanupIntervalMs;
    std::wstring issuesPaneSortColumnId;
    bool issuesPaneSortDescending = false;
    std::vector<GridColumnLayoutEntry> issuesPaneGridLayout;
};

[[nodiscard]] COMMON_API bool HasNonDefaultFileOperationsSettings(const FileOperationsSettings& fileOperations) noexcept;

struct CompareDirectoriesSettings
{
    bool compareSize       = false;
    bool compareDateTime   = false;
    bool compareAttributes = false;
    bool compareContent    = false;

    bool compareSubdirectories         = false;
    bool compareSubdirectoryAttributes = false;
    bool selectSubdirsOnlyInOnePane    = true;

    bool ignoreFiles = false;
    std::wstring ignoreFilesPatterns;
    bool ignoreDirectories = false;
    std::wstring ignoreDirectoriesPatterns;

    // Keep identical entries in cached decisions so "Show Identical Items" can be toggled as a view filter (no rescan).
    // Uses more memory on very large trees.
    bool keepIdenticalItems = false;

    // Display identical items (requires keepIdenticalItems).
    bool showIdenticalItems = false;

    // 0 = Auto (current behavior, clamped to <= 4), 1..4 = fixed worker count for background content compare.
    uint32_t contentCompareWorkerCount = 0;
};

struct HotPathSlot
{
    std::wstring path;
    std::wstring label; // empty = derive display name from path
    bool showInMenu = false;
};

struct HotPathsSettings
{
    std::array<std::optional<HotPathSlot>, 10> slots{}; // [0]=Ctrl+1 .. [9]=Ctrl+0
    bool openPrefsOnAssign = false;
};

struct SelectionMasksSettings
{
    // Most-recent-first history (max 10 recommended).
    std::vector<std::wstring> selectHistory;
    std::vector<std::wstring> unselectHistory;
    std::vector<std::wstring> filterHistory;
};

enum class SearchNameMode : uint8_t
{
    Wildcard,
    Literal,
    Regex,
};

enum class SearchContentMode : uint8_t
{
    Disabled,
    TextLiteral,
    TextRegex,
};

struct SearchDialogSettings
{
    std::vector<std::wstring> recentRoots;
    std::vector<std::wstring> recentNamePatterns;
    std::vector<std::wstring> recentContentPatterns;

    std::wstring lastRoot;
    std::wstring lastNamePattern;
    std::wstring lastContentPattern;

    bool recursive          = true;
    bool includeFiles       = true;
    bool includeDirectories = false;
    bool followSymlinks     = false;
    bool matchCaseName      = false;
    bool matchCaseContent   = false;
    bool preferIndex        = true;
    bool wantSnippets       = false;

    SearchNameMode nameMode       = SearchNameMode::Wildcard;
    SearchContentMode contentMode = SearchContentMode::Disabled;

    uint64_t maxResults = 0;
    std::wstring sortColumnId;
    bool sortDescending = false;
    std::vector<GridColumnLayoutEntry> resultsGridLayout;
};

enum class BatchRenameCaseStyle : uint8_t
{
    None,
    Lower,
    Upper,
    Mixed,
};

struct BatchRenameSettings
{
    std::wstring lastRoot;
    std::vector<std::wstring> recentMasks;
    std::vector<std::wstring> recentNameTemplates;
    std::vector<std::wstring> recentSearchPatterns;
    std::vector<std::wstring> recentReplacePatterns;

    bool includeSubdirectories = false;
    bool includeFiles          = true;
    bool includeFolders        = false;
    bool regexEnabled          = false;
    bool caseSensitive         = true;
    bool wholeWords            = false;
    bool replaceOnce           = false;
    bool excludeExtension      = false;

    std::wstring flattenSeparator           = L" - ";
    BatchRenameCaseStyle fileNameCaseStyle  = BatchRenameCaseStyle::None;
    BatchRenameCaseStyle extensionCaseStyle = BatchRenameCaseStyle::None;

    std::wstring previewSortColumnId;
    bool previewSortDescending = false;
    std::vector<GridColumnLayoutEntry> previewGridLayout;

    bool operator==(const BatchRenameSettings&) const = default;
};

enum class FileActionKind : uint8_t
{
    ViewerPlugin,
    ExternalProgram,
};

enum class FileActionMatchKind : uint8_t
{
    Default,
    Extension,
    Pattern,
};

struct FileActionMatch
{
    FileActionMatchKind kind = FileActionMatchKind::Default;
    std::wstring value;

    bool operator==(const FileActionMatch&) const = default;
};

struct FileActionDefinition
{
    std::wstring id;
    std::wstring displayName;
    bool enabled        = true;
    FileActionKind kind = FileActionKind::ExternalProgram;
    std::wstring pluginId;
    std::wstring executablePath;
    std::wstring arguments;
    std::wstring workingDirectory;
    struct Applicability
    {
        std::vector<FileActionMatch> matches;
        std::vector<std::wstring> computerNames;

        bool operator==(const Applicability&) const = default;
    } appliesTo;

    bool operator==(const FileActionDefinition&) const = default;
};

struct ViewerAssociationRule
{
    FileActionMatch match;
    std::wstring computerName;
    std::wstring viewActionId;
    std::wstring alternateViewActionId;

    bool operator==(const ViewerAssociationRule&) const = default;
};

struct EditorAssociationRule
{
    FileActionMatch match;
    std::wstring computerName;
    std::wstring editActionId;
    std::wstring alternateEditActionId;
    std::wstring editNewActionId;

    bool operator==(const EditorAssociationRule&) const = default;
};

struct ViewerFileActionsSettings
{
    std::vector<FileActionDefinition> actions;
    std::vector<ViewerAssociationRule> associations;

    bool operator==(const ViewerFileActionsSettings&) const = default;
};

struct EditorFileActionsSettings
{
    std::vector<FileActionDefinition> actions;
    std::vector<EditorAssociationRule> associations;

    bool operator==(const EditorFileActionsSettings&) const = default;
};

struct FileActionsSettings
{
    ViewerFileActionsSettings viewers;
    EditorFileActionsSettings editors;

    bool operator==(const FileActionsSettings&) const = default;
};

struct UserMenuSettings
{
    std::vector<FileActionDefinition> actions;

    bool operator==(const UserMenuSettings&) const = default;
};

inline FileActionDefinition MakeViewerPluginAction(std::wstring id,
                                                   std::wstring displayName,
                                                   std::initializer_list<const wchar_t*> extensions,
                                                   bool defaultMatch = false)
{
    FileActionDefinition action{};
    action.id          = std::move(id);
    action.displayName = std::move(displayName);
    action.enabled     = true;
    action.kind        = FileActionKind::ViewerPlugin;
    action.pluginId    = action.id;
    if (defaultMatch)
    {
        action.appliesTo.matches.push_back(FileActionMatch{.kind = FileActionMatchKind::Default});
    }
    action.appliesTo.matches.reserve(action.appliesTo.matches.size() + extensions.size());
    for (const wchar_t* extension : extensions)
    {
        if (extension && extension[0] != L'\0')
        {
            action.appliesTo.matches.push_back(FileActionMatch{.kind = FileActionMatchKind::Extension, .value = extension});
        }
    }
    return action;
}

inline void AddViewerAssociationExtensions(ViewerFileActionsSettings& settings, std::wstring_view actionId, std::initializer_list<const wchar_t*> extensions)
{
    for (const wchar_t* extension : extensions)
    {
        if (! extension || extension[0] == L'\0')
        {
            continue;
        }

        ViewerAssociationRule rule{};
        rule.match.kind   = FileActionMatchKind::Extension;
        rule.match.value  = extension;
        rule.viewActionId = std::wstring(actionId);
        settings.associations.push_back(std::move(rule));
    }
}

inline ViewerFileActionsSettings DefaultViewerFileActionsSettings()
{
    ViewerFileActionsSettings settings{};

    const std::initializer_list<const wchar_t*> kTextExtensions{L".txt", L".log", L".xml", L".ini", L".cfg", L".csv", L".diff", L".patch", L".rej"};
    const std::initializer_list<const wchar_t*> kMarkdownExtensions{L".md"};
    const std::initializer_list<const wchar_t*> kJsonExtensions{L".json", L".json5", L".jsonl", L".ndjson"};
    const std::initializer_list<const wchar_t*> kWebExtensions{L".html", L".htm", L".pdf"};
    const std::initializer_list<const wchar_t*> kSqliteExtensions{L".db", L".db3", L".s3db", L".sqlite", L".sqlite3"};
    const std::initializer_list<const wchar_t*> kImageExtensions{
        L".bmp", L".dib", L".gif", L".ico",  L".jpe", L".jpeg", L".jpg", L".png", L".tif",  L".tiff", L".hdp", L".jxr", L".wdp", L".3fr",
        L".ari", L".arw", L".bay", L".braw", L".cap", L".cr2",  L".cr3", L".crw", L".data", L".dcr",  L".dcs", L".dng", L".drf", L".eip",
        L".erf", L".fff", L".gpr", L".iiq",  L".k25", L".kdc",  L".mdc", L".mef", L".mos",  L".mrw",  L".nef", L".nrw", L".obm", L".orf",
        L".pef", L".ptx", L".pxn", L".r3d",  L".raf", L".raw",  L".rwl", L".rw2", L".rwz",  L".sr2",  L".srf", L".srw", L".x3f"};
    const std::initializer_list<const wchar_t*> kVideoExtensions{L".avi", L".mp4",  L".mkv",  L".mka", L".mov",  L".wmv",  L".flv", L".mpg", L".mpeg",
                                                                 L".m4v", L".webm", L".3gp",  L".ts",  L".m2ts", L".mts",  L".vob", L".ogv", L".m4a",
                                                                 L".mp3", L".aac",  L".flac", L".wav", L".ogg",  L".opus", L".wma", L".aif", L".aiff"};
    const std::initializer_list<const wchar_t*> kPeExtensions{L".cpl", L".dll", L".drv", L".exe", L".ocx", L".scr", L".spl", L".sys"};

    settings.actions.reserve(8u);
    settings.actions.push_back(MakeViewerPluginAction(L"builtin/viewer-text", L"Text Viewer", kTextExtensions, true));
    settings.actions.push_back(MakeViewerPluginAction(L"builtin/viewer-markdown", L"Markdown Viewer", kMarkdownExtensions));
    settings.actions.push_back(MakeViewerPluginAction(L"builtin/viewer-json", L"JSON Viewer", kJsonExtensions));
    settings.actions.push_back(MakeViewerPluginAction(L"builtin/viewer-web", L"Web Viewer", kWebExtensions));
    settings.actions.push_back(MakeViewerPluginAction(L"builtin/viewer-sqlite", L"SQLite Viewer", kSqliteExtensions));
    settings.actions.push_back(MakeViewerPluginAction(L"builtin/viewer-imgraw", L"Image Viewer", kImageExtensions));
    settings.actions.push_back(MakeViewerPluginAction(L"builtin/viewer-vlc", L"Media Viewer", kVideoExtensions));
    settings.actions.push_back(MakeViewerPluginAction(L"builtin/viewer-pe", L"PE Viewer", kPeExtensions));

    settings.associations.reserve(1u + kTextExtensions.size() + kMarkdownExtensions.size() + kJsonExtensions.size() + kWebExtensions.size() +
                                  kSqliteExtensions.size() + kImageExtensions.size() + kVideoExtensions.size() + kPeExtensions.size());
    ViewerAssociationRule defaultRule{};
    defaultRule.match.kind   = FileActionMatchKind::Default;
    defaultRule.viewActionId = L"builtin/viewer-text";
    settings.associations.push_back(std::move(defaultRule));

    AddViewerAssociationExtensions(settings, L"builtin/viewer-text", kTextExtensions);
    AddViewerAssociationExtensions(settings, L"builtin/viewer-markdown", kMarkdownExtensions);
    AddViewerAssociationExtensions(settings, L"builtin/viewer-json", kJsonExtensions);
    AddViewerAssociationExtensions(settings, L"builtin/viewer-web", kWebExtensions);
    AddViewerAssociationExtensions(settings, L"builtin/viewer-sqlite", kSqliteExtensions);
    AddViewerAssociationExtensions(settings, L"builtin/viewer-imgraw", kImageExtensions);
    AddViewerAssociationExtensions(settings, L"builtin/viewer-vlc", kVideoExtensions);
    AddViewerAssociationExtensions(settings, L"builtin/viewer-pe", kPeExtensions);

    return settings;
}

inline EditorFileActionsSettings DefaultEditorFileActionsSettings()
{
    return {};
}

inline FileActionsSettings DefaultFileActionsSettings()
{
    FileActionsSettings settings{};
    settings.viewers = DefaultViewerFileActionsSettings();
    settings.editors = DefaultEditorFileActionsSettings();
    return settings;
}

enum class MakeFileListSourceMode : uint8_t
{
    Selection,
    CurrentFolder,
};

enum class MakeFileListFormat : uint8_t
{
    Text,
    Csv,
    Json,
};

enum class MakeFileListOutputTarget : uint8_t
{
    Clipboard,
    File,
};

struct MakeFileListSettings
{
    MakeFileListSourceMode sourceMode     = MakeFileListSourceMode::Selection;
    bool recursive                        = false;
    MakeFileListFormat format             = MakeFileListFormat::Text;
    MakeFileListOutputTarget outputTarget = MakeFileListOutputTarget::Clipboard;
    std::wstring textMacro                = L"{fullPath}\t{size}\t{modified}";
    std::filesystem::path outputFile;
    bool includeName        = true;
    bool includeFullPath    = true;
    bool includeSize        = true;
    bool includeModified    = true;
    bool includeAttributes  = false;
    bool includeDirectories = true;

    bool operator==(const MakeFileListSettings&) const = default;
};

struct ExtensionsSettings
{
    // Map a file extension (lowercase, with leading dot like ".7z") to a file system plugin ID.
    // Used by the host to open matching files as a virtual file system instead of ShellExecute.
    std::unordered_map<std::wstring, std::wstring> openWithFileSystemByExtension{
        // read / write
        {L".7z", L"builtin/file-system-7z"},
        {L".zip", L"builtin/file-system-7z"},
        {L".rar", L"builtin/file-system-7z"},
        {L".xz", L"builtin/file-system-7z"},
        {L".bzip2", L"builtin/file-system-7z"},
        {L".gzip", L"builtin/file-system-7z"},
        {L".tar", L"builtin/file-system-7z"},
        {L".wim", L"builtin/file-system-7z"},
        // read only
        {L".rar", L"builtin/file-system-7z"},
        {L".apfs", L"builtin/file-system-7z"},
        {L".ar", L"builtin/file-system-7z"},
        {L".arj", L"builtin/file-system-7z"},
        {L".cab", L"builtin/file-system-7z"},
        {L".chm", L"builtin/file-system-7z"},
        {L".cpio", L"builtin/file-system-7z"},
        {L".cramfs", L"builtin/file-system-7z"},
        {L".dmg", L"builtin/file-system-7z"},
        {L".ext", L"builtin/file-system-7z"},
        {L".fat", L"builtin/file-system-7z"},
        {L".gpt", L"builtin/file-system-7z"},
        {L".hfs", L"builtin/file-system-7z"},
        {L".ihex", L"builtin/file-system-7z"},
        {L".iso", L"builtin/file-system-7z"},
        {L".lzh", L"builtin/file-system-7z"},
        {L".lzma", L"builtin/file-system-7z"},
        {L".mbr", L"builtin/file-system-7z"},
        {L".msi", L"builtin/file-system-7z"},
        {L".nsis", L"builtin/file-system-7z"},
        {L".ntfs", L"builtin/file-system-7z"},
        {L".qcow2", L"builtin/file-system-7z"},
        {L".rar", L"builtin/file-system-7z"},
        {L".rpm", L"builtin/file-system-7z"},
        {L".squashfs", L"builtin/file-system-7z"},
        {L".udf", L"builtin/file-system-7z"},
        {L".uefi", L"builtin/file-system-7z"},
        {L".vdi", L"builtin/file-system-7z"},
        {L".vhd", L"builtin/file-system-7z"},
        {L".vhdx", L"builtin/file-system-7z"},
        {L".vmdk", L"builtin/file-system-7z"},
        {L".xar", L"builtin/file-system-7z"},
        {L".z", L"builtin/file-system-7z"},
    };
};

struct ShortcutBinding
{
    uint32_t vk        = 0; // Win32 virtual-key code (0..255 recommended)
    uint32_t modifiers = 0; // bitmask: 1=Ctrl, 2=Alt, 4=Shift
    std::wstring commandId;
};

struct ShortcutsSettings
{
    std::vector<ShortcutBinding> functionBar;
    std::vector<ShortcutBinding> folderView;
    bool functionBarCollapsed = false;
    bool folderViewCollapsed  = false;
    std::wstring sortColumnId;
    bool sortDescending = false;
    std::vector<GridColumnLayoutEntry> gridLayout;
};

enum class SettingsSavePermission : uint8_t
{
    Automatic,
    ExplicitReplacementRequired,
};

struct SettingsFileStamp
{
    uint32_t volumeSerialNumber = 0;
    uint64_t fileIndexHigh      = 0;
    uint64_t fileIndexLow       = 0;
    uint64_t lastWriteTime      = 0;
    uint64_t fileSize           = 0;

    bool operator==(const SettingsFileStamp&) const noexcept = default;
};

struct SettingsPersistenceState
{
    // Unknown top-level members and malformed optional sections are copied out of the yyjson
    // document so a later canonical save does not silently discard data this build cannot use.
    std::vector<std::pair<std::string, JsonValue>> opaqueTopLevelMembers;
    // Identity of the canonical save target when this snapshot was loaded or last committed.
    // nullopt means the target did not exist. Writers compare this under the cross-process
    // commit lock and fail with ERROR_REVISION_MISMATCH instead of overwriting a newer file.
    std::optional<SettingsFileStamp> expectedFileStamp;
    SettingsSavePermission savePermission = SettingsSavePermission::Automatic;
    int64_t sourceSchemaVersion           = 16;
};

struct Settings
{
    uint32_t schemaVersion = 16;
    std::unordered_map<std::wstring, WindowPlacement> windows;
    ThemeSettings theme;
    PluginsSettings plugins;
    FileActionsSettings fileActions = DefaultFileActionsSettings();
    UserMenuSettings userMenu;
    std::optional<MakeFileListSettings> makeFileList;
    ExtensionsSettings extensions;
    std::optional<ShortcutsSettings> shortcuts;
    std::optional<MainMenuState> mainMenu;
    std::optional<StartupSettings> startup;
    std::optional<UiSettings> ui;
    std::optional<CacheSettings> cache;
    std::optional<FoldersSettings> folders;
    std::optional<MonitorSettings> monitor;
    std::optional<ConnectionsSettings> connections;
    std::optional<FileOperationsSettings> fileOperations;
    std::optional<CompareDirectoriesSettings> compareDirectories;
    std::optional<HotPathsSettings> hotPaths;
    std::optional<SelectionMasksSettings> selectionMasks;
    std::optional<SearchDialogSettings> search;
    std::optional<BatchRenameSettings> batchRename;
    SettingsPersistenceState persistence;
};

enum class SettingsLoadRecoveryReason : uint8_t
{
    None,
    SettingsFileMissing,
    ReadFailed,
    InvalidJson,
    InvalidRoot,
    MissingSchemaVersion,
    UnsupportedSchemaVersion,
    LegacyShape,
    FileActionsInvalid,
    UserMenuInvalid,
    ShortcutsInvalid,
    ConnectionProfileIdsMigrated,
};

struct ConnectionProfileIdMigration
{
    std::wstring profileName;
    std::wstring previousId;
    std::wstring replacementId;
    bool savedSecretReferenceCleared = false;
};

struct SettingsSectionRecoveryInfo
{
    SettingsLoadRecoveryReason reason = SettingsLoadRecoveryReason::None;
    HRESULT hr                        = S_OK;
};

struct SettingsLoadRecoveryInfo
{
    SettingsLoadRecoveryReason reason = SettingsLoadRecoveryReason::None;
    HRESULT hr                        = S_OK;
    std::filesystem::path settingsPath;
    std::filesystem::path backupPath;
    int64_t unsupportedSchemaVersion = 0;
    bool usedDefaults                = false;
    bool backedUp                    = false;
    std::vector<SettingsSectionRecoveryInfo> sectionRecoveries;
    std::vector<ConnectionProfileIdMigration> connectionProfileIdMigrations;
};

COMMON_API std::filesystem::path GetSettingsPath(std::wstring_view appId) noexcept;
COMMON_API std::filesystem::path GetSettingsSchemaPath(std::wstring_view appId) noexcept;

// Returns the canonical Settings Store JSON Schema (UTF-8 JSON, no BOM).
COMMON_API std::string_view GetSettingsStoreSchemaJsonUtf8() noexcept;

// Returns:
// - S_OK: loaded successfully
// - S_FALSE: defaults used (missing/invalid/unreadable)
COMMON_API HRESULT LoadSettings(std::wstring_view appId, Settings& out) noexcept;

// Same recovery behavior as LoadSettings(...), but reports when an existing
// settings file was backed up and defaults were restored.
COMMON_API HRESULT LoadSettingsWithRecoveryInfo(std::wstring_view appId, Settings& out, SettingsLoadRecoveryInfo* recovery) noexcept;

// Returns:
// - S_OK: loaded successfully
// - S_FALSE: file missing
// - failure HRESULT: file unreadable / invalid JSON / unsupported schema version
//
// Unlike LoadSettings(...), this does not back up bad files and does not silently
// fall back to defaults when the on-disk file is invalid.
COMMON_API HRESULT TryLoadSettingsNoRecovery(std::wstring_view appId, Settings& out) noexcept;

// Returns:
// - S_OK: stamp retrieved successfully
// - S_FALSE: file missing
// - failure HRESULT: unexpected I/O error while querying the file
COMMON_API HRESULT TryGetSettingsFileStamp(std::wstring_view appId, SettingsFileStamp& out) noexcept;

// Moves the current settings file to the standard timestamped backup path. This is the required
// first step for an explicit replacement of settings written by a newer schema version.
COMMON_API HRESULT BackupSettingsForExplicitReplacement(std::wstring_view appId, std::filesystem::path& backupPath) noexcept;

// On success the snapshot's expectedFileStamp advances to the committed file identity.
COMMON_API HRESULT SaveSettings(std::wstring_view appId, Settings& settings) noexcept;
// One-shot compatibility overload for immutable shutdown/test snapshots. Prefer the mutable
// overload whenever the caller can save the same in-memory snapshot again.
COMMON_API HRESULT SaveSettings(std::wstring_view appId, const Settings& settings) noexcept;

// Saves only the settings JSON. Use when an existing schema remains valid and must not be regenerated.
COMMON_API HRESULT SaveSettingsValuesOnly(std::wstring_view appId, Settings& settings) noexcept;

// Saves only the settings JSON and returns the identity of the exact flushed temporary file moved
// into place. Callers that watch the settings directory use this to distinguish their own atomic
// replacement from a later external replacement without re-statting the path.
COMMON_API HRESULT SaveSettingsValuesOnlyWithStamp(std::wstring_view appId,
                                                   Settings& settings,
                                                   SettingsFileStamp& writtenStamp) noexcept;

// Saves a JSON Schema file alongside the settings file (UTF-8 JSON, no BOM).
// Path: `<AppId>.settings.schema.json` in the same directory as `GetSettingsPath(appId)`.
COMMON_API HRESULT SaveSettingsSchema(std::wstring_view appId, std::string_view schemaJsonUtf8) noexcept;

// Parses a JSON/JSON5 value from UTF-8 text.
// Returns S_OK on success, otherwise an HRESULT error.
COMMON_API HRESULT ParseJsonValue(std::string_view jsonText, JsonValue& out) noexcept;

// Strict typed accessors for parsed JSON objects. Numeric reads reject negative and overflowing
// values; UTF-16 member reads reject malformed UTF-8 instead of substituting replacement characters.
COMMON_API const JsonValue* FindMember(const JsonValue& value, std::string_view key) noexcept;
COMMON_API std::optional<std::string> GetString(const JsonValue& value, std::string_view key) noexcept;
COMMON_API std::optional<std::wstring> GetWString(const JsonValue& value, std::string_view key) noexcept;
COMMON_API std::optional<bool> GetBool(const JsonValue& value, std::string_view key) noexcept;
COMMON_API std::optional<uint32_t> GetUInt32(const JsonValue& value, std::string_view key) noexcept;
COMMON_API const JsonArray* GetArray(const JsonValue& value, std::string_view key) noexcept;

// Serializes a JSON value to UTF-8 JSON text.
// Returns S_OK on success, otherwise an HRESULT error.
COMMON_API HRESULT SerializeJsonValue(const JsonValue& value, std::string& outJsonText) noexcept;

COMMON_API bool TryParseColor(std::wstring_view hex, uint32_t& argb) noexcept;
COMMON_API std::wstring FormatColor(uint32_t argb);

// Normalizes the placement for the current monitor configuration:
// - Optionally scales width/height by currentDpi/savedDpi
// - Ensures the window is fully visible inside a monitor work area
COMMON_API WindowPlacement NormalizeWindowPlacement(const WindowPlacement& saved, unsigned int currentDpi) noexcept;

// Loads theme definitions from a directory (expects files named `*.theme.json5`).
// Returns:
// - S_OK: one or more themes loaded
// - S_FALSE: directory missing/empty, or no valid themes
COMMON_API HRESULT LoadThemeDefinitionsFromDirectory(const std::filesystem::path& directory, std::vector<ThemeDefinition>& out) noexcept;
} // namespace Common::Settings
