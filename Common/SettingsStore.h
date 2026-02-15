#pragma once

#include <cstdint>
#include <filesystem>
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

struct ThemeDefinition
{
    std::wstring id;
    std::wstring name;
    std::wstring baseThemeId;                          // builtin/*
    std::unordered_map<std::wstring, uint32_t> colors; // key -> 0xAARRGGBB
};

struct ThemeSettings
{
    std::wstring currentThemeId = L"builtin/system";
    std::vector<ThemeDefinition> themes;
};

enum class FolderDisplayMode : uint8_t
{
    Brief,
    Detailed,
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

struct FoldersSettings
{
    std::wstring active;
    FolderLayoutSettings layout;
    uint32_t historyMax = 20u;
    std::vector<std::filesystem::path> history;
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

enum class MonitorFilterPreset : uint8_t
{
    Custom,
    ErrorsOnly,
    ErrorsWarnings,
    AllTypes,
};

struct MonitorFilterState
{
    uint32_t mask              = 31u; // 0..31
    MonitorFilterPreset preset = MonitorFilterPreset::Custom;
};

struct MonitorSettings
{
    MonitorMenuState menu;
    MonitorFilterState filter;
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
    uint32_t windowsHelloReauthTimeoutMinute = 10;
};

 struct FileOperationsSettings
 {
     bool autoDismissSuccess         = false;
     uint32_t maxDiagnosticsLogFiles = 14;
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
 };

 struct CompareDirectoriesSettings
 {
     bool compareSize       = false;
     bool compareDateTime   = false;
     bool compareAttributes = false;
     bool compareContent    = false;
 
     bool compareSubdirectories          = false;
     bool compareSubdirectoryAttributes  = false;
     bool selectSubdirsOnlyInOnePane     = true;
 
     bool ignoreFiles       = false;
     std::wstring ignoreFilesPatterns;
     bool ignoreDirectories = false;
     std::wstring ignoreDirectoriesPatterns;
 
     bool showIdenticalItems = false; // Show full list (not just differences).
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

    // Map a file extension (lowercase, with leading dot like ".txt") to a viewer plugin ID.
    // Used by the host to open matching files in a viewer window on F3.
    std::unordered_map<std::wstring, std::wstring> openWithViewerByExtension{
        {L".txt", L"builtin/viewer-text"},
        {L".log", L"builtin/viewer-text"},
        {L".md", L"builtin/viewer-markdown"},
        {L".json", L"builtin/viewer-json"},
        {L".json5", L"builtin/viewer-json"},
        {L".html", L"builtin/viewer-web"},
        {L".htm", L"builtin/viewer-web"},
        {L".pdf", L"builtin/viewer-web"},
        {L".xml", L"builtin/viewer-text"},
        {L".ini", L"builtin/viewer-text"},
        {L".cfg", L"builtin/viewer-text"},
        {L".csv", L"builtin/viewer-text"},

        // Default image formats (built-in WIC codecs)
        {L".bmp", L"builtin/viewer-imgraw"},
        {L".dib", L"builtin/viewer-imgraw"},
        {L".gif", L"builtin/viewer-imgraw"},
        {L".ico", L"builtin/viewer-imgraw"},
        {L".jpe", L"builtin/viewer-imgraw"},
        {L".jpeg", L"builtin/viewer-imgraw"},
        {L".jpg", L"builtin/viewer-imgraw"},
        {L".png", L"builtin/viewer-imgraw"},
        {L".tif", L"builtin/viewer-imgraw"},
        {L".tiff", L"builtin/viewer-imgraw"},
        {L".hdp", L"builtin/viewer-imgraw"},
        {L".jxr", L"builtin/viewer-imgraw"},
        {L".wdp", L"builtin/viewer-imgraw"},

        // Default video formats (VLC / libVLC)
        {L".avi", L"builtin/viewer-vlc"},
        {L".mp4", L"builtin/viewer-vlc"},
        {L".mkv", L"builtin/viewer-vlc"},
        {L".mka", L"builtin/viewer-vlc"},
        {L".mov", L"builtin/viewer-vlc"},
        {L".wmv", L"builtin/viewer-vlc"},
        {L".flv", L"builtin/viewer-vlc"},
        {L".mpg", L"builtin/viewer-vlc"},
        {L".mpeg", L"builtin/viewer-vlc"},
        {L".m4v", L"builtin/viewer-vlc"},
        {L".webm", L"builtin/viewer-vlc"},
        {L".3gp", L"builtin/viewer-vlc"},
        {L".ts", L"builtin/viewer-vlc"},
        {L".m2ts", L"builtin/viewer-vlc"},
        {L".mts", L"builtin/viewer-vlc"},
        {L".vob", L"builtin/viewer-vlc"},
        {L".ogv", L"builtin/viewer-vlc"},
        {L".m4a", L"builtin/viewer-vlc"},
        {L".mp3", L"builtin/viewer-vlc"},
        {L".aac", L"builtin/viewer-vlc"},
        {L".flac", L"builtin/viewer-vlc"},
        {L".wav", L"builtin/viewer-vlc"},
        {L".ogg", L"builtin/viewer-vlc"},
        {L".opus", L"builtin/viewer-vlc"},
        {L".wma", L"builtin/viewer-vlc"},
        {L".aif", L"builtin/viewer-vlc"},
        {L".aiff", L"builtin/viewer-vlc"},

        // Portable Executable formats (PE)
        {L".cpl", L"builtin/viewer-pe"},
        {L".dll", L"builtin/viewer-pe"},
        {L".drv", L"builtin/viewer-pe"},
        {L".exe", L"builtin/viewer-pe"},
        {L".ocx", L"builtin/viewer-pe"},
        {L".scr", L"builtin/viewer-pe"},
        {L".spl", L"builtin/viewer-pe"},
        {L".sys", L"builtin/viewer-pe"},

        // RAW camera formats (LibRaw)
        {L".3fr", L"builtin/viewer-imgraw"},
        {L".ari", L"builtin/viewer-imgraw"},
        {L".arw", L"builtin/viewer-imgraw"},
        {L".bay", L"builtin/viewer-imgraw"},
        {L".braw", L"builtin/viewer-imgraw"},
        {L".cap", L"builtin/viewer-imgraw"},
        {L".cr2", L"builtin/viewer-imgraw"},
        {L".cr3", L"builtin/viewer-imgraw"},
        {L".crw", L"builtin/viewer-imgraw"},
        {L".data", L"builtin/viewer-imgraw"},
        {L".dcr", L"builtin/viewer-imgraw"},
        {L".dcs", L"builtin/viewer-imgraw"},
        {L".dng", L"builtin/viewer-imgraw"},
        {L".drf", L"builtin/viewer-imgraw"},
        {L".eip", L"builtin/viewer-imgraw"},
        {L".erf", L"builtin/viewer-imgraw"},
        {L".fff", L"builtin/viewer-imgraw"},
        {L".gpr", L"builtin/viewer-imgraw"},
        {L".iiq", L"builtin/viewer-imgraw"},
        {L".k25", L"builtin/viewer-imgraw"},
        {L".kdc", L"builtin/viewer-imgraw"},
        {L".mdc", L"builtin/viewer-imgraw"},
        {L".mef", L"builtin/viewer-imgraw"},
        {L".mos", L"builtin/viewer-imgraw"},
        {L".mrw", L"builtin/viewer-imgraw"},
        {L".nef", L"builtin/viewer-imgraw"},
        {L".nrw", L"builtin/viewer-imgraw"},
        {L".obm", L"builtin/viewer-imgraw"},
        {L".orf", L"builtin/viewer-imgraw"},
        {L".pef", L"builtin/viewer-imgraw"},
        {L".ptx", L"builtin/viewer-imgraw"},
        {L".pxn", L"builtin/viewer-imgraw"},
        {L".r3d", L"builtin/viewer-imgraw"},
        {L".raf", L"builtin/viewer-imgraw"},
        {L".raw", L"builtin/viewer-imgraw"},
        {L".rwl", L"builtin/viewer-imgraw"},
        {L".rw2", L"builtin/viewer-imgraw"},
        {L".rwz", L"builtin/viewer-imgraw"},
        {L".sr2", L"builtin/viewer-imgraw"},
        {L".srf", L"builtin/viewer-imgraw"},
        {L".srw", L"builtin/viewer-imgraw"},
        {L".x3f", L"builtin/viewer-imgraw"},
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
};

 struct Settings
 {
     uint32_t schemaVersion = 9;
     std::unordered_map<std::wstring, WindowPlacement> windows;
     ThemeSettings theme;
     PluginsSettings plugins;
     ExtensionsSettings extensions;
     std::optional<ShortcutsSettings> shortcuts;
     std::optional<MainMenuState> mainMenu;
     std::optional<StartupSettings> startup;
     std::optional<CacheSettings> cache;
     std::optional<FoldersSettings> folders;
     std::optional<MonitorSettings> monitor;
     std::optional<ConnectionsSettings> connections;
     std::optional<FileOperationsSettings> fileOperations;
     std::optional<CompareDirectoriesSettings> compareDirectories;
 };

COMMON_API std::filesystem::path GetSettingsPath(std::wstring_view appId) noexcept;
COMMON_API std::filesystem::path GetSettingsSchemaPath(std::wstring_view appId) noexcept;

// Returns the canonical Settings Store JSON Schema (UTF-8 JSON, no BOM).
COMMON_API std::string_view GetSettingsStoreSchemaJsonUtf8() noexcept;

// Returns:
// - S_OK: loaded successfully
// - S_FALSE: defaults used (missing/invalid/unreadable)
COMMON_API HRESULT LoadSettings(std::wstring_view appId, Settings& out) noexcept;

COMMON_API HRESULT SaveSettings(std::wstring_view appId, const Settings& settings) noexcept;

// Saves a JSON Schema file alongside the settings file (UTF-8 JSON, no BOM).
// Path: `<AppId>.settings.schema.json` in the same directory as `GetSettingsPath(appId)`.
COMMON_API HRESULT SaveSettingsSchema(std::wstring_view appId, std::string_view schemaJsonUtf8) noexcept;

// Parses a JSON/JSON5 value from UTF-8 text.
// Returns S_OK on success, otherwise an HRESULT error.
COMMON_API HRESULT ParseJsonValue(std::string_view jsonText, JsonValue& out) noexcept;

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
