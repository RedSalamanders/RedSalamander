#include <algorithm>
#include <array>
#include <charconv>
#include <cmath>
#include <cstdint>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <mutex>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <strsafe.h>
#include <system_error>
#include <unordered_set>
#include <utility>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <KnownFolders.h>
#include <pathcch.h>
#include <shlobj_core.h>
#pragma comment(lib, "pathcch.lib")

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/filesystem.h>
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.'cur_key->next' contains the same NULL value as 'removed_item' did..
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include "SettingsStore.h"

#include "Helpers.h"
#include "Version.h"

namespace
{
constexpr wchar_t kCompanyDirectoryName[]  = L"RedSalamander";
constexpr wchar_t kSettingsDirectoryName[] = L"Settings";

constexpr size_t kMaxSettingsFileBytes = 16u * 1024u * 1024u; // 16 MiB safety limit

constexpr wchar_t kSettingsStoreSchemaFileName[] = L"SettingsStore.schema.json";

[[nodiscard]] std::filesystem::path GetShippedSettingsStoreSchemaPath() noexcept
{
    wil::unique_cotaskmem_string modulePath = wil::GetModuleFileNameW();
    if (! modulePath)
    {
        return {};
    }

    std::filesystem::path exePath(modulePath.get());
    if (exePath.empty())
    {
        return {};
    }

    return exePath.parent_path() / kSettingsStoreSchemaFileName;
}

void StripUtf8BomInPlace(std::string& text) noexcept
{
    if (text.size() < 3u)
    {
        return;
    }

    if (static_cast<unsigned char>(text[0]) == 0xEFu && static_cast<unsigned char>(text[1]) == 0xBBu && static_cast<unsigned char>(text[2]) == 0xBFu)
    {
        text.erase(0, 3);
    }
}

// Settings store schema is shipped as `SettingsStore.schema.json` next to the executable.
// Source of truth: `Specs/SettingsStore.schema.json`.

std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return result;
}

std::wstring_view DescribeJsonReadError(yyjson_read_code code) noexcept
{
    switch (code)
    {
        case YYJSON_READ_ERROR_INVALID_PARAMETER: return L"Invalid parameter (NULL input or zero length)";
        case YYJSON_READ_ERROR_MEMORY_ALLOCATION: return L"Memory allocation failed";
        case YYJSON_READ_ERROR_EMPTY_CONTENT: return L"Input JSON string is empty";
        case YYJSON_READ_ERROR_UNEXPECTED_CONTENT: return L"Unexpected content after document";
        case YYJSON_READ_ERROR_UNEXPECTED_END: return L"Unexpected end of input";
        case YYJSON_READ_ERROR_UNEXPECTED_CHARACTER: return L"Unexpected character inside document";
        case YYJSON_READ_ERROR_JSON_STRUCTURE: return L"Invalid JSON structure";
        case YYJSON_READ_ERROR_INVALID_COMMENT: return L"Invalid comment";
        case YYJSON_READ_ERROR_INVALID_NUMBER: return L"Invalid number";
        case YYJSON_READ_ERROR_INVALID_STRING: return L"Invalid string";
        case YYJSON_READ_ERROR_LITERAL: return L"Invalid JSON literal";
        case YYJSON_READ_ERROR_FILE_OPEN: return L"Failed to open file";
        case YYJSON_READ_ERROR_FILE_READ: return L"Failed to read file";
        case YYJSON_READ_ERROR_MORE: return L"Incomplete input (more data expected)";
        default: return L"Unknown parse error";
    }
}

void LogJsonParseError(const wchar_t* context, const std::filesystem::path& path, const yyjson_read_err& err) noexcept
{
    const std::wstring message   = (err.msg && err.msg[0] != '\0') ? Utf16FromUtf8(err.msg) : std::wstring{};
    const std::wstring_view desc = DescribeJsonReadError(static_cast<yyjson_read_code>(err.code));

    std::wstring details;
    if (! desc.empty())
    {
        details.assign(desc);
    }
    if (! message.empty())
    {
        if (! details.empty())
        {
            details.append(L"; ");
        }
        details.append(message);
    }

    if (details.empty())
    {
        Debug::Error(L"Failed to parse {} '{}' (code {}, pos {})", context, path.c_str(), err.code, err.pos);
    }
    else
    {
        Debug::Error(L"Failed to parse {} '{}' (code {}, pos {}): '{}'", context, path.c_str(), err.code, err.pos, details.c_str());
    }
}

bool TryHexNibble(char c, uint8_t& out) noexcept
{
    if (c >= '0' && c <= '9')
    {
        out = static_cast<uint8_t>(c - '0');
        return true;
    }
    if (c >= 'a' && c <= 'f')
    {
        out = static_cast<uint8_t>(10 + (c - 'a'));
        return true;
    }
    if (c >= 'A' && c <= 'F')
    {
        out = static_cast<uint8_t>(10 + (c - 'A'));
        return true;
    }
    return false;
}

bool TryHexByte(std::string_view text, size_t offset, uint8_t& out) noexcept
{
    if (offset + 2 > text.size())
    {
        return false;
    }

    uint8_t hi = 0;
    uint8_t lo = 0;
    if (! TryHexNibble(text[offset], hi))
    {
        return false;
    }
    if (! TryHexNibble(text[offset + 1], lo))
    {
        return false;
    }
    out = static_cast<uint8_t>((hi << 4) | lo);
    return true;
}

bool TryParseColorUtf8(std::string_view text, uint32_t& argb) noexcept
{
    if (text.size() != 7 && text.size() != 9)
    {
        return false;
    }
    if (text[0] != '#')
    {
        return false;
    }

    uint8_t a  = 0xFF;
    size_t pos = 1;
    if (text.size() == 9)
    {
        if (! TryHexByte(text, pos, a))
        {
            return false;
        }
        pos += 2;
    }

    uint8_t r = 0;
    uint8_t g = 0;
    uint8_t b = 0;
    if (! TryHexByte(text, pos, r))
    {
        return false;
    }
    pos += 2;
    if (! TryHexByte(text, pos, g))
    {
        return false;
    }
    pos += 2;
    if (! TryHexByte(text, pos, b))
    {
        return false;
    }

    argb = (static_cast<uint32_t>(a) << 24) | (static_cast<uint32_t>(r) << 16) | (static_cast<uint32_t>(g) << 8) | static_cast<uint32_t>(b);
    return true;
}

std::wstring MakeUtcTimestamp() noexcept
{
    SYSTEMTIME st{};
    GetSystemTime(&st);
    return std::format(L"{:04}{:02}{:02}T{:02}{:02}{:02}Z", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
}

std::filesystem::path MakeBackupPath(const std::filesystem::path& settingsPath) noexcept
{
    const std::wstring baseName = settingsPath.filename().wstring();
    const std::wstring stamp    = MakeUtcTimestamp();
    std::wstring backupName     = std::format(L"{}.bad.{}", baseName, stamp);

    std::filesystem::path candidate = settingsPath.parent_path() / backupName;
    for (int i = 1; std::filesystem::exists(candidate) && i < 100; ++i)
    {
        backupName = std::format(L"{}.bad.{}.{}", baseName, stamp, i);
        candidate  = settingsPath.parent_path() / backupName;
    }
    return candidate;
}

void BackupBadSettingsFile(const std::filesystem::path& path) noexcept
{
    const std::filesystem::path backup = MakeBackupPath(path);
    BOOL res                           = MoveFileExW(path.c_str(), backup.c_str(), MOVEFILE_COPY_ALLOWED);
    if (! res)
    {
        Debug::ErrorWithLastError(L"Failed to back up bad settings file from '{}' to '{}'", path.c_str(), backup.c_str());
    }
}

HRESULT ReadFileBytes(const std::filesystem::path& path, std::string& out) noexcept
{
    out.clear();

    wil::unique_handle file(CreateFileW(
        path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        auto lastError = Debug::ErrorWithLastError(L"Failed to open file '{}'", path.c_str());
        return HRESULT_FROM_WIN32(lastError);
    }

    LARGE_INTEGER size{};
    if (! GetFileSizeEx(file.get(), &size))
    {
        auto lastError = Debug::ErrorWithLastError(L"Failed to get size of file '{}'", path.c_str());
        return HRESULT_FROM_WIN32(lastError);
    }

    if (size.QuadPart < 0 || static_cast<uint64_t>(size.QuadPart) > kMaxSettingsFileBytes)
    {
        Debug::Error(L"File '{}' has invalid size {}", path.c_str(), size.QuadPart);
        return HRESULT_FROM_WIN32(ERROR_FILE_INVALID);
    }

    const size_t length = static_cast<size_t>(size.QuadPart);
    out.resize(length);

    size_t totalRead = 0;
    while (totalRead < length)
    {
        DWORD chunkRead    = 0;
        const DWORD toRead = static_cast<DWORD>(std::min<size_t>(length - totalRead, static_cast<size_t>(std::numeric_limits<DWORD>::max())));
        if (! ReadFile(file.get(), out.data() + totalRead, toRead, &chunkRead, nullptr))
        {
            auto lastError = Debug::ErrorWithLastError(L"Failed to read file '{}'", path.c_str());
            return HRESULT_FROM_WIN32(lastError);
        }
        if (chunkRead == 0)
        {
            break;
        }
        totalRead += static_cast<size_t>(chunkRead);
    }

    if (totalRead != length)
    {
        Debug::Error(L"Failed to read complete settings file '{}'", path.c_str());
        return HRESULT_FROM_WIN32(ERROR_READ_FAULT);
    }

    return S_OK;
}

HRESULT WriteFileBytesAtomic(const std::filesystem::path& path, std::string_view bytes) noexcept
{
    const std::filesystem::path directory = path.parent_path();
    if (directory.empty())
    {
        return E_INVALIDARG;
    }

    HRESULT hr = wil::CreateDirectoryDeepNoThrow(directory.c_str());
    if (FAILED(hr))
    {
        return hr;
    }

    const DWORD dirAttrs = GetFileAttributesW(directory.c_str());
    if (dirAttrs == INVALID_FILE_ATTRIBUTES)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    if ((dirAttrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }

    std::filesystem::path tmpPath = path;
    tmpPath += L".tmp";

    wil::unique_handle file(CreateFileW(tmpPath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    size_t totalWritten = 0;
    while (totalWritten < bytes.size())
    {
        DWORD chunkWritten  = 0;
        const DWORD toWrite = static_cast<DWORD>(std::min<size_t>(bytes.size() - totalWritten, static_cast<size_t>(std::numeric_limits<DWORD>::max())));
        if (! WriteFile(file.get(), bytes.data() + totalWritten, toWrite, &chunkWritten, nullptr))
        {
            auto lastError = Debug::ErrorWithLastError(L"Failed to write settings file '{}'", tmpPath.c_str());
            file.reset();
            DeleteFileW(tmpPath.c_str());
            return HRESULT_FROM_WIN32(lastError);
        }
        totalWritten += static_cast<size_t>(chunkWritten);
    }

    if (! FlushFileBuffers(file.get()))
    {
        auto lastError = Debug::ErrorWithLastError(L"Failed to flush settings file '{}'", tmpPath.c_str());
        file.reset();
        DeleteFileW(tmpPath.c_str());
        return HRESULT_FROM_WIN32(lastError);
    }

    file.reset();

    if (! MoveFileExW(tmpPath.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
    {
        auto lastError = Debug::ErrorWithLastError(L"Failed to replace settings file '{}' with temporary file '{}'", path.c_str(), tmpPath.c_str());
        DeleteFileW(tmpPath.c_str());
        return HRESULT_FROM_WIN32(lastError);
    }

    return S_OK;
}

bool GetBool(yyjson_val* obj, const char* key, bool& out) noexcept
{
    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v || ! yyjson_is_bool(v))
    {
        Debug::Error(L"Expected boolean value for key '{}'", Utf16FromUtf8(key).c_str());
        return false;
    }
    out = yyjson_get_bool(v);
    return true;
}

bool GetUInt32(yyjson_val* obj, const char* key, uint32_t& out) noexcept
{
    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v || ! yyjson_is_uint(v))
    {
        Debug::Error(L"Expected unsigned integer value for key '{}'", Utf16FromUtf8(key).c_str());
        return false;
    }
    out = static_cast<uint32_t>(yyjson_get_uint(v));
    return true;
}

bool IsAsciiSpace(char ch) noexcept
{
    return ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r' || ch == '\f' || ch == '\v';
}

std::string_view TrimAscii(std::string_view text) noexcept
{
    while (! text.empty() && IsAsciiSpace(text.front()))
    {
        text.remove_prefix(1);
    }

    while (! text.empty() && IsAsciiSpace(text.back()))
    {
        text.remove_suffix(1);
    }

    return text;
}

char FoldAsciiCase(char ch) noexcept
{
    if (ch >= 'A' && ch <= 'Z')
    {
        return static_cast<char>(ch - 'A' + 'a');
    }
    return ch;
}

bool EqualsIgnoreAsciiCase(std::string_view a, std::string_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        if (FoldAsciiCase(a[i]) != FoldAsciiCase(b[i]))
        {
            return false;
        }
    }

    return true;
}

std::string VkToStableName(uint32_t vk)
{
    const uint32_t clampedVk = vk & 0xFFu;

    if (clampedVk >= static_cast<uint32_t>(VK_F1) && clampedVk <= static_cast<uint32_t>(VK_F24))
    {
        const uint32_t number = clampedVk - static_cast<uint32_t>(VK_F1) + 1u;
        return std::format("F{}", number);
    }

    if ((clampedVk >= static_cast<uint32_t>('0') && clampedVk <= static_cast<uint32_t>('9')) ||
        (clampedVk >= static_cast<uint32_t>('A') && clampedVk <= static_cast<uint32_t>('Z')))
    {
        char buf[2]{};
        buf[0] = static_cast<char>(clampedVk);
        buf[1] = '\0';
        return std::string(buf);
    }

    switch (clampedVk)
    {
        case VK_BACK: return "Backspace";
        case VK_TAB: return "Tab";
        case VK_RETURN: return "Enter";
        case VK_SPACE: return "Space";
        case VK_PRIOR: return "PageUp";
        case VK_NEXT: return "PageDown";
        case VK_END: return "End";
        case VK_HOME: return "Home";
        case VK_LEFT: return "Left";
        case VK_UP: return "Up";
        case VK_RIGHT: return "Right";
        case VK_DOWN: return "Down";
        case VK_INSERT: return "Insert";
        case VK_DELETE: return "Delete";
        case VK_ESCAPE: return "Escape";
    }

    return std::format("VK_{:02X}", clampedVk);
}

bool TryParseVkFromText(std::string_view text, uint32_t& outVk) noexcept
{
    text = TrimAscii(text);
    if (text.empty())
    {
        return false;
    }

    if (text.size() == 1)
    {
        char ch = text[0];
        if (ch >= 'a' && ch <= 'z')
        {
            ch = static_cast<char>(ch - 'a' + 'A');
        }
        if ((ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'Z'))
        {
            outVk = static_cast<uint32_t>(static_cast<unsigned char>(ch));
            return true;
        }
    }

    if (text.size() >= 2 && (text[0] == 'F' || text[0] == 'f'))
    {
        const std::string_view numberText = text.substr(1);
        uint32_t number                   = 0;
        const auto [ptr, ec]              = std::from_chars(numberText.data(), numberText.data() + numberText.size(), number);
        if (ec == std::errc{} && ptr == numberText.data() + numberText.size() && number >= 1u && number <= 24u)
        {
            outVk = static_cast<uint32_t>(VK_F1) + (number - 1u);
            return true;
        }
    }

    if (text.size() == 5 && (text[0] == 'V' || text[0] == 'v') && (text[1] == 'K' || text[1] == 'k') && text[2] == '_')
    {
        const std::string_view hexText = text.substr(3, 2);
        uint32_t vk                    = 0;
        const auto [ptr, ec]           = std::from_chars(hexText.data(), hexText.data() + hexText.size(), vk, 16);
        if (ec == std::errc{} && ptr == hexText.data() + hexText.size() && vk <= 0xFFu)
        {
            outVk = vk;
            return true;
        }
    }

    struct NamedVk
    {
        std::string_view name;
        uint32_t vk = 0;
    };

    constexpr std::array<NamedVk, 16> kNamedVks = {
        NamedVk{"Backspace", static_cast<uint32_t>(VK_BACK)},
        NamedVk{"Tab", static_cast<uint32_t>(VK_TAB)},
        NamedVk{"Enter", static_cast<uint32_t>(VK_RETURN)},
        NamedVk{"Return", static_cast<uint32_t>(VK_RETURN)},
        NamedVk{"Space", static_cast<uint32_t>(VK_SPACE)},
        NamedVk{"PageUp", static_cast<uint32_t>(VK_PRIOR)},
        NamedVk{"PageDown", static_cast<uint32_t>(VK_NEXT)},
        NamedVk{"End", static_cast<uint32_t>(VK_END)},
        NamedVk{"Home", static_cast<uint32_t>(VK_HOME)},
        NamedVk{"Left", static_cast<uint32_t>(VK_LEFT)},
        NamedVk{"Up", static_cast<uint32_t>(VK_UP)},
        NamedVk{"Right", static_cast<uint32_t>(VK_RIGHT)},
        NamedVk{"Down", static_cast<uint32_t>(VK_DOWN)},
        NamedVk{"Insert", static_cast<uint32_t>(VK_INSERT)},
        NamedVk{"Delete", static_cast<uint32_t>(VK_DELETE)},
        NamedVk{"Escape", static_cast<uint32_t>(VK_ESCAPE)},
    };

    for (const auto& item : kNamedVks)
    {
        if (EqualsIgnoreAsciiCase(text, item.name))
        {
            outVk = item.vk;
            return true;
        }
    }

    return false;
}

uint64_t MultiplyOrSaturate(uint64_t a, uint64_t b) noexcept
{
    if (a == 0 || b == 0)
    {
        return 0;
    }

    if (a > (std::numeric_limits<uint64_t>::max() / b))
    {
        return std::numeric_limits<uint64_t>::max();
    }

    return a * b;
}

bool TryParseByteSizeText(std::string_view text, uint64_t& outBytes) noexcept
{
    constexpr uint64_t kKiB = 1024ull;
    constexpr uint64_t kMiB = 1024ull * 1024ull;
    constexpr uint64_t kGiB = 1024ull * 1024ull * 1024ull;

    text = TrimAscii(text);
    if (text.empty())
    {
        return false;
    }

    uint64_t number   = 0;
    const char* begin = text.data();
    const char* end   = begin + text.size();

    const auto [ptr, ec] = std::from_chars(begin, end, number);
    if (ec != std::errc{})
    {
        return false;
    }

    std::string_view unit(ptr, static_cast<size_t>(end - ptr));
    unit = TrimAscii(unit);

    uint64_t multiplier = 0;
    if (unit.empty() || EqualsIgnoreAsciiCase(unit, "kb"))
    {
        // Bare numeric strings are interpreted as KiB for user-friendliness.
        multiplier = kKiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, "mb"))
    {
        multiplier = kMiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, "gb"))
    {
        multiplier = kGiB;
    }
    else
    {
        return false;
    }

    outBytes = MultiplyOrSaturate(number, multiplier);
    return true;
}

bool GetDirectoryCacheBytes(yyjson_val* obj, const char* key, uint64_t& outBytes) noexcept
{
    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v)
    {
        return false;
    }

    constexpr uint64_t kKiB = 1024ull;

    if (yyjson_is_uint(v))
    {
        const uint64_t kiloBytes = yyjson_get_uint(v);
        outBytes                 = MultiplyOrSaturate(kiloBytes, kKiB);
        return true;
    }

    if (yyjson_is_sint(v))
    {
        const int64_t kiloBytesSigned = yyjson_get_sint(v);
        if (kiloBytesSigned < 0)
        {
            return false;
        }
        outBytes = MultiplyOrSaturate(static_cast<uint64_t>(kiloBytesSigned), kKiB);
        return true;
    }

    if (yyjson_is_str(v))
    {
        const char* s = yyjson_get_str(v);
        if (! s)
        {
            return false;
        }

        uint64_t parsed = 0;
        if (! TryParseByteSizeText(std::string_view(s), parsed))
        {
            return false;
        }

        outBytes = parsed;
        return true;
    }

    return false;
}

bool GetDouble(yyjson_val* obj, const char* key, double& out) noexcept
{
    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v)
    {
        return false;
    }

    if (yyjson_is_real(v))
    {
        out = yyjson_get_real(v);
        return true;
    }
    if (yyjson_is_sint(v))
    {
        out = static_cast<double>(yyjson_get_sint(v));
        return true;
    }
    if (yyjson_is_uint(v))
    {
        out = static_cast<double>(yyjson_get_uint(v));
        return true;
    }

    return false;
}

yyjson_val* GetObj(yyjson_val* obj, const char* key) noexcept
{
    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v || ! yyjson_is_obj(v))
    {
        return nullptr;
    }
    return v;
}

yyjson_val* GetArr(yyjson_val* obj, const char* key) noexcept
{
    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v || ! yyjson_is_arr(v))
    {
        return nullptr;
    }
    return v;
}

std::optional<std::string_view> GetString(yyjson_val* obj, const char* key) noexcept
{
    yyjson_val* v = yyjson_obj_get(obj, key);
    if (! v || ! yyjson_is_str(v))
    {
        return std::nullopt;
    }
    const char* s = yyjson_get_str(v);
    if (! s)
    {
        return std::nullopt;
    }
    return std::string_view(s);
}

HRESULT ConvertYyjsonToJsonValue(yyjson_val* val, Common::Settings::JsonValue& out) noexcept
{
    if (! val || yyjson_is_null(val))
    {
        out.value = std::monostate{};
        return S_OK;
    }

    {
        if (yyjson_is_bool(val))
        {
            out.value = yyjson_get_bool(val) != 0;
            return S_OK;
        }

        if (yyjson_is_sint(val))
        {
            out.value = yyjson_get_sint(val);
            return S_OK;
        }

        if (yyjson_is_uint(val))
        {
            out.value = yyjson_get_uint(val);
            return S_OK;
        }

        if (yyjson_is_real(val))
        {
            out.value = yyjson_get_real(val);
            return S_OK;
        }

        if (yyjson_is_str(val))
        {
            const char* s = yyjson_get_str(val);
            if (! s)
            {
                out.value = std::string{};
                return S_OK;
            }
            out.value = std::string(s, yyjson_get_len(val));
            return S_OK;
        }

        if (yyjson_is_arr(val))
        {
            auto arr = std::make_shared<Common::Settings::JsonArray>();

            const size_t count = yyjson_arr_size(val);
            arr->items.reserve(count);

            for (size_t i = 0; i < count; ++i)
            {
                yyjson_val* item = yyjson_arr_get(val, i);
                Common::Settings::JsonValue converted;
                const HRESULT hr = ConvertYyjsonToJsonValue(item, converted);
                if (FAILED(hr))
                {
                    return hr;
                }
                arr->items.push_back(std::move(converted));
            }

            out.value = std::move(arr);
            return S_OK;
        }

        if (yyjson_is_obj(val))
        {
            auto obj = std::make_shared<Common::Settings::JsonObject>();

            const size_t count = yyjson_obj_size(val);
            obj->members.reserve(count);

            yyjson_val* key      = nullptr;
            yyjson_obj_iter iter = yyjson_obj_iter_with(val);
            while ((key = yyjson_obj_iter_next(&iter)))
            {
                yyjson_val* memberVal = yyjson_obj_iter_get_val(key);
                if (! key || ! yyjson_is_str(key) || ! memberVal)
                {
                    continue;
                }

                const char* keyStr = yyjson_get_str(key);
                if (! keyStr)
                {
                    continue;
                }
                std::string keyText(keyStr, yyjson_get_len(key));

                Common::Settings::JsonValue converted;
                const HRESULT hr = ConvertYyjsonToJsonValue(memberVal, converted);
                if (FAILED(hr))
                {
                    return hr;
                }

                obj->members.emplace_back(std::move(keyText), std::move(converted));
            }

            out.value = std::move(obj);
            return S_OK;
        }

        out.value = std::monostate{};
        return S_OK;
    }
}

yyjson_mut_val* NewYyjsonFromJsonValue(yyjson_mut_doc* doc, const Common::Settings::JsonValue& value, HRESULT& outHr) noexcept
{
    outHr = S_OK;

    if (! doc)
    {
        outHr = E_POINTER;
        return nullptr;
    }

    {
        if (std::holds_alternative<std::monostate>(value.value))
        {
            yyjson_mut_val* v = yyjson_mut_null(doc);
            if (! v)
            {
                outHr = E_OUTOFMEMORY;
            }
            return v;
        }

        if (const bool* b = std::get_if<bool>(&value.value))
        {
            yyjson_mut_val* v = yyjson_mut_bool(doc, *b ? 1 : 0);
            if (! v)
            {
                outHr = E_OUTOFMEMORY;
            }
            return v;
        }

        if (const int64_t* i = std::get_if<int64_t>(&value.value))
        {
            yyjson_mut_val* v = yyjson_mut_sint(doc, *i);
            if (! v)
            {
                outHr = E_OUTOFMEMORY;
            }
            return v;
        }

        if (const uint64_t* u = std::get_if<uint64_t>(&value.value))
        {
            yyjson_mut_val* v = yyjson_mut_uint(doc, *u);
            if (! v)
            {
                outHr = E_OUTOFMEMORY;
            }
            return v;
        }

        if (const double* d = std::get_if<double>(&value.value))
        {
            yyjson_mut_val* v = yyjson_mut_real(doc, *d);
            if (! v)
            {
                outHr = E_OUTOFMEMORY;
            }
            return v;
        }

        if (const std::string* s = std::get_if<std::string>(&value.value))
        {
            yyjson_mut_val* v = yyjson_mut_strncpy(doc, s->c_str(), s->size());
            if (! v)
            {
                outHr = E_OUTOFMEMORY;
            }
            return v;
        }

        if (const Common::Settings::JsonValue::ArrayPtr* arrPtr = std::get_if<Common::Settings::JsonValue::ArrayPtr>(&value.value))
        {
            yyjson_mut_val* arr = yyjson_mut_arr(doc);
            if (! arr)
            {
                outHr = E_OUTOFMEMORY;
                return nullptr;
            }

            if (*arrPtr)
            {
                for (const auto& item : (*arrPtr)->items)
                {
                    HRESULT itemHr        = S_OK;
                    yyjson_mut_val* entry = NewYyjsonFromJsonValue(doc, item, itemHr);
                    if (! entry)
                    {
                        outHr = itemHr;
                        return nullptr;
                    }
                    yyjson_mut_arr_add_val(arr, entry);
                }
            }

            return arr;
        }

        if (const Common::Settings::JsonValue::ObjectPtr* objPtr = std::get_if<Common::Settings::JsonValue::ObjectPtr>(&value.value))
        {
            yyjson_mut_val* obj = yyjson_mut_obj(doc);
            if (! obj)
            {
                outHr = E_OUTOFMEMORY;
                return nullptr;
            }

            if (*objPtr)
            {
                for (const auto& [k, v] : (*objPtr)->members)
                {
                    yyjson_mut_val* key = yyjson_mut_strncpy(doc, k.c_str(), k.size());
                    if (! key)
                    {
                        outHr = E_OUTOFMEMORY;
                        return nullptr;
                    }

                    HRESULT valHr       = S_OK;
                    yyjson_mut_val* val = NewYyjsonFromJsonValue(doc, v, valHr);
                    if (! val)
                    {
                        outHr = valHr;
                        return nullptr;
                    }

                    yyjson_mut_obj_add(obj, key, val);
                }
            }

            return obj;
        }

        outHr = E_UNEXPECTED;
        return nullptr;
    }
}

void ParseWindows(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* windows = yyjson_obj_get(root, "windows");
    if (! windows || ! yyjson_is_obj(windows))
    {
        return;
    }

    yyjson_val* key      = nullptr;
    yyjson_val* val      = nullptr;
    yyjson_obj_iter iter = yyjson_obj_iter_with(windows);
    while ((key = yyjson_obj_iter_next(&iter)))
    {
        val = yyjson_obj_iter_get_val(key);
        if (! val || ! yyjson_is_obj(val) || ! yyjson_is_str(key))
        {
            continue;
        }

        const char* keyStr = yyjson_get_str(key);
        if (! keyStr)
        {
            continue;
        }

        Common::Settings::WindowPlacement placement;

        const auto stateText = GetString(val, "state");
        if (stateText)
        {
            if (*stateText == "maximized")
            {
                placement.state = Common::Settings::WindowState::Maximized;
            }
        }

        yyjson_val* bounds = GetObj(val, "bounds");
        if (bounds)
        {
            yyjson_val* vx = yyjson_obj_get(bounds, "x");
            yyjson_val* vy = yyjson_obj_get(bounds, "y");
            yyjson_val* vw = yyjson_obj_get(bounds, "width");
            yyjson_val* vh = yyjson_obj_get(bounds, "height");
            if (vx && vy && vw && vh && yyjson_is_int(vx) && yyjson_is_int(vy) && yyjson_is_int(vw) && yyjson_is_int(vh))
            {
                placement.bounds.x      = static_cast<int>(yyjson_get_int(vx));
                placement.bounds.y      = static_cast<int>(yyjson_get_int(vy));
                placement.bounds.width  = static_cast<int>(yyjson_get_int(vw));
                placement.bounds.height = static_cast<int>(yyjson_get_int(vh));
            }
        }

        uint32_t dpiValue = 0;
        if (GetUInt32(val, "dpi", dpiValue) && dpiValue > 0)
        {
            placement.dpi = dpiValue;
        }

        const std::wstring id = Utf16FromUtf8(keyStr);
        if (! id.empty())
        {
            out.windows[id] = std::move(placement);
        }
    }
}

void ParseTheme(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* theme = yyjson_obj_get(root, "theme");
    if (! theme || ! yyjson_is_obj(theme))
    {
        return;
    }

    const auto currentId = GetString(theme, "currentThemeId");
    if (currentId)
    {
        const std::wstring currentWide = Utf16FromUtf8(*currentId);
        if (! currentWide.empty())
        {
            out.theme.currentThemeId = currentWide;
        }
    }

    yyjson_val* themes = GetArr(theme, "themes");
    if (! themes)
    {
        return;
    }

    const size_t count = yyjson_arr_size(themes);
    out.theme.themes.clear();
    out.theme.themes.reserve(count);

    for (size_t i = 0; i < count; ++i)
    {
        yyjson_val* item = yyjson_arr_get(themes, i);
        if (! item || ! yyjson_is_obj(item))
        {
            continue;
        }

        const auto idText   = GetString(item, "id");
        const auto nameText = GetString(item, "name");
        const auto baseText = GetString(item, "baseThemeId");
        yyjson_val* colors  = GetObj(item, "colors");
        if (! idText || ! nameText || ! baseText || ! colors)
        {
            continue;
        }

        Common::Settings::ThemeDefinition def;
        def.id          = Utf16FromUtf8(*idText);
        def.name        = Utf16FromUtf8(*nameText);
        def.baseThemeId = Utf16FromUtf8(*baseText);
        if (def.id.empty() || def.name.empty() || def.baseThemeId.empty())
        {
            continue;
        }

        yyjson_val* colorKey = nullptr;
        yyjson_val* colorVal = nullptr;
        yyjson_obj_iter iter = yyjson_obj_iter_with(colors);
        while ((colorKey = yyjson_obj_iter_next(&iter)))
        {
            colorVal = yyjson_obj_iter_get_val(colorKey);
            if (! colorVal || ! yyjson_is_str(colorKey) || ! yyjson_is_str(colorVal))
            {
                continue;
            }

            const char* keyStr = yyjson_get_str(colorKey);
            const char* valStr = yyjson_get_str(colorVal);
            if (! keyStr || ! valStr)
            {
                continue;
            }

            uint32_t argb = 0;
            if (! TryParseColorUtf8(valStr, argb))
            {
                continue;
            }

            const std::wstring keyWide = Utf16FromUtf8(keyStr);
            if (keyWide.empty())
            {
                continue;
            }

            def.colors[keyWide] = argb;
        }

        out.theme.themes.push_back(std::move(def));
    }
}

void ParsePlugins(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* plugins = yyjson_obj_get(root, "plugins");
    if (! plugins || ! yyjson_is_obj(plugins))
    {
        return;
    }

    const auto currentId = GetString(plugins, "currentFileSystemPluginId");
    if (currentId)
    {
        const std::wstring currentWide = Utf16FromUtf8(*currentId);
        if (! currentWide.empty())
        {
            out.plugins.currentFileSystemPluginId = currentWide;
        }
    }

    yyjson_val* disabled = GetArr(plugins, "disabledPluginIds");
    if (disabled)
    {
        out.plugins.disabledPluginIds.clear();

        const size_t count = yyjson_arr_size(disabled);
        out.plugins.disabledPluginIds.reserve(count);

        for (size_t i = 0; i < count; ++i)
        {
            yyjson_val* v = yyjson_arr_get(disabled, i);
            if (! v || ! yyjson_is_str(v))
            {
                continue;
            }
            const char* s = yyjson_get_str(v);
            if (! s)
            {
                continue;
            }
            std::wstring id = Utf16FromUtf8(s);
            if (id.empty())
            {
                continue;
            }
            out.plugins.disabledPluginIds.push_back(std::move(id));
        }
    }

    yyjson_val* custom = GetArr(plugins, "customPluginPaths");
    if (custom)
    {
        out.plugins.customPluginPaths.clear();

        const size_t count = yyjson_arr_size(custom);
        out.plugins.customPluginPaths.reserve(count);

        for (size_t i = 0; i < count; ++i)
        {
            yyjson_val* v = yyjson_arr_get(custom, i);
            if (! v || ! yyjson_is_str(v))
            {
                continue;
            }
            const char* s = yyjson_get_str(v);
            if (! s)
            {
                continue;
            }
            std::wstring pathText = Utf16FromUtf8(s);
            if (pathText.empty())
            {
                continue;
            }
            out.plugins.customPluginPaths.emplace_back(std::filesystem::path(std::move(pathText)));
        }
    }

    yyjson_val* configs = GetObj(plugins, "configurationByPluginId");
    if (configs)
    {
        out.plugins.configurationByPluginId.clear();

        yyjson_val* key      = nullptr;
        yyjson_val* val      = nullptr;
        yyjson_obj_iter iter = yyjson_obj_iter_with(configs);
        while ((key = yyjson_obj_iter_next(&iter)))
        {
            val = yyjson_obj_iter_get_val(key);
            if (! key || ! yyjson_is_str(key) || ! val)
            {
                continue;
            }

            const char* keyStr = yyjson_get_str(key);
            if (! keyStr)
            {
                continue;
            }

            std::wstring id = Utf16FromUtf8(keyStr);
            if (id.empty())
            {
                continue;
            }

            Common::Settings::JsonValue config;
            if (yyjson_is_str(val))
            {
                const char* valStr = yyjson_get_str(val);
                if (! valStr)
                {
                    continue;
                }

                std::string configText(valStr, yyjson_get_len(val));

                yyjson_read_err configErr{};
                yyjson_doc* configDoc = yyjson_read_opts(configText.data(), configText.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &configErr);
                if (! configDoc)
                {
                    continue;
                }

                auto freeConfigDoc = wil::scope_exit([&] { yyjson_doc_free(configDoc); });

                yyjson_val* configRoot = yyjson_doc_get_root(configDoc);
                if (! configRoot)
                {
                    continue;
                }

                const HRESULT cfgHr = ConvertYyjsonToJsonValue(configRoot, config);
                if (FAILED(cfgHr))
                {
                    continue;
                }
            }
            else
            {
                const HRESULT cfgHr = ConvertYyjsonToJsonValue(val, config);
                if (FAILED(cfgHr))
                {
                    continue;
                }
            }

            out.plugins.configurationByPluginId.emplace(std::move(id), std::move(config));
        }
    }

    const auto migratePluginId = [](std::wstring& id)
    {
        if (id == L"builtin/filesystem" || id == L"file")
        {
            id = L"builtin/file-system";
            return;
        }

        if (id == L"optional/filesystemDummy" || id == L"fk")
        {
            id = L"builtin/file-system-dummy";
        }
    };

    migratePluginId(out.plugins.currentFileSystemPluginId);

    for (auto& id : out.plugins.disabledPluginIds)
    {
        migratePluginId(id);
    }

    if (! out.plugins.configurationByPluginId.empty())
    {
        std::unordered_map<std::wstring, Common::Settings::JsonValue> migrated;
        migrated.reserve(out.plugins.configurationByPluginId.size());

        for (auto& [id, config] : out.plugins.configurationByPluginId)
        {
            std::wstring newId = id;
            migratePluginId(newId);
            migrated.emplace(std::move(newId), std::move(config));
        }

        out.plugins.configurationByPluginId = std::move(migrated);
    }
}

void ParseExtensions(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* extensions = yyjson_obj_get(root, "extensions");
    if (! extensions || ! yyjson_is_obj(extensions))
    {
        return;
    }

    auto parseExtMap = [&](const char* field, std::unordered_map<std::wstring, std::wstring>& target)
    {
        yyjson_val* openWith = yyjson_obj_get(extensions, field);
        if (! openWith || ! yyjson_is_obj(openWith))
        {
            return;
        }

        target.clear();

        yyjson_val* key      = nullptr;
        yyjson_val* val      = nullptr;
        yyjson_obj_iter iter = yyjson_obj_iter_with(openWith);
        while ((key = yyjson_obj_iter_next(&iter)))
        {
            val = yyjson_obj_iter_get_val(key);
            if (! key || ! yyjson_is_str(key) || ! val || ! yyjson_is_str(val))
            {
                continue;
            }

            const char* keyStr = yyjson_get_str(key);
            if (! keyStr)
            {
                continue;
            }

            const char* valStr = yyjson_get_str(val);
            if (! valStr)
            {
                continue;
            }

            std::wstring ext = Utf16FromUtf8(keyStr);
            if (ext.empty())
            {
                continue;
            }

            if (ext.front() != L'.')
            {
                ext.insert(ext.begin(), L'.');
            }

            std::transform(ext.begin(), ext.end(), ext.begin(), [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))); });

            std::wstring pluginId = Utf16FromUtf8(valStr);
            target[ext]           = std::move(pluginId);
        }
    };

    parseExtMap("openWithFileSystemByExtension", out.extensions.openWithFileSystemByExtension);
    parseExtMap("openWithViewerByExtension", out.extensions.openWithViewerByExtension);
}

void NormalizeHistory(std::vector<std::filesystem::path>& history, size_t maxItems)
{
    std::vector<std::filesystem::path> normalized;
    normalized.reserve(std::min(history.size(), maxItems));

    for (const auto& path : history)
    {
        if (path.empty())
        {
            continue;
        }
        const bool already = std::find(normalized.begin(), normalized.end(), path) != normalized.end();
        if (already)
        {
            continue;
        }
        normalized.push_back(path);
        if (normalized.size() >= maxItems)
        {
            break;
        }
    }

    history = std::move(normalized);
}

Common::Settings::FolderDisplayMode ParseFolderDisplayMode(std::string_view display) noexcept
{
    if (display == "detailed")
    {
        return Common::Settings::FolderDisplayMode::Detailed;
    }
    return Common::Settings::FolderDisplayMode::Brief;
}

const char* FolderDisplayModeToString(Common::Settings::FolderDisplayMode display) noexcept
{
    switch (display)
    {
        case Common::Settings::FolderDisplayMode::Brief: return "brief";
        case Common::Settings::FolderDisplayMode::Detailed: return "detailed";
    }
    return "brief";
}

Common::Settings::ConnectionAuthMode ParseConnectionAuthMode(std::string_view auth) noexcept
{
    if (auth == "anonymous")
    {
        return Common::Settings::ConnectionAuthMode::Anonymous;
    }
    if (auth == "sshKey")
    {
        return Common::Settings::ConnectionAuthMode::SshKey;
    }
    return Common::Settings::ConnectionAuthMode::Password;
}

const char* ConnectionAuthModeToString(Common::Settings::ConnectionAuthMode auth) noexcept
{
    switch (auth)
    {
        case Common::Settings::ConnectionAuthMode::Anonymous: return "anonymous";
        case Common::Settings::ConnectionAuthMode::Password: return "password";
        case Common::Settings::ConnectionAuthMode::SshKey: return "sshKey";
    }
    return "password";
}

Common::Settings::FolderSortBy ParseFolderSortBy(std::string_view sortBy) noexcept
{
    if (sortBy == "none")
    {
        return Common::Settings::FolderSortBy::None;
    }
    if (sortBy == "extension")
    {
        return Common::Settings::FolderSortBy::Extension;
    }
    if (sortBy == "time")
    {
        return Common::Settings::FolderSortBy::Time;
    }
    if (sortBy == "size")
    {
        return Common::Settings::FolderSortBy::Size;
    }
    if (sortBy == "attributes")
    {
        return Common::Settings::FolderSortBy::Attributes;
    }
    return Common::Settings::FolderSortBy::Name;
}

const char* FolderSortByToString(Common::Settings::FolderSortBy sortBy) noexcept
{
    switch (sortBy)
    {
        case Common::Settings::FolderSortBy::Name: return "name";
        case Common::Settings::FolderSortBy::Extension: return "extension";
        case Common::Settings::FolderSortBy::Time: return "time";
        case Common::Settings::FolderSortBy::Size: return "size";
        case Common::Settings::FolderSortBy::Attributes: return "attributes";
        case Common::Settings::FolderSortBy::None: return "none";
    }
    return "name";
}

Common::Settings::FolderSortDirection DefaultFolderSortDirection(Common::Settings::FolderSortBy sortBy) noexcept
{
    switch (sortBy)
    {
        case Common::Settings::FolderSortBy::Time:
        case Common::Settings::FolderSortBy::Size: return Common::Settings::FolderSortDirection::Descending;
        case Common::Settings::FolderSortBy::Name:
        case Common::Settings::FolderSortBy::Extension:
        case Common::Settings::FolderSortBy::Attributes:
        case Common::Settings::FolderSortBy::None: return Common::Settings::FolderSortDirection::Ascending;
    }
    return Common::Settings::FolderSortDirection::Ascending;
}

Common::Settings::FolderSortDirection ParseFolderSortDirection(std::string_view direction) noexcept
{
    if (direction == "descending")
    {
        return Common::Settings::FolderSortDirection::Descending;
    }
    return Common::Settings::FolderSortDirection::Ascending;
}

const char* FolderSortDirectionToString(Common::Settings::FolderSortDirection direction) noexcept
{
    switch (direction)
    {
        case Common::Settings::FolderSortDirection::Ascending: return "ascending";
        case Common::Settings::FolderSortDirection::Descending: return "descending";
    }
    return "ascending";
}

void ParseFolders(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* folders = yyjson_obj_get(root, "folders");
    if (! folders || ! yyjson_is_obj(folders))
    {
        return;
    }

    Common::Settings::FoldersSettings folderSettings;

    const auto activeText = GetString(folders, "active");
    if (activeText)
    {
        folderSettings.active = Utf16FromUtf8(*activeText);
    }

    yyjson_val* layout = GetObj(folders, "layout");
    if (layout)
    {
        double splitRatio = static_cast<double>(folderSettings.layout.splitRatio);
        if (GetDouble(layout, "splitRatio", splitRatio))
        {
            splitRatio                       = std::clamp(splitRatio, 0.0, 1.0);
            folderSettings.layout.splitRatio = static_cast<float>(splitRatio);
        }

        if (const auto zoomedPaneText = GetString(layout, "zoomedPane"))
        {
            std::wstring zoomedPane = Utf16FromUtf8(*zoomedPaneText);
            if (! zoomedPane.empty())
            {
                folderSettings.layout.zoomedPane = std::move(zoomedPane);
            }
        }

        double zoomRestoreSplitRatio = 0.0;
        if (GetDouble(layout, "zoomRestoreSplitRatio", zoomRestoreSplitRatio))
        {
            zoomRestoreSplitRatio                       = std::clamp(zoomRestoreSplitRatio, 0.0, 1.0);
            folderSettings.layout.zoomRestoreSplitRatio = static_cast<float>(zoomRestoreSplitRatio);
        }
    }

    uint32_t historyMax = folderSettings.historyMax;
    GetUInt32(folders, "historyMax", historyMax);
    historyMax                = std::clamp(historyMax, 1u, 50u);
    folderSettings.historyMax = historyMax;

    yyjson_val* historyArr = GetArr(folders, "history");
    if (historyArr)
    {
        const size_t histCount = yyjson_arr_size(historyArr);
        folderSettings.history.reserve(std::min(histCount, static_cast<size_t>(historyMax)));

        for (size_t h = 0; h < histCount && folderSettings.history.size() < static_cast<size_t>(historyMax); ++h)
        {
            yyjson_val* hv = yyjson_arr_get(historyArr, h);
            if (! hv || ! yyjson_is_str(hv))
            {
                continue;
            }

            const char* hvStr = yyjson_get_str(hv);
            if (! hvStr)
            {
                continue;
            }

            const std::wstring hvWide = Utf16FromUtf8(hvStr);
            if (hvWide.empty())
            {
                continue;
            }

            folderSettings.history.push_back(std::filesystem::path(hvWide));
        }

        NormalizeHistory(folderSettings.history, static_cast<size_t>(historyMax));
    }

    yyjson_val* items = GetArr(folders, "items");
    if (! items)
    {
        return;
    }

    const size_t count = yyjson_arr_size(items);
    folderSettings.items.reserve(count);

    for (size_t i = 0; i < count; ++i)
    {
        yyjson_val* item = yyjson_arr_get(items, i);
        if (! item || ! yyjson_is_obj(item))
        {
            continue;
        }

        const auto slotText    = GetString(item, "slot");
        const auto currentText = GetString(item, "current");
        if (! slotText || ! currentText)
        {
            continue;
        }

        Common::Settings::FolderPane pane;
        pane.slot = Utf16FromUtf8(*slotText);
        if (pane.slot.empty())
        {
            continue;
        }

        const std::wstring currentWide = Utf16FromUtf8(*currentText);
        if (currentWide.empty())
        {
            continue;
        }
        pane.current = std::filesystem::path(currentWide);

        yyjson_val* view = GetObj(item, "view");
        if (view)
        {
            const auto displayText = GetString(view, "display");
            if (displayText)
            {
                pane.view.display = ParseFolderDisplayMode(*displayText);
            }

            const auto sortByText = GetString(view, "sortBy");
            if (sortByText)
            {
                pane.view.sortBy = ParseFolderSortBy(*sortByText);
            }

            bool sawSortDirection        = false;
            const auto sortDirectionText = GetString(view, "sortDirection");
            if (sortDirectionText)
            {
                pane.view.sortDirection = ParseFolderSortDirection(*sortDirectionText);
                sawSortDirection        = true;
            }
            if (! sawSortDirection)
            {
                pane.view.sortDirection = DefaultFolderSortDirection(pane.view.sortBy);
            }

            yyjson_val* statusBarVisible = yyjson_obj_get(view, "statusBarVisible");
            if (statusBarVisible && yyjson_is_bool(statusBarVisible))
            {
                pane.view.statusBarVisible = yyjson_get_bool(statusBarVisible);
            }
        }

        folderSettings.items.push_back(std::move(pane));
    }

    if (! folderSettings.items.empty())
    {
        if (folderSettings.active.empty())
        {
            folderSettings.active = folderSettings.items.front().slot;
        }
        out.folders = std::move(folderSettings);
    }
}

Common::Settings::MonitorFilterPreset ParsePreset(std::string_view preset)
{
    if (preset == "errorsOnly")
    {
        return Common::Settings::MonitorFilterPreset::ErrorsOnly;
    }
    if (preset == "errorsWarnings")
    {
        return Common::Settings::MonitorFilterPreset::ErrorsWarnings;
    }
    if (preset == "allTypes")
    {
        return Common::Settings::MonitorFilterPreset::AllTypes;
    }
    return Common::Settings::MonitorFilterPreset::Custom;
}

const char* PresetToString(Common::Settings::MonitorFilterPreset preset)
{
    switch (preset)
    {
        case Common::Settings::MonitorFilterPreset::ErrorsOnly: return "errorsOnly";
        case Common::Settings::MonitorFilterPreset::ErrorsWarnings: return "errorsWarnings";
        case Common::Settings::MonitorFilterPreset::AllTypes: return "allTypes";
        case Common::Settings::MonitorFilterPreset::Custom: return "custom";
    }
    return "custom";
}

void ParseMonitor(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* monitor = yyjson_obj_get(root, "monitor");
    if (! monitor || ! yyjson_is_obj(monitor))
    {
        return;
    }

    Common::Settings::MonitorSettings settings;

    yyjson_val* menu = GetObj(monitor, "menu");
    if (menu)
    {
        GetBool(menu, "toolbarVisible", settings.menu.toolbarVisible);
        GetBool(menu, "lineNumbersVisible", settings.menu.lineNumbersVisible);
        GetBool(menu, "alwaysOnTop", settings.menu.alwaysOnTop);
        GetBool(menu, "showIds", settings.menu.showIds);
        GetBool(menu, "autoScroll", settings.menu.autoScroll);
    }

    yyjson_val* filter = GetObj(monitor, "filter");
    if (filter)
    {
        GetUInt32(filter, "mask", settings.filter.mask);
        settings.filter.mask &= 31u;

        const auto preset = GetString(filter, "preset");
        if (preset)
        {
            settings.filter.preset = ParsePreset(*preset);
        }
    }

    out.monitor = std::move(settings);
}

void ParseMainMenu(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* mainMenu = yyjson_obj_get(root, "mainMenu");
    if (! mainMenu || ! yyjson_is_obj(mainMenu))
    {
        return;
    }

    Common::Settings::MainMenuState state;
    GetBool(mainMenu, "menuBarVisible", state.menuBarVisible);
    GetBool(mainMenu, "functionBarVisible", state.functionBarVisible);
    out.mainMenu = std::move(state);
}

void ParseStartup(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* startup = yyjson_obj_get(root, "startup");
    if (! startup || ! yyjson_is_obj(startup))
    {
        return;
    }

    Common::Settings::StartupSettings settings;
    GetBool(startup, "showSplash", settings.showSplash);
    out.startup = std::move(settings);
}

void ParseConnections(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* connections = yyjson_obj_get(root, "connections");
    if (! connections || ! yyjson_is_obj(connections))
    {
        return;
    }

    Common::Settings::ConnectionsSettings settings;

    if (yyjson_val* v = yyjson_obj_get(connections, "bypassWindowsHello"); v && yyjson_is_bool(v))
    {
        settings.bypassWindowsHello = yyjson_get_bool(v) != 0;
    }

    const auto parseTimeoutValue = [&](yyjson_val* v, uint64_t defaultValue) noexcept -> uint64_t
    {
        if (! v)
        {
            return defaultValue;
        }

        if (yyjson_is_uint(v))
        {
            return yyjson_get_uint(v);
        }
        if (yyjson_is_int(v))
        {
            const int64_t signedValue = yyjson_get_int(v);
            return signedValue < 0 ? 0u : static_cast<uint64_t>(signedValue);
        }

        return defaultValue;
    };

    if (yyjson_val* timeoutVal = yyjson_obj_get(connections, "windowsHelloReauthTimeoutMinute"); timeoutVal)
    {
        const uint64_t timeoutMinutes            = parseTimeoutValue(timeoutVal, settings.windowsHelloReauthTimeoutMinute);
        settings.windowsHelloReauthTimeoutMinute = static_cast<uint32_t>(std::min<uint64_t>(timeoutMinutes, 0xFFFFFFFFull));
    }
    else if (yyjson_val* legacyTimeoutVal = yyjson_obj_get(connections, "windowsHelloReauthTimeoutMs"); legacyTimeoutVal)
    {
        // Backward compatibility: accept old millisecond key.
        const uint64_t timeoutMs      = parseTimeoutValue(legacyTimeoutVal, static_cast<uint64_t>(settings.windowsHelloReauthTimeoutMinute) * 60'000ull);
        const uint64_t timeoutMinutes = timeoutMs / 60'000ull;
        settings.windowsHelloReauthTimeoutMinute = static_cast<uint32_t>(std::min<uint64_t>(timeoutMinutes, 0xFFFFFFFFull));
    }

    yyjson_val* items = GetArr(connections, "items");

    const auto trimWhitespace = [](std::wstring_view text) -> std::wstring
    {
        size_t start = 0;
        while (start < text.size() && std::iswspace(static_cast<wint_t>(text[start])) != 0)
        {
            ++start;
        }

        size_t end = text.size();
        while (end > start && std::iswspace(static_cast<wint_t>(text[end - 1])) != 0)
        {
            --end;
        }

        return std::wstring(text.substr(start, end - start));
    };

    const auto normalizeNameKey = [](std::wstring_view text) -> std::wstring
    {
        std::wstring key;
        key.reserve(text.size());
        for (const wchar_t ch : text)
        {
            key.push_back(static_cast<wchar_t>(std::towlower(ch)));
        }
        return key;
    };

    if (items)
    {
        const size_t count = yyjson_arr_size(items);
        settings.items.reserve(count);

        for (size_t i = 0; i < count; ++i)
        {
            yyjson_val* item = yyjson_arr_get(items, i);
            if (! item || ! yyjson_is_obj(item))
            {
                continue;
            }

            Common::Settings::ConnectionProfile profile;

            if (const auto idText = GetString(item, "id"); idText.has_value())
            {
                profile.id = Utf16FromUtf8(*idText);
            }
            if (const auto nameText = GetString(item, "name"); nameText.has_value())
            {
                profile.name = trimWhitespace(Utf16FromUtf8(*nameText));
            }
            if (const auto pluginIdText = GetString(item, "pluginId"); pluginIdText.has_value())
            {
                profile.pluginId = Utf16FromUtf8(*pluginIdText);
            }
            if (const auto hostText = GetString(item, "host"); hostText.has_value())
            {
                profile.host = Utf16FromUtf8(*hostText);
            }

            if (yyjson_val* v = yyjson_obj_get(item, "port"); v && yyjson_is_uint(v))
            {
                profile.port = static_cast<uint32_t>(yyjson_get_uint(v));
            }

            if (const auto initialPathText = GetString(item, "initialPath"); initialPathText.has_value())
            {
                profile.initialPath = Utf16FromUtf8(*initialPathText);
            }
            if (profile.initialPath.empty())
            {
                profile.initialPath = L"/";
            }

            if (const auto userNameText = GetString(item, "userName"); userNameText.has_value())
            {
                profile.userName = Utf16FromUtf8(*userNameText);
            }

            if (const auto authModeText = GetString(item, "authMode"); authModeText.has_value())
            {
                profile.authMode = ParseConnectionAuthMode(*authModeText);
            }

            if (yyjson_val* v = yyjson_obj_get(item, "savePassword"); v && yyjson_is_bool(v))
            {
                profile.savePassword = yyjson_get_bool(v) != 0;
            }
            if (yyjson_val* v = yyjson_obj_get(item, "requireWindowsHello"); v && yyjson_is_bool(v))
            {
                profile.requireWindowsHello = yyjson_get_bool(v) != 0;
            }

            if (yyjson_val* v = yyjson_obj_get(item, "extra"))
            {
                static_cast<void>(ConvertYyjsonToJsonValue(v, profile.extra));
            }

            const bool hostRequired = profile.pluginId != L"builtin/file-system-s3" && profile.pluginId != L"builtin/file-system-s3table";
            if (profile.id.empty() || profile.name.empty() || profile.pluginId.empty() || (hostRequired && profile.host.empty()))
            {
                continue;
            }

            settings.items.push_back(std::move(profile));
        }
    }

    if (! settings.items.empty())
    {
        std::unordered_set<std::wstring> usedNames;
        usedNames.reserve(settings.items.size());

        for (auto& profile : settings.items)
        {
            profile.name = trimWhitespace(profile.name);

            for (auto& ch : profile.name)
            {
                if (ch == L'/' || ch == L'\\')
                {
                    ch = L'-';
                }
            }

            if (profile.name.empty())
            {
                continue;
            }

            const std::wstring base = profile.name;
            std::wstring unique     = base;
            if (usedNames.contains(normalizeNameKey(unique)))
            {
                for (int suffix = 2; suffix < 10'000; ++suffix)
                {
                    unique = std::format(L"{} ({})", base, suffix);
                    if (! usedNames.contains(normalizeNameKey(unique)))
                    {
                        break;
                    }
                }
            }

            usedNames.insert(normalizeNameKey(unique));
            profile.name = std::move(unique);
        }
    }

    const Common::Settings::ConnectionsSettings defaults;
    const bool hasNonDefaultGlobals =
        settings.bypassWindowsHello != defaults.bypassWindowsHello || settings.windowsHelloReauthTimeoutMinute != defaults.windowsHelloReauthTimeoutMinute;

    if (! settings.items.empty() || hasNonDefaultGlobals)
    {
        out.connections = std::move(settings);
    }
}

void ParseFileOperations(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* fileOperations = yyjson_obj_get(root, "fileOperations");
    if (! fileOperations || ! yyjson_is_obj(fileOperations))
    {
        return;
    }

    Common::Settings::FileOperationsSettings settings;
    GetBool(fileOperations, "autoDismissSuccess", settings.autoDismissSuccess);
    GetBool(fileOperations, "diagnosticsInfoEnabled", settings.diagnosticsInfoEnabled);
    GetBool(fileOperations, "diagnosticsDebugEnabled", settings.diagnosticsDebugEnabled);

    uint32_t maxDiagnosticsLogFiles = settings.maxDiagnosticsLogFiles;
    if (GetUInt32(fileOperations, "maxDiagnosticsLogFiles", maxDiagnosticsLogFiles))
    {
        settings.maxDiagnosticsLogFiles = maxDiagnosticsLogFiles;
    }

    uint32_t maxIssueReportFiles = 0;
    if (GetUInt32(fileOperations, "maxIssueReportFiles", maxIssueReportFiles))
    {
        settings.maxIssueReportFiles = maxIssueReportFiles;
    }

    uint32_t maxDiagnosticsInMemory = 0;
    if (GetUInt32(fileOperations, "maxDiagnosticsInMemory", maxDiagnosticsInMemory))
    {
        settings.maxDiagnosticsInMemory = maxDiagnosticsInMemory;
    }

    uint32_t maxDiagnosticsPerFlush = 0;
    if (GetUInt32(fileOperations, "maxDiagnosticsPerFlush", maxDiagnosticsPerFlush))
    {
        settings.maxDiagnosticsPerFlush = maxDiagnosticsPerFlush;
    }

    uint32_t diagnosticsFlushIntervalMs = 0;
    if (GetUInt32(fileOperations, "diagnosticsFlushIntervalMs", diagnosticsFlushIntervalMs))
    {
        settings.diagnosticsFlushIntervalMs = diagnosticsFlushIntervalMs;
    }

    uint32_t diagnosticsCleanupIntervalMs = 0;
    if (GetUInt32(fileOperations, "diagnosticsCleanupIntervalMs", diagnosticsCleanupIntervalMs))
    {
        settings.diagnosticsCleanupIntervalMs = diagnosticsCleanupIntervalMs;
    }

    out.fileOperations = std::move(settings);
}

void ParseCompareDirectories(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* compare = yyjson_obj_get(root, "compareDirectories");
    if (! compare || ! yyjson_is_obj(compare))
    {
        return;
    }

    Common::Settings::CompareDirectoriesSettings settings;
    GetBool(compare, "compareSize", settings.compareSize);
    GetBool(compare, "compareDateTime", settings.compareDateTime);
    GetBool(compare, "compareAttributes", settings.compareAttributes);
    GetBool(compare, "compareContent", settings.compareContent);
    GetBool(compare, "compareSubdirectories", settings.compareSubdirectories);
    GetBool(compare, "compareSubdirectoryAttributes", settings.compareSubdirectoryAttributes);
    GetBool(compare, "selectSubdirsOnlyInOnePane", settings.selectSubdirsOnlyInOnePane);
    GetBool(compare, "ignoreFiles", settings.ignoreFiles);
    GetBool(compare, "ignoreDirectories", settings.ignoreDirectories);
    GetBool(compare, "showIdenticalItems", settings.showIdenticalItems);

    if (const auto ignoreFilesPatterns = GetString(compare, "ignoreFilesPatterns"))
    {
        settings.ignoreFilesPatterns = Utf16FromUtf8(ignoreFilesPatterns.value());
    }
    if (const auto ignoreDirectoriesPatterns = GetString(compare, "ignoreDirectoriesPatterns"))
    {
        settings.ignoreDirectoriesPatterns = Utf16FromUtf8(ignoreDirectoriesPatterns.value());
    }

    const Common::Settings::CompareDirectoriesSettings defaults{};
    const bool hasNonDefault = settings.compareSize != defaults.compareSize || settings.compareDateTime != defaults.compareDateTime ||
                               settings.compareAttributes != defaults.compareAttributes || settings.compareContent != defaults.compareContent ||
                               settings.compareSubdirectories != defaults.compareSubdirectories ||
                               settings.compareSubdirectoryAttributes != defaults.compareSubdirectoryAttributes ||
                               settings.selectSubdirsOnlyInOnePane != defaults.selectSubdirsOnlyInOnePane || settings.ignoreFiles != defaults.ignoreFiles ||
                               settings.ignoreDirectories != defaults.ignoreDirectories || settings.showIdenticalItems != defaults.showIdenticalItems ||
                               ! settings.ignoreFilesPatterns.empty() || ! settings.ignoreDirectoriesPatterns.empty();

    if (hasNonDefault)
    {
        out.compareDirectories = std::move(settings);
    }
}

void ParseShortcuts(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* shortcuts = yyjson_obj_get(root, "shortcuts");
    if (! shortcuts || ! yyjson_is_obj(shortcuts))
    {
        return;
    }

    Common::Settings::ShortcutsSettings settings;
    const bool isSchemaV5OrLater = out.schemaVersion >= 5;

    auto parseBindings = [&](const char* name, std::vector<Common::Settings::ShortcutBinding>& dest)
    {
        yyjson_val* arr = GetArr(shortcuts, name);
        if (! arr)
        {
            return;
        }

        const size_t count = yyjson_arr_size(arr);
        dest.reserve(count);

        for (size_t i = 0; i < count; ++i)
        {
            yyjson_val* binding = yyjson_arr_get(arr, i);
            if (! binding || ! yyjson_is_obj(binding))
            {
                continue;
            }

            uint32_t vk              = 0;
            uint32_t modifiers       = 0;
            const auto commandIdText = GetString(binding, "commandId");

            yyjson_val* vkVal = yyjson_obj_get(binding, "vk");
            if (! vkVal || ! commandIdText.has_value())
            {
                continue;
            }

            if (isSchemaV5OrLater)
            {
                if (! yyjson_is_str(vkVal))
                {
                    continue;
                }

                if (yyjson_obj_get(binding, "modifiers"))
                {
                    continue;
                }

                const char* vkText = yyjson_get_str(vkVal);
                if (! vkText || ! TryParseVkFromText(std::string_view(vkText), vk))
                {
                    continue;
                }

                if (yyjson_val* ctrlVal = yyjson_obj_get(binding, "ctrl"))
                {
                    if (! yyjson_is_bool(ctrlVal))
                    {
                        continue;
                    }
                    if (yyjson_get_bool(ctrlVal))
                    {
                        modifiers |= 1u;
                    }
                }

                if (yyjson_val* altVal = yyjson_obj_get(binding, "alt"))
                {
                    if (! yyjson_is_bool(altVal))
                    {
                        continue;
                    }
                    if (yyjson_get_bool(altVal))
                    {
                        modifiers |= 2u;
                    }
                }

                if (yyjson_val* shiftVal = yyjson_obj_get(binding, "shift"))
                {
                    if (! yyjson_is_bool(shiftVal))
                    {
                        continue;
                    }
                    if (yyjson_get_bool(shiftVal))
                    {
                        modifiers |= 4u;
                    }
                }
            }
            else
            {
                if (! yyjson_is_uint(vkVal))
                {
                    continue;
                }

                vk = static_cast<uint32_t>(yyjson_get_uint(vkVal));

                if (! GetUInt32(binding, "modifiers", modifiers))
                {
                    continue;
                }
            }

            if (vk > 0xFFu || modifiers > 7u)
            {
                continue;
            }

            const std::wstring commandId = Utf16FromUtf8(commandIdText.value());
            if (commandId.empty() || commandId.rfind(L"cmd/", 0) != 0)
            {
                continue;
            }

            Common::Settings::ShortcutBinding entry;
            entry.vk        = vk;
            entry.modifiers = modifiers;
            entry.commandId = commandId;
            dest.push_back(std::move(entry));
        }
    };

    parseBindings("functionBar", settings.functionBar);
    parseBindings("folderView", settings.folderView);

    out.shortcuts = std::move(settings);
}

void ParseCache(yyjson_val* root, Common::Settings::Settings& out)
{
    yyjson_val* cache = yyjson_obj_get(root, "cache");
    if (! cache || ! yyjson_is_obj(cache))
    {
        return;
    }

    Common::Settings::CacheSettings settings;

    yyjson_val* directoryInfo = GetObj(cache, "directoryInfo");
    if (directoryInfo)
    {
        uint64_t maxBytes = 0;
        if (GetDirectoryCacheBytes(directoryInfo, "maxBytes", maxBytes) && maxBytes > 0)
        {
            settings.directoryInfo.maxBytes = maxBytes;
        }

        uint32_t maxWatchers = 0;
        if (GetUInt32(directoryInfo, "maxWatchers", maxWatchers))
        {
            settings.directoryInfo.maxWatchers = maxWatchers;
        }

        uint32_t mruWatched = 0;
        if (GetUInt32(directoryInfo, "mruWatched", mruWatched))
        {
            settings.directoryInfo.mruWatched = mruWatched;
        }
    }

    out.cache = std::move(settings);
}

yyjson_mut_val* NewString(yyjson_mut_doc* doc, const std::wstring& value)
{
    const std::string utf8 = Utf8FromUtf16(value);
    if (utf8.empty() && ! value.empty())
    {
        return yyjson_mut_strcpy(doc, "");
    }
    return yyjson_mut_strncpy(doc, utf8.c_str(), utf8.size());
}

} // namespace

namespace Common::Settings
{
namespace
{
[[nodiscard]] std::filesystem::path GetSettingsDirectoryPath() noexcept
{
    std::filesystem::path base;

    wil::unique_cotaskmem_string localAppData;
    const HRESULT hr = SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, localAppData.put());
    if (SUCCEEDED(hr) && localAppData)
    {
        base = std::filesystem::path(localAppData.get());
    }
    else
    {
        const DWORD required = GetEnvironmentVariableW(L"LOCALAPPDATA", nullptr, 0);
        if (required == 0)
        {
            return {};
        }

        std::wstring buffer(static_cast<size_t>(required), L'\0');
        const DWORD written = GetEnvironmentVariableW(L"LOCALAPPDATA", buffer.data(), required);
        if (written == 0 || written >= required)
        {
            return {};
        }

        buffer.resize(static_cast<size_t>(written));
        base = std::filesystem::path(buffer);
    }

    base /= kCompanyDirectoryName;
    base /= kSettingsDirectoryName;
    return base;
}

[[nodiscard]] std::wstring GetLegacySettingsFileName(std::wstring_view appId)
{
    std::wstring fileName(appId);
    fileName += L".settings.json";
    return fileName;
}

[[nodiscard]] std::wstring GetVersionedSettingsFileName(std::wstring_view appId)
{
    std::wstring fileName(appId);
    fileName += std::format(L"-{}.{}", VERSINFO_MAJOR, VERSINFO_MINORA);
    fileName += L".settings.json";
    return fileName;
}

[[nodiscard]] std::wstring GetDebugSettingsFileName(std::wstring_view appId)
{
    std::wstring fileName(appId);
    fileName += L"-debug.settings.json";
    return fileName;
}

[[nodiscard]] std::filesystem::path GetLegacySettingsPath(std::wstring_view appId) noexcept
{
    const std::filesystem::path base = GetSettingsDirectoryPath();
    if (base.empty())
    {
        return {};
    }
    return base / GetLegacySettingsFileName(appId);
}

[[nodiscard]] std::filesystem::path GetVersionedSettingsPath(std::wstring_view appId) noexcept
{
    const std::filesystem::path base = GetSettingsDirectoryPath();
    if (base.empty())
    {
        return {};
    }
    return base / GetVersionedSettingsFileName(appId);
}

[[nodiscard]] bool IsSettingsFilePresent(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    const DWORD attrs = GetFileAttributesW(path.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES)
    {
        return false;
    }

    return (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0;
}
} // namespace

std::filesystem::path GetSettingsPath(std::wstring_view appId) noexcept
{
    if (appId.empty())
    {
        return {};
    }

    const std::filesystem::path base = GetSettingsDirectoryPath();
    if (base.empty())
    {
        return {};
    }

#ifdef _DEBUG
    return base / GetDebugSettingsFileName(appId);
#else
    return base / GetVersionedSettingsFileName(appId);
#endif
}

std::filesystem::path GetSettingsSchemaPath(std::wstring_view appId) noexcept
{
    if (appId.empty())
    {
        return {};
    }

    const std::filesystem::path settingsPath = GetSettingsPath(appId);
    if (settingsPath.empty())
    {
        return {};
    }

    std::wstring fileName(appId);
    fileName += L".settings.schema.json";
    return settingsPath.parent_path() / fileName;
}

std::string_view GetSettingsStoreSchemaJsonUtf8() noexcept
{
    static std::once_flag once;
    static std::string cached;

    std::call_once(once,
                   []
                   {
                       const std::filesystem::path schemaPath = GetShippedSettingsStoreSchemaPath();
                       if (schemaPath.empty())
                       {
                           return;
                       }

                       const DWORD attrs = GetFileAttributesW(schemaPath.c_str());
                       if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0)
                       {
                           Debug::Warning(L"Shipped settings schema file is missing: '{}'", schemaPath.c_str());
                           return;
                       }

                       std::string bytes;
                       const HRESULT hr = ReadFileBytes(schemaPath, bytes);
                       if (FAILED(hr))
                       {
                           Debug::Warning(L"Failed to read shipped settings schema '{}': hr=0x{:08X}", schemaPath.c_str(), static_cast<unsigned long>(hr));
                           return;
                       }

                       StripUtf8BomInPlace(bytes);
                       cached = std::move(bytes);
                   });

    return cached;
}

HRESULT LoadSettings(std::wstring_view appId, Settings& out) noexcept
{
    out = Settings{};

    std::filesystem::path path = GetSettingsPath(appId);
    if (path.empty())
    {
        return E_FAIL;
    }

#ifdef _DEBUG
    if (! IsSettingsFilePresent(path))
    {
        const std::filesystem::path versionedPath = GetVersionedSettingsPath(appId);
        if (IsSettingsFilePresent(versionedPath))
        {
            path = versionedPath;
        }
        else
        {
            const std::filesystem::path legacyPath = GetLegacySettingsPath(appId);
            if (IsSettingsFilePresent(legacyPath))
            {
                path = legacyPath;
            }
            else
            {
                return S_FALSE;
            }
        }
    }
#else
    if (! IsSettingsFilePresent(path))
    {
        const std::filesystem::path legacyPath = GetLegacySettingsPath(appId);
        if (IsSettingsFilePresent(legacyPath))
        {
            path = legacyPath;
        }
        else
        {
            return S_FALSE;
        }
    }
#endif

    if (path.empty())
    {
        return S_FALSE;
    }

    std::string bytes;
    const HRESULT readHr = ReadFileBytes(path, bytes);
    if (FAILED(readHr))
    {
        return S_FALSE;
    }

    yyjson_read_err err{};
    yyjson_doc* doc = yyjson_read_opts(bytes.data(), bytes.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
    if (! doc)
    {
        LogJsonParseError(L"settings file", path, err);
        BackupBadSettingsFile(path);
        return S_FALSE;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        Debug::Error(L"Failed to parse settings file '{}': expected object at root", path.c_str());
        BackupBadSettingsFile(path);
        return S_FALSE;
    }

    yyjson_val* schema = yyjson_obj_get(root, "schemaVersion");
    if (! schema || ! yyjson_is_int(schema))
    {
        Debug::Error(L"Unsupported schema version in settings file '{}'", path.c_str());
        BackupBadSettingsFile(path);
        return S_FALSE;
    }

    const int64_t schemaVersion = yyjson_get_int(schema);
    if (schemaVersion != 6 && schemaVersion != 7 && schemaVersion != 8 && schemaVersion != 9)
    {
        Debug::Error(L"Unsupported schema version in settings file '{}'", path.c_str());
        BackupBadSettingsFile(path);
        return S_FALSE;
    }

    out.schemaVersion = static_cast<uint32_t>(schemaVersion);

    ParseWindows(root, out);
    ParseTheme(root, out);
    if (schemaVersion >= 2)
    {
        ParsePlugins(root, out);
    }
    ParseExtensions(root, out);
    ParseShortcuts(root, out);
    ParseCache(root, out);
    ParseFolders(root, out);
    ParseMonitor(root, out);
    ParseMainMenu(root, out);
    ParseStartup(root, out);
    ParseConnections(root, out);
    ParseFileOperations(root, out);
    ParseCompareDirectories(root, out);

    out.schemaVersion = 9;

    return S_OK;
}

std::wstring FormatColor(uint32_t argb)
{
    const uint8_t a = static_cast<uint8_t>((argb >> 24) & 0xFFu);
    const uint8_t r = static_cast<uint8_t>((argb >> 16) & 0xFFu);
    const uint8_t g = static_cast<uint8_t>((argb >> 8) & 0xFFu);
    const uint8_t b = static_cast<uint8_t>(argb & 0xFFu);

    if (a == 0xFFu)
    {
        return std::format(L"#{:02X}{:02X}{:02X}", r, g, b);
    }
    return std::format(L"#{:02X}{:02X}{:02X}{:02X}", a, r, g, b);
}

bool TryParseColor(std::wstring_view hex, uint32_t& argb) noexcept
{
    const std::string utf8 = Utf8FromUtf16(hex);
    if (utf8.empty() && ! hex.empty())
    {
        return false;
    }
    return TryParseColorUtf8(utf8, argb);
}

HRESULT SaveSettings(std::wstring_view appId, const Settings& settings) noexcept
{
    const std::filesystem::path settingsPath = GetSettingsPath(appId);
    if (settingsPath.empty())
    {
        return E_FAIL;
    }

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    if (! root)
    {
        return E_OUTOFMEMORY;
    }
    yyjson_mut_doc_set_root(doc, root);

    std::wstring schemaRef(L"./");
    schemaRef.append(appId);
    schemaRef.append(L".settings.schema.json");
    yyjson_mut_obj_add_val(doc, root, "$schema", NewString(doc, schemaRef));

    yyjson_mut_obj_add_int(doc, root, "schemaVersion", 9);

    yyjson_mut_val* windows = nullptr;
    {
        std::vector<std::wstring> windowIds;
        windowIds.reserve(settings.windows.size());
        for (const auto& [id, _] : settings.windows)
        {
            windowIds.push_back(id);
        }
        std::sort(windowIds.begin(), windowIds.end());

        for (const auto& id : windowIds)
        {
            const auto it = settings.windows.find(id);
            if (it == settings.windows.end())
            {
                continue;
            }

            const WindowPlacement& wp = it->second;
            yyjson_mut_val* wpObj     = yyjson_mut_obj(doc);
            if (! wpObj)
            {
                return E_OUTOFMEMORY;
            }

            const char* stateText = wp.state == WindowState::Maximized ? "maximized" : "normal";
            yyjson_mut_obj_add_str(doc, wpObj, "state", stateText);

            yyjson_mut_val* bounds = yyjson_mut_obj(doc);
            if (! bounds)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_int(doc, bounds, "x", wp.bounds.x);
            yyjson_mut_obj_add_int(doc, bounds, "y", wp.bounds.y);
            yyjson_mut_obj_add_int(doc, bounds, "width", std::max(1, wp.bounds.width));
            yyjson_mut_obj_add_int(doc, bounds, "height", std::max(1, wp.bounds.height));
            yyjson_mut_obj_add_val(doc, wpObj, "bounds", bounds);

            if (wp.dpi)
            {
                yyjson_mut_obj_add_int(doc, wpObj, "dpi", static_cast<int>(*wp.dpi));
            }

            const std::string idUtf8 = Utf8FromUtf16(id);
            if (idUtf8.empty() && ! id.empty())
            {
                continue;
            }

            yyjson_mut_val* key = yyjson_mut_strncpy(doc, idUtf8.c_str(), idUtf8.size());
            if (! key)
            {
                return E_OUTOFMEMORY;
            }

            if (! windows)
            {
                windows = yyjson_mut_obj(doc);
                if (! windows)
                {
                    return E_OUTOFMEMORY;
                }
            }
            yyjson_mut_obj_add(windows, key, wpObj);
        }
    }

    if (windows)
    {
        yyjson_mut_obj_add_val(doc, root, "windows", windows);
    }

    {
        const ThemeSettings defaults{};
        const std::wstring currentThemeId = settings.theme.currentThemeId.empty() ? defaults.currentThemeId : settings.theme.currentThemeId;

        const bool writeThemeId = (currentThemeId != defaults.currentThemeId);
        const bool writeThemes  = ! settings.theme.themes.empty();
        if (writeThemeId || writeThemes)
        {
            yyjson_mut_val* theme = yyjson_mut_obj(doc);
            if (! theme)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, root, "theme", theme);

            if (writeThemeId)
            {
                yyjson_mut_obj_add_val(doc, theme, "currentThemeId", NewString(doc, currentThemeId));
            }

            if (writeThemes)
            {
                yyjson_mut_val* themeArr = yyjson_mut_arr(doc);
                if (! themeArr)
                {
                    return E_OUTOFMEMORY;
                }
                yyjson_mut_obj_add_val(doc, theme, "themes", themeArr);

                std::vector<const ThemeDefinition*> defs;
                defs.reserve(settings.theme.themes.size());
                for (const ThemeDefinition& def : settings.theme.themes)
                {
                    defs.push_back(&def);
                }
                std::sort(defs.begin(), defs.end(), [](const ThemeDefinition* a, const ThemeDefinition* b) { return a->id < b->id; });

                for (const ThemeDefinition* def : defs)
                {
                    yyjson_mut_val* defObj = yyjson_mut_obj(doc);
                    if (! defObj)
                    {
                        return E_OUTOFMEMORY;
                    }
                    yyjson_mut_obj_add_val(doc, defObj, "id", NewString(doc, def->id));
                    yyjson_mut_obj_add_val(doc, defObj, "name", NewString(doc, def->name));
                    yyjson_mut_obj_add_val(doc, defObj, "baseThemeId", NewString(doc, def->baseThemeId));

                    yyjson_mut_val* colors = yyjson_mut_obj(doc);
                    if (! colors)
                    {
                        return E_OUTOFMEMORY;
                    }
                    yyjson_mut_obj_add_val(doc, defObj, "colors", colors);

                    std::vector<std::wstring> colorKeys;
                    colorKeys.reserve(def->colors.size());
                    for (const auto& [k, _] : def->colors)
                    {
                        colorKeys.push_back(k);
                    }
                    std::sort(colorKeys.begin(), colorKeys.end());

                    for (const auto& k : colorKeys)
                    {
                        const auto it = def->colors.find(k);
                        if (it == def->colors.end())
                        {
                            continue;
                        }
                        const std::wstring colorText = FormatColor(it->second);
                        const std::string keyUtf8    = Utf8FromUtf16(k);
                        if (keyUtf8.empty() && ! k.empty())
                        {
                            continue;
                        }

                        yyjson_mut_val* key = yyjson_mut_strncpy(doc, keyUtf8.c_str(), keyUtf8.size());
                        if (! key)
                        {
                            return E_OUTOFMEMORY;
                        }

                        yyjson_mut_val* value = NewString(doc, colorText);
                        if (! value)
                        {
                            return E_OUTOFMEMORY;
                        }

                        yyjson_mut_obj_add(colors, key, value);
                    }

                    yyjson_mut_arr_add_val(themeArr, defObj);
                }
            }
        }
    }

    {
        const PluginsSettings defaults{};

        const std::wstring currentPluginId =
            settings.plugins.currentFileSystemPluginId.empty() ? defaults.currentFileSystemPluginId : settings.plugins.currentFileSystemPluginId;

        std::vector<std::wstring> disabledIds = settings.plugins.disabledPluginIds;
        std::sort(disabledIds.begin(), disabledIds.end());
        disabledIds.erase(std::unique(disabledIds.begin(), disabledIds.end()), disabledIds.end());
        std::erase_if(disabledIds, [](const std::wstring& id) { return id.empty(); });

        std::vector<std::wstring> customPaths;
        customPaths.reserve(settings.plugins.customPluginPaths.size());
        for (const auto& customPath : settings.plugins.customPluginPaths)
        {
            if (customPath.empty())
            {
                continue;
            }
            customPaths.push_back(customPath.wstring());
        }
        std::sort(customPaths.begin(), customPaths.end());
        customPaths.erase(std::unique(customPaths.begin(), customPaths.end()), customPaths.end());
        std::erase_if(customPaths, [](const std::wstring& path) { return path.empty(); });

        std::vector<std::wstring> configIds;
        configIds.reserve(settings.plugins.configurationByPluginId.size());
        for (const auto& [id, _] : settings.plugins.configurationByPluginId)
        {
            configIds.push_back(id);
        }
        std::sort(configIds.begin(), configIds.end());
        configIds.erase(std::unique(configIds.begin(), configIds.end()), configIds.end());
        std::erase_if(configIds, [](const std::wstring& id) { return id.empty(); });

        const bool writeCurrentPluginId = (currentPluginId != defaults.currentFileSystemPluginId);
        const bool writeDisabledIds     = ! disabledIds.empty();
        const bool writeCustomPaths     = ! customPaths.empty();
        const bool writeConfigs         = ! configIds.empty();

        if (writeCurrentPluginId || writeDisabledIds || writeCustomPaths || writeConfigs)
        {
            yyjson_mut_val* plugins = yyjson_mut_obj(doc);
            if (! plugins)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, root, "plugins", plugins);

            if (writeCurrentPluginId)
            {
                yyjson_mut_obj_add_val(doc, plugins, "currentFileSystemPluginId", NewString(doc, currentPluginId));
            }

            if (writeDisabledIds)
            {
                yyjson_mut_val* disabled = yyjson_mut_arr(doc);
                if (! disabled)
                {
                    return E_OUTOFMEMORY;
                }
                yyjson_mut_obj_add_val(doc, plugins, "disabledPluginIds", disabled);
                for (const auto& id : disabledIds)
                {
                    yyjson_mut_arr_add_val(disabled, NewString(doc, id));
                }
            }

            if (writeCustomPaths)
            {
                yyjson_mut_val* custom = yyjson_mut_arr(doc);
                if (! custom)
                {
                    return E_OUTOFMEMORY;
                }
                yyjson_mut_obj_add_val(doc, plugins, "customPluginPaths", custom);
                for (const auto& customPath : customPaths)
                {
                    yyjson_mut_arr_add_val(custom, NewString(doc, customPath));
                }
            }

            if (writeConfigs)
            {
                yyjson_mut_val* configs = yyjson_mut_obj(doc);
                if (! configs)
                {
                    return E_OUTOFMEMORY;
                }
                bool wroteAny = false;

                for (const auto& id : configIds)
                {
                    const auto it = settings.plugins.configurationByPluginId.find(id);
                    if (it == settings.plugins.configurationByPluginId.end())
                    {
                        continue;
                    }

                    const Common::Settings::JsonValue& configValue = it->second;

                    const std::string idUtf8 = Utf8FromUtf16(id);
                    if (idUtf8.empty() && ! id.empty())
                    {
                        continue;
                    }

                    yyjson_mut_val* key = yyjson_mut_strncpy(doc, idUtf8.c_str(), idUtf8.size());
                    if (! key)
                    {
                        return E_OUTOFMEMORY;
                    }

                    HRESULT valueHr       = S_OK;
                    yyjson_mut_val* value = NewYyjsonFromJsonValue(doc, configValue, valueHr);
                    if (! value)
                    {
                        return FAILED(valueHr) ? valueHr : E_OUTOFMEMORY;
                    }

                    yyjson_mut_obj_add(configs, key, value);
                    wroteAny = true;
                }

                if (wroteAny)
                {
                    yyjson_mut_obj_add_val(doc, plugins, "configurationByPluginId", configs);
                }
            }
        }
    }

    {
        const ExtensionsSettings defaults{};
        const bool writeFileSystems = (settings.extensions.openWithFileSystemByExtension != defaults.openWithFileSystemByExtension);
        const bool writeViewers     = (settings.extensions.openWithViewerByExtension != defaults.openWithViewerByExtension);

        if (writeFileSystems || writeViewers)
        {
            yyjson_mut_val* extensions = yyjson_mut_obj(doc);
            if (! extensions)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, root, "extensions", extensions);

            auto writeExtMap = [&](const char* field, const std::unordered_map<std::wstring, std::wstring>& map) -> HRESULT
            {
                yyjson_mut_val* openWith = yyjson_mut_obj(doc);
                if (! openWith)
                {
                    return E_OUTOFMEMORY;
                }
                yyjson_mut_obj_add_val(doc, extensions, field, openWith);

                std::vector<std::wstring> exts;
                exts.reserve(map.size());
                for (const auto& [ext, _] : map)
                {
                    exts.push_back(ext);
                }

                std::sort(exts.begin(), exts.end());
                exts.erase(std::unique(exts.begin(), exts.end()), exts.end());

                for (const auto& ext : exts)
                {
                    if (ext.empty())
                    {
                        continue;
                    }

                    const auto it = map.find(ext);
                    if (it == map.end())
                    {
                        continue;
                    }

                    const std::string extUtf8 = Utf8FromUtf16(ext);
                    if (extUtf8.empty() && ! ext.empty())
                    {
                        continue;
                    }

                    yyjson_mut_val* key = yyjson_mut_strncpy(doc, extUtf8.c_str(), extUtf8.size());
                    if (! key)
                    {
                        return E_OUTOFMEMORY;
                    }

                    yyjson_mut_val* value = NewString(doc, it->second);
                    if (! value)
                    {
                        return E_OUTOFMEMORY;
                    }

                    yyjson_mut_obj_add(openWith, key, value);
                }

                return S_OK;
            };

            if (writeFileSystems)
            {
                HRESULT hr = writeExtMap("openWithFileSystemByExtension", settings.extensions.openWithFileSystemByExtension);
                if (FAILED(hr))
                {
                    return hr;
                }
            }

            if (writeViewers)
            {
                HRESULT hr = writeExtMap("openWithViewerByExtension", settings.extensions.openWithViewerByExtension);
                if (FAILED(hr))
                {
                    return hr;
                }
            }
        }
    }

    if (settings.shortcuts)
    {
        yyjson_mut_val* shortcuts = yyjson_mut_obj(doc);
        if (! shortcuts)
        {
            return E_OUTOFMEMORY;
        }
        yyjson_mut_obj_add_val(doc, root, "shortcuts", shortcuts);

        auto addBindings = [&](const char* name, const std::vector<ShortcutBinding>& bindings) -> HRESULT
        {
            yyjson_mut_val* arr = yyjson_mut_arr(doc);
            if (! arr)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, shortcuts, name, arr);

            std::vector<const ShortcutBinding*> items;
            items.reserve(bindings.size());
            for (const auto& binding : bindings)
            {
                if (binding.commandId.empty())
                {
                    continue;
                }
                items.push_back(&binding);
            }

            std::sort(items.begin(),
                      items.end(),
                      [](const ShortcutBinding* a, const ShortcutBinding* b)
                      {
                          if (a->vk != b->vk)
                          {
                              return a->vk < b->vk;
                          }
                          if (a->modifiers != b->modifiers)
                          {
                              return a->modifiers < b->modifiers;
                          }
                          return a->commandId < b->commandId;
                      });

            for (const ShortcutBinding* binding : items)
            {
                if (! binding)
                {
                    continue;
                }

                yyjson_mut_val* obj = yyjson_mut_obj(doc);
                if (! obj)
                {
                    return E_OUTOFMEMORY;
                }

                const std::string vkText = VkToStableName(binding->vk);
                yyjson_mut_val* vkVal    = yyjson_mut_strncpy(doc, vkText.c_str(), vkText.size());
                if (! vkVal)
                {
                    return E_OUTOFMEMORY;
                }
                yyjson_mut_obj_add_val(doc, obj, "vk", vkVal);

                const uint32_t modifiers = binding->modifiers & 0x7u;
                if ((modifiers & 1u) != 0u)
                {
                    yyjson_mut_obj_add_bool(doc, obj, "ctrl", true);
                }
                if ((modifiers & 2u) != 0u)
                {
                    yyjson_mut_obj_add_bool(doc, obj, "alt", true);
                }
                if ((modifiers & 4u) != 0u)
                {
                    yyjson_mut_obj_add_bool(doc, obj, "shift", true);
                }

                yyjson_mut_val* commandId = NewString(doc, binding->commandId);
                if (! commandId)
                {
                    return E_OUTOFMEMORY;
                }
                yyjson_mut_obj_add_val(doc, obj, "commandId", commandId);
                yyjson_mut_arr_add_val(arr, obj);
            }

            return S_OK;
        };

        HRESULT hr = addBindings("functionBar", settings.shortcuts->functionBar);
        if (FAILED(hr))
        {
            return hr;
        }
        hr = addBindings("folderView", settings.shortcuts->folderView);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    if (settings.mainMenu)
    {
        const MainMenuState defaults{};
        const bool writeMenuBarVisible     = (settings.mainMenu->menuBarVisible != defaults.menuBarVisible);
        const bool writeFunctionBarVisible = (settings.mainMenu->functionBarVisible != defaults.functionBarVisible);
        if (writeMenuBarVisible || writeFunctionBarVisible)
        {
            yyjson_mut_val* mainMenu = yyjson_mut_obj(doc);
            if (! mainMenu)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, root, "mainMenu", mainMenu);

            if (writeMenuBarVisible)
            {
                yyjson_mut_obj_add_bool(doc, mainMenu, "menuBarVisible", settings.mainMenu->menuBarVisible);
            }
            if (writeFunctionBarVisible)
            {
                yyjson_mut_obj_add_bool(doc, mainMenu, "functionBarVisible", settings.mainMenu->functionBarVisible);
            }
        }
    }

    if (settings.startup)
    {
        const StartupSettings defaults{};
        const bool writeShowSplash = (settings.startup->showSplash != defaults.showSplash);
        if (writeShowSplash)
        {
            yyjson_mut_val* startup = yyjson_mut_obj(doc);
            if (! startup)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, root, "startup", startup);
            yyjson_mut_obj_add_bool(doc, startup, "showSplash", settings.startup->showSplash);
        }
    }

    if (settings.cache)
    {
        const DirectoryInfoCacheSettings& directoryInfoSettings = settings.cache->directoryInfo;

        yyjson_mut_val* cache         = nullptr;
        yyjson_mut_val* directoryInfo = nullptr;
        bool wroteDirectoryInfo       = false;

        auto ensureDirectoryInfo = [&]() noexcept -> bool
        {
            if (! cache)
            {
                cache = yyjson_mut_obj(doc);
                if (! cache)
                {
                    return false;
                }
            }
            if (! directoryInfo)
            {
                directoryInfo = yyjson_mut_obj(doc);
                if (! directoryInfo)
                {
                    return false;
                }
            }
            return true;
        };

        if (directoryInfoSettings.maxBytes && *directoryInfoSettings.maxBytes > 0)
        {
            if (! ensureDirectoryInfo())
            {
                return E_OUTOFMEMORY;
            }

            // Persist as KiB so the JSON is readable and matches the accepted input format.
            const uint64_t valueBytes = *directoryInfoSettings.maxBytes;
            const uint64_t kiloBytes  = (valueBytes + 1023ull) / 1024ull;
            if (kiloBytes <= static_cast<uint64_t>(std::numeric_limits<int64_t>::max()))
            {
                yyjson_mut_obj_add_int(doc, directoryInfo, "maxBytes", static_cast<int64_t>(kiloBytes));
            }
            else
            {
                yyjson_mut_obj_add_int(doc, directoryInfo, "maxBytes", std::numeric_limits<int64_t>::max());
            }
            wroteDirectoryInfo = true;
        }

        if (directoryInfoSettings.maxWatchers)
        {
            if (! ensureDirectoryInfo())
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_int(doc, directoryInfo, "maxWatchers", static_cast<int64_t>(*directoryInfoSettings.maxWatchers));
            wroteDirectoryInfo = true;
        }

        if (directoryInfoSettings.mruWatched)
        {
            if (! ensureDirectoryInfo())
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_int(doc, directoryInfo, "mruWatched", static_cast<int64_t>(*directoryInfoSettings.mruWatched));
            wroteDirectoryInfo = true;
        }

        if (wroteDirectoryInfo && cache && directoryInfo)
        {
            yyjson_mut_obj_add_val(doc, cache, "directoryInfo", directoryInfo);
            yyjson_mut_obj_add_val(doc, root, "cache", cache);
        }
    }

    if (settings.folders && ! settings.folders->items.empty())
    {
        const FoldersSettings defaults{};

        std::vector<const FolderPane*> panes;
        panes.reserve(settings.folders->items.size());
        for (const auto& pane : settings.folders->items)
        {
            if (pane.slot.empty() || pane.current.empty())
            {
                continue;
            }
            panes.push_back(&pane);
        }

        if (! panes.empty())
        {
            yyjson_mut_val* folders = yyjson_mut_obj(doc);
            if (! folders)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, root, "folders", folders);

            std::sort(panes.begin(), panes.end(), [](const FolderPane* a, const FolderPane* b) { return a->slot < b->slot; });

            const std::wstring defaultActiveSlot = panes.front()->slot;
            const std::wstring activeSlot        = settings.folders->active.empty() ? defaultActiveSlot : settings.folders->active;
            if (! activeSlot.empty() && activeSlot != defaultActiveSlot)
            {
                yyjson_mut_obj_add_val(doc, folders, "active", NewString(doc, activeSlot));
            }

            const float splitRatio                = std::clamp(settings.folders->layout.splitRatio, 0.0f, 1.0f);
            const bool writeSplitRatio            = (std::abs(splitRatio - defaults.layout.splitRatio) > 0.0001f);
            const bool writeZoomedPane            = settings.folders->layout.zoomedPane.has_value() && ! settings.folders->layout.zoomedPane.value().empty();
            const bool writeZoomRestoreSplitRatio = settings.folders->layout.zoomRestoreSplitRatio.has_value();

            if (writeSplitRatio || writeZoomedPane || writeZoomRestoreSplitRatio)
            {
                yyjson_mut_val* layout = yyjson_mut_obj(doc);
                if (! layout)
                {
                    return E_OUTOFMEMORY;
                }
                yyjson_mut_obj_add_val(doc, folders, "layout", layout);

                if (writeSplitRatio)
                {
                    yyjson_mut_obj_add_real(doc, layout, "splitRatio", static_cast<double>(splitRatio));
                }
                if (writeZoomedPane)
                {
                    yyjson_mut_obj_add_val(doc, layout, "zoomedPane", NewString(doc, settings.folders->layout.zoomedPane.value()));
                }
                if (writeZoomRestoreSplitRatio)
                {
                    const float zoomRestoreSplitRatio = std::clamp(settings.folders->layout.zoomRestoreSplitRatio.value(), 0.0f, 1.0f);
                    yyjson_mut_obj_add_real(doc, layout, "zoomRestoreSplitRatio", static_cast<double>(zoomRestoreSplitRatio));
                }
            }

            const uint32_t historyMax = std::clamp(settings.folders->historyMax, 1u, 50u);
            if (historyMax != defaults.historyMax)
            {
                yyjson_mut_obj_add_int(doc, folders, "historyMax", static_cast<int64_t>(historyMax));
            }

            {
                yyjson_mut_val* history = yyjson_mut_arr(doc);
                if (! history)
                {
                    return E_OUTOFMEMORY;
                }
                size_t historyWritten = 0;
                for (const auto& entry : settings.folders->history)
                {
                    if (entry.empty())
                    {
                        continue;
                    }

                    yyjson_mut_arr_add_val(history, NewString(doc, entry.wstring()));
                    ++historyWritten;
                    if (historyWritten >= static_cast<size_t>(historyMax))
                    {
                        break;
                    }
                }

                if (historyWritten > 0)
                {
                    yyjson_mut_obj_add_val(doc, folders, "history", history);
                }
            }

            yyjson_mut_val* items = yyjson_mut_arr(doc);
            if (! items)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, folders, "items", items);

            for (const FolderPane* pane : panes)
            {
                yyjson_mut_val* paneObj = yyjson_mut_obj(doc);
                if (! paneObj)
                {
                    return E_OUTOFMEMORY;
                }
                yyjson_mut_obj_add_val(doc, paneObj, "slot", NewString(doc, pane->slot));
                yyjson_mut_obj_add_val(doc, paneObj, "current", NewString(doc, pane->current.wstring()));

                const FolderViewSettings viewDefaults{};

                yyjson_mut_val* view = yyjson_mut_obj(doc);
                if (! view)
                {
                    return E_OUTOFMEMORY;
                }
                bool wroteView = false;

                if (pane->view.display != viewDefaults.display)
                {
                    yyjson_mut_obj_add_str(doc, view, "display", FolderDisplayModeToString(pane->view.display));
                    wroteView = true;
                }

                if (pane->view.sortBy != viewDefaults.sortBy)
                {
                    yyjson_mut_obj_add_str(doc, view, "sortBy", FolderSortByToString(pane->view.sortBy));
                    wroteView = true;
                }

                const FolderSortDirection defaultDirection = DefaultFolderSortDirection(pane->view.sortBy);
                if (pane->view.sortDirection != defaultDirection)
                {
                    yyjson_mut_obj_add_str(doc, view, "sortDirection", FolderSortDirectionToString(pane->view.sortDirection));
                    wroteView = true;
                }

                if (pane->view.statusBarVisible != viewDefaults.statusBarVisible)
                {
                    yyjson_mut_obj_add_bool(doc, view, "statusBarVisible", pane->view.statusBarVisible);
                    wroteView = true;
                }

                if (wroteView)
                {
                    yyjson_mut_obj_add_val(doc, paneObj, "view", view);
                }

                yyjson_mut_arr_add_val(items, paneObj);
            }
        }
    }

    if (settings.monitor)
    {
        const MonitorSettings defaults{};

        yyjson_mut_val* monitor = nullptr;

        auto ensureMonitor = [&]() noexcept -> bool
        {
            if (! monitor)
            {
                monitor = yyjson_mut_obj(doc);
                if (! monitor)
                {
                    return false;
                }
            }
            return true;
        };

        yyjson_mut_val* menu = yyjson_mut_obj(doc);
        if (! menu)
        {
            return E_OUTOFMEMORY;
        }
        bool wroteMenu = false;
        if (settings.monitor->menu.toolbarVisible != defaults.menu.toolbarVisible)
        {
            yyjson_mut_obj_add_bool(doc, menu, "toolbarVisible", settings.monitor->menu.toolbarVisible);
            wroteMenu = true;
        }
        if (settings.monitor->menu.lineNumbersVisible != defaults.menu.lineNumbersVisible)
        {
            yyjson_mut_obj_add_bool(doc, menu, "lineNumbersVisible", settings.monitor->menu.lineNumbersVisible);
            wroteMenu = true;
        }
        if (settings.monitor->menu.alwaysOnTop != defaults.menu.alwaysOnTop)
        {
            yyjson_mut_obj_add_bool(doc, menu, "alwaysOnTop", settings.monitor->menu.alwaysOnTop);
            wroteMenu = true;
        }
        if (settings.monitor->menu.showIds != defaults.menu.showIds)
        {
            yyjson_mut_obj_add_bool(doc, menu, "showIds", settings.monitor->menu.showIds);
            wroteMenu = true;
        }
        if (settings.monitor->menu.autoScroll != defaults.menu.autoScroll)
        {
            yyjson_mut_obj_add_bool(doc, menu, "autoScroll", settings.monitor->menu.autoScroll);
            wroteMenu = true;
        }

        if (wroteMenu)
        {
            if (! ensureMonitor())
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, monitor, "menu", menu);
        }

        yyjson_mut_val* filter = yyjson_mut_obj(doc);
        if (! filter)
        {
            return E_OUTOFMEMORY;
        }
        bool wroteFilter    = false;
        const uint32_t mask = settings.monitor->filter.mask & 31u;
        if (mask != (defaults.filter.mask & 31u))
        {
            yyjson_mut_obj_add_int(doc, filter, "mask", static_cast<int>(mask));
            wroteFilter = true;
        }
        if (settings.monitor->filter.preset != defaults.filter.preset)
        {
            yyjson_mut_obj_add_str(doc, filter, "preset", PresetToString(settings.monitor->filter.preset));
            wroteFilter = true;
        }

        if (wroteFilter)
        {
            if (! ensureMonitor())
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, monitor, "filter", filter);
        }

        if (monitor)
        {
            yyjson_mut_obj_add_val(doc, root, "monitor", monitor);
        }
    }

    if (settings.connections)
    {
        const Common::Settings::ConnectionsSettings defaults;
        constexpr std::wstring_view kQuickConnectConnectionId = L"00000000-0000-0000-0000-000000000001";
        const auto isQuickConnect = [&](const Common::Settings::ConnectionProfile& profile) noexcept { return profile.id == kQuickConnectConnectionId; };

        const auto isAwsS3Profile = [&](const Common::Settings::ConnectionProfile& profile) noexcept
        { return profile.pluginId == L"builtin/file-system-s3" || profile.pluginId == L"builtin/file-system-s3table"; };

        const auto isProfilePersistable = [&](const Common::Settings::ConnectionProfile& profile) noexcept
        {
            if (isQuickConnect(profile))
            {
                return false;
            }
            if (profile.id.empty() || profile.name.empty() || profile.pluginId.empty())
            {
                return false;
            }
            if (profile.host.empty() && ! isAwsS3Profile(profile))
            {
                return false;
            }
            return true;
        };

        const auto prunedConnectionExtraForPersist = [&](const Common::Settings::ConnectionProfile& profile) noexcept -> Common::Settings::JsonValue
        {
            Common::Settings::JsonValue pruned;

            const auto* objPtr = std::get_if<Common::Settings::JsonValue::ObjectPtr>(&profile.extra.value);
            if (! objPtr || ! *objPtr)
            {
                pruned.value = std::monostate{};
                return pruned;
            }

            const bool isS3      = profile.pluginId == L"builtin/file-system-s3";
            const bool isS3Table = profile.pluginId == L"builtin/file-system-s3table";
            const bool isAwsS3   = isS3 || isS3Table;
            const bool isImap    = profile.pluginId == L"builtin/file-system-imap";

            auto obj = std::make_shared<Common::Settings::JsonObject>();
            obj->members.reserve((*objPtr)->members.size());

            for (const auto& [k, v] : (*objPtr)->members)
            {
                if (k == "sshPrivateKey" || k == "sshKnownHosts")
                {
                    if (const auto* str = std::get_if<std::string>(&v.value); str && str->empty())
                    {
                        continue;
                    }
                }

                if (isAwsS3)
                {
                    if (k == "endpointOverride")
                    {
                        if (const auto* str = std::get_if<std::string>(&v.value); str && str->empty())
                        {
                            continue;
                        }
                    }

                    if (k == "useHttps" || k == "verifyTls" || (isS3 && k == "useVirtualAddressing"))
                    {
                        if (const auto* b = std::get_if<bool>(&v.value); b && *b)
                        {
                            continue;
                        }
                    }
                }

                if (isImap && k == "ignoreSslTrust")
                {
                    if (const auto* b = std::get_if<bool>(&v.value); b && ! *b)
                    {
                        continue;
                    }
                }

                obj->members.emplace_back(k, v);
            }

            if (obj->members.empty())
            {
                pruned.value = std::monostate{};
                return pruned;
            }

            pruned.value = std::move(obj);
            return pruned;
        };

        const bool hasProfilesToPersist =
            std::any_of(settings.connections->items.begin(),
                        settings.connections->items.end(),
                        [&](const Common::Settings::ConnectionProfile& profile) noexcept { return isProfilePersistable(profile); });

        const bool wroteConnections = hasProfilesToPersist || settings.connections->bypassWindowsHello != defaults.bypassWindowsHello ||
                                      settings.connections->windowsHelloReauthTimeoutMinute != defaults.windowsHelloReauthTimeoutMinute;

        if (wroteConnections)
        {
            yyjson_mut_val* connections = yyjson_mut_obj(doc);
            if (! connections)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, root, "connections", connections);

            if (settings.connections->bypassWindowsHello != defaults.bypassWindowsHello)
            {
                yyjson_mut_obj_add_bool(doc, connections, "bypassWindowsHello", settings.connections->bypassWindowsHello);
            }

            if (settings.connections->windowsHelloReauthTimeoutMinute != defaults.windowsHelloReauthTimeoutMinute)
            {
                yyjson_mut_obj_add_uint(doc, connections, "windowsHelloReauthTimeoutMinute", settings.connections->windowsHelloReauthTimeoutMinute);
            }

            if (hasProfilesToPersist)
            {
                yyjson_mut_val* items = yyjson_mut_arr(doc);
                if (! items)
                {
                    return E_OUTOFMEMORY;
                }
                yyjson_mut_obj_add_val(doc, connections, "items", items);

                const Common::Settings::ConnectionProfile profileDefaults{};

                for (const auto& profile : settings.connections->items)
                {
                    if (! isProfilePersistable(profile))
                    {
                        continue;
                    }

                    yyjson_mut_val* item = yyjson_mut_obj(doc);
                    if (! item)
                    {
                        return E_OUTOFMEMORY;
                    }

                    yyjson_mut_obj_add_val(doc, item, "id", NewString(doc, profile.id));
                    yyjson_mut_obj_add_val(doc, item, "name", NewString(doc, profile.name));
                    yyjson_mut_obj_add_val(doc, item, "pluginId", NewString(doc, profile.pluginId));
                    if (! profile.host.empty())
                    {
                        yyjson_mut_obj_add_val(doc, item, "host", NewString(doc, profile.host));
                    }
                    if (profile.port != 0)
                    {
                        yyjson_mut_obj_add_uint(doc, item, "port", profile.port);
                    }
                    if (! profile.initialPath.empty() && profile.initialPath != profileDefaults.initialPath)
                    {
                        yyjson_mut_obj_add_val(doc, item, "initialPath", NewString(doc, profile.initialPath));
                    }

                    if (! profile.userName.empty())
                    {
                        yyjson_mut_obj_add_val(doc, item, "userName", NewString(doc, profile.userName));
                    }

                    if (profile.authMode != profileDefaults.authMode)
                    {
                        yyjson_mut_obj_add_str(doc, item, "authMode", ConnectionAuthModeToString(profile.authMode));
                    }

                    if (profile.savePassword != profileDefaults.savePassword)
                    {
                        yyjson_mut_obj_add_bool(doc, item, "savePassword", profile.savePassword);
                    }

                    if (profile.requireWindowsHello != profileDefaults.requireWindowsHello)
                    {
                        yyjson_mut_obj_add_bool(doc, item, "requireWindowsHello", profile.requireWindowsHello);
                    }

                    if (! std::holds_alternative<std::monostate>(profile.extra.value))
                    {
                        const Common::Settings::JsonValue prunedExtra = prunedConnectionExtraForPersist(profile);
                        if (! std::holds_alternative<std::monostate>(prunedExtra.value))
                        {
                            HRESULT extraHr       = S_OK;
                            yyjson_mut_val* extra = NewYyjsonFromJsonValue(doc, prunedExtra, extraHr);
                            if (! extra)
                            {
                                return FAILED(extraHr) ? extraHr : E_OUTOFMEMORY;
                            }
                            yyjson_mut_obj_add_val(doc, item, "extra", extra);
                        }
                    }

                    yyjson_mut_arr_add_val(items, item);
                }
            }
        }
    }

    if (settings.fileOperations)
    {
        const Common::Settings::FileOperationsSettings defaults{};
        const bool wroteFileOperations =
            settings.fileOperations->autoDismissSuccess != defaults.autoDismissSuccess ||
            settings.fileOperations->maxDiagnosticsLogFiles != defaults.maxDiagnosticsLogFiles ||
            settings.fileOperations->diagnosticsInfoEnabled != defaults.diagnosticsInfoEnabled ||
            settings.fileOperations->diagnosticsDebugEnabled != defaults.diagnosticsDebugEnabled || settings.fileOperations->maxIssueReportFiles.has_value() ||
            settings.fileOperations->maxDiagnosticsInMemory.has_value() || settings.fileOperations->maxDiagnosticsPerFlush.has_value() ||
            settings.fileOperations->diagnosticsFlushIntervalMs.has_value() || settings.fileOperations->diagnosticsCleanupIntervalMs.has_value();

        if (wroteFileOperations)
        {
            yyjson_mut_val* fileOperations = yyjson_mut_obj(doc);
            if (! fileOperations)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, root, "fileOperations", fileOperations);

            if (settings.fileOperations->autoDismissSuccess != defaults.autoDismissSuccess)
            {
                yyjson_mut_obj_add_bool(doc, fileOperations, "autoDismissSuccess", settings.fileOperations->autoDismissSuccess);
            }

            if (settings.fileOperations->maxDiagnosticsLogFiles != defaults.maxDiagnosticsLogFiles)
            {
                yyjson_mut_obj_add_uint(doc, fileOperations, "maxDiagnosticsLogFiles", settings.fileOperations->maxDiagnosticsLogFiles);
            }

            if (settings.fileOperations->diagnosticsInfoEnabled != defaults.diagnosticsInfoEnabled)
            {
                yyjson_mut_obj_add_bool(doc, fileOperations, "diagnosticsInfoEnabled", settings.fileOperations->diagnosticsInfoEnabled);
            }

            if (settings.fileOperations->diagnosticsDebugEnabled != defaults.diagnosticsDebugEnabled)
            {
                yyjson_mut_obj_add_bool(doc, fileOperations, "diagnosticsDebugEnabled", settings.fileOperations->diagnosticsDebugEnabled);
            }

            if (settings.fileOperations->maxIssueReportFiles.has_value())
            {
                yyjson_mut_obj_add_uint(doc, fileOperations, "maxIssueReportFiles", settings.fileOperations->maxIssueReportFiles.value());
            }

            if (settings.fileOperations->maxDiagnosticsInMemory.has_value())
            {
                yyjson_mut_obj_add_uint(doc, fileOperations, "maxDiagnosticsInMemory", settings.fileOperations->maxDiagnosticsInMemory.value());
            }

            if (settings.fileOperations->maxDiagnosticsPerFlush.has_value())
            {
                yyjson_mut_obj_add_uint(doc, fileOperations, "maxDiagnosticsPerFlush", settings.fileOperations->maxDiagnosticsPerFlush.value());
            }

            if (settings.fileOperations->diagnosticsFlushIntervalMs.has_value())
            {
                yyjson_mut_obj_add_uint(doc, fileOperations, "diagnosticsFlushIntervalMs", settings.fileOperations->diagnosticsFlushIntervalMs.value());
            }

            if (settings.fileOperations->diagnosticsCleanupIntervalMs.has_value())
            {
                yyjson_mut_obj_add_uint(doc, fileOperations, "diagnosticsCleanupIntervalMs", settings.fileOperations->diagnosticsCleanupIntervalMs.value());
            }
        }
    }

    if (settings.compareDirectories)
    {
        const Common::Settings::CompareDirectoriesSettings defaults{};
        const auto& compare     = settings.compareDirectories.value();
        const bool wroteCompare = compare.compareSize != defaults.compareSize || compare.compareDateTime != defaults.compareDateTime ||
                                  compare.compareAttributes != defaults.compareAttributes || compare.compareContent != defaults.compareContent ||
                                  compare.compareSubdirectories != defaults.compareSubdirectories ||
                                  compare.compareSubdirectoryAttributes != defaults.compareSubdirectoryAttributes ||
                                  compare.selectSubdirsOnlyInOnePane != defaults.selectSubdirsOnlyInOnePane || compare.ignoreFiles != defaults.ignoreFiles ||
                                  compare.ignoreDirectories != defaults.ignoreDirectories || compare.showIdenticalItems != defaults.showIdenticalItems ||
                                  ! compare.ignoreFilesPatterns.empty() || ! compare.ignoreDirectoriesPatterns.empty();

        if (wroteCompare)
        {
            yyjson_mut_val* compareObj = yyjson_mut_obj(doc);
            if (! compareObj)
            {
                return E_OUTOFMEMORY;
            }
            yyjson_mut_obj_add_val(doc, root, "compareDirectories", compareObj);

            if (compare.compareSize != defaults.compareSize)
            {
                yyjson_mut_obj_add_bool(doc, compareObj, "compareSize", compare.compareSize);
            }
            if (compare.compareDateTime != defaults.compareDateTime)
            {
                yyjson_mut_obj_add_bool(doc, compareObj, "compareDateTime", compare.compareDateTime);
            }
            if (compare.compareAttributes != defaults.compareAttributes)
            {
                yyjson_mut_obj_add_bool(doc, compareObj, "compareAttributes", compare.compareAttributes);
            }
            if (compare.compareContent != defaults.compareContent)
            {
                yyjson_mut_obj_add_bool(doc, compareObj, "compareContent", compare.compareContent);
            }
            if (compare.compareSubdirectories != defaults.compareSubdirectories)
            {
                yyjson_mut_obj_add_bool(doc, compareObj, "compareSubdirectories", compare.compareSubdirectories);
            }
            if (compare.compareSubdirectoryAttributes != defaults.compareSubdirectoryAttributes)
            {
                yyjson_mut_obj_add_bool(doc, compareObj, "compareSubdirectoryAttributes", compare.compareSubdirectoryAttributes);
            }
            if (compare.selectSubdirsOnlyInOnePane != defaults.selectSubdirsOnlyInOnePane)
            {
                yyjson_mut_obj_add_bool(doc, compareObj, "selectSubdirsOnlyInOnePane", compare.selectSubdirsOnlyInOnePane);
            }
            if (compare.ignoreFiles != defaults.ignoreFiles)
            {
                yyjson_mut_obj_add_bool(doc, compareObj, "ignoreFiles", compare.ignoreFiles);
            }
            if (! compare.ignoreFilesPatterns.empty())
            {
                yyjson_mut_obj_add_val(doc, compareObj, "ignoreFilesPatterns", NewString(doc, compare.ignoreFilesPatterns));
            }
            if (compare.ignoreDirectories != defaults.ignoreDirectories)
            {
                yyjson_mut_obj_add_bool(doc, compareObj, "ignoreDirectories", compare.ignoreDirectories);
            }
            if (! compare.ignoreDirectoriesPatterns.empty())
            {
                yyjson_mut_obj_add_val(doc, compareObj, "ignoreDirectoriesPatterns", NewString(doc, compare.ignoreDirectoriesPatterns));
            }
            if (compare.showIdenticalItems != defaults.showIdenticalItems)
            {
                yyjson_mut_obj_add_bool(doc, compareObj, "showIdenticalItems", compare.showIdenticalItems);
            }
        }
    }

    yyjson_write_err writeErr{};
    size_t jsonLen                = 0;
    const yyjson_write_flag flags = YYJSON_WRITE_PRETTY_TWO_SPACES | YYJSON_WRITE_NEWLINE_AT_END;
    char* json                    = yyjson_mut_write_opts(doc, flags, nullptr, &jsonLen, &writeErr);
    if (! json)
    {
        Debug::Error(L"Failed to serialize settings to JSON: code: {}", writeErr.code);
        return E_FAIL;
    }

    auto freeJson = wil::scope_exit([&] { free(json); });

    std::string output;
    output.reserve(3u + jsonLen);
    output.push_back(static_cast<char>(0xEF));
    output.push_back(static_cast<char>(0xBB));
    output.push_back(static_cast<char>(0xBF));
    output.append(json, jsonLen);

    const HRESULT writeHr = WriteFileBytesAtomic(settingsPath, output);
    if (FAILED(writeHr))
    {
        return writeHr;
    }

    if (const std::string_view baseSchema = GetSettingsStoreSchemaJsonUtf8(); ! baseSchema.empty())
    {
        const HRESULT schemaHr = SaveSettingsSchema(appId, baseSchema);
        if (FAILED(schemaHr))
        {
            Debug::Warning(L"Failed to write settings schema file (hr=0x{:08X}) for appId={}", static_cast<unsigned long>(schemaHr), appId);
        }
    }

    return S_OK;
}

HRESULT SaveSettingsSchema(std::wstring_view appId, std::string_view schemaJsonUtf8) noexcept
{
    const std::filesystem::path schemaPath = GetSettingsSchemaPath(appId);
    if (schemaPath.empty())
    {
        return E_FAIL;
    }

    return WriteFileBytesAtomic(schemaPath, schemaJsonUtf8);
}

HRESULT ParseJsonValue(std::string_view jsonText, JsonValue& out) noexcept
{
    out.value = std::monostate{};

    if (jsonText.empty())
    {
        return S_OK;
    }

    std::string mutableJson;
    mutableJson.assign(jsonText.data(), jsonText.size());

    yyjson_read_err err{};
    yyjson_doc* doc = yyjson_read_opts(mutableJson.data(), mutableJson.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
    if (! doc)
    {
        if (err.code == YYJSON_READ_ERROR_MEMORY_ALLOCATION)
        {
            return E_OUTOFMEMORY;
        }
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    return ConvertYyjsonToJsonValue(root, out);
}

HRESULT SerializeJsonValue(const JsonValue& value, std::string& outJsonText) noexcept
{
    outJsonText.clear();

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    HRESULT rootHr       = S_OK;
    yyjson_mut_val* root = NewYyjsonFromJsonValue(doc, value, rootHr);
    if (! root)
    {
        return FAILED(rootHr) ? rootHr : E_OUTOFMEMORY;
    }
    yyjson_mut_doc_set_root(doc, root);

    yyjson_write_err writeErr{};
    size_t jsonLen = 0;
    char* json     = yyjson_mut_write_opts(doc, YYJSON_WRITE_NOFLAG, nullptr, &jsonLen, &writeErr);
    if (! json)
    {
        if (writeErr.code == YYJSON_WRITE_ERROR_MEMORY_ALLOCATION)
        {
            return E_OUTOFMEMORY;
        }
        return E_FAIL;
    }

    auto freeJson = wil::scope_exit([&] { free(json); });

    outJsonText.assign(json, jsonLen);
    return S_OK;
}

HRESULT LoadThemeDefinitionsFromDirectory(const std::filesystem::path& directory, std::vector<ThemeDefinition>& out) noexcept
{
    out.clear();

    if (directory.empty())
    {
        Debug::Error(L"Themes folder is empty {}", directory.c_str());
        return S_FALSE;
    }

    std::error_code ec;
    const bool exists = std::filesystem::exists(directory, ec);
    if (ec || ! exists)
    {
        Debug::Error(L"Themes folder does not exist {}", directory.c_str());
        return S_FALSE;
    }

    const bool isDir = std::filesystem::is_directory(directory, ec);
    if (ec || ! isDir)
    {
        Debug::Error(L"Themes folder is not a directory {}", directory.c_str());
        return S_FALSE;
    }

    std::vector<std::filesystem::path> paths;
    std::filesystem::directory_iterator it(directory, ec);
    if (ec)
    {
        Debug::Error(L"Failed to iterate themes folder {}", directory.c_str());
        return S_FALSE;
    }

    const auto isThemeFile = [](const std::filesystem::path& path) noexcept
    {
        const std::wstring fileName        = path.filename().wstring();
        constexpr std::wstring_view suffix = L".theme.json5";
        if (fileName.size() < suffix.size())
        {
            Debug::Error(L"Invalid theme file name {}", fileName.c_str());
            return false;
        }

        const size_t offset = fileName.size() - suffix.size();
        for (size_t index = 0; index < suffix.size(); ++index)
        {
            wchar_t fileChar   = fileName[offset + index];
            wchar_t suffixChar = suffix[index];

            if (fileChar >= L'A' && fileChar <= L'Z')
            {
                fileChar = static_cast<wchar_t>((fileChar - L'A') + L'a');
            }
            if (suffixChar >= L'A' && suffixChar <= L'Z')
            {
                suffixChar = static_cast<wchar_t>((suffixChar - L'A') + L'a');
            }

            if (fileChar != suffixChar)
            {
                return false;
            }
        }

        return true;
    };

    for (std::filesystem::directory_iterator end; it != end; it.increment(ec))
    {
        if (ec)
        {
            break;
        }

        const std::filesystem::directory_entry& entry = *it;

        const bool isFile = entry.is_regular_file(ec);
        if (ec || ! isFile)
        {
            ec.clear();
            continue;
        }

        if (! isThemeFile(entry.path()))
        {
            continue;
        }

        paths.push_back(entry.path());
    }

    if (paths.empty())
    {
        return S_FALSE;
    }

    std::sort(paths.begin(),
              paths.end(),
              [](const std::filesystem::path& left, const std::filesystem::path& right) noexcept
              {
                  const std::wstring leftText  = left.filename().wstring();
                  const std::wstring rightText = right.filename().wstring();
                  return CompareStringOrdinal(leftText.c_str(), -1, rightText.c_str(), -1, TRUE) == CSTR_LESS_THAN;
              });

    for (const auto& path : paths)
    {
        std::string bytes;
        const HRESULT readHr = ReadFileBytes(path, bytes);
        if (FAILED(readHr))
        {
            Debug::Error(L"Failed to read theme file {}", path.c_str());
            continue;
        }

        yyjson_read_err err{};
        yyjson_doc* doc = yyjson_read_opts(bytes.data(), bytes.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
        if (! doc)
        {
            LogJsonParseError(L"theme file", path, err);
            continue;
        }

        auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

        yyjson_val* root = yyjson_doc_get_root(doc);
        if (! root || ! yyjson_is_obj(root))
        {
            Debug::Error(L"Failed to get root of theme file {}", path.c_str());
            continue;
        }

        const auto idText   = GetString(root, "id");
        const auto nameText = GetString(root, "name");
        const auto baseText = GetString(root, "baseThemeId");
        yyjson_val* colors  = GetObj(root, "colors");
        if (! idText || ! nameText || ! baseText || ! colors)
        {
            Debug::Error(L"Failed to get theme properties from file {}", path.c_str());
            continue;
        }

        ThemeDefinition def;
        def.id          = Utf16FromUtf8(*idText);
        def.name        = Utf16FromUtf8(*nameText);
        def.baseThemeId = Utf16FromUtf8(*baseText);
        if (def.id.empty() || def.name.empty() || def.baseThemeId.empty())
        {
            Debug::Error(L"Invalid theme properties in file {}", path.c_str());
            continue;
        }

        const auto duplicate = std::find_if(out.begin(), out.end(), [&](const ThemeDefinition& existing) { return existing.id == def.id; });
        if (duplicate != out.end())
        {
            Debug::Error(L"Duplicate theme ID '{}' in file {}", def.id.c_str(), path.c_str());
            continue;
        }

        yyjson_val* colorKey = nullptr;
        yyjson_val* colorVal = nullptr;
        yyjson_obj_iter iter = yyjson_obj_iter_with(colors);
        while ((colorKey = yyjson_obj_iter_next(&iter)))
        {
            colorVal = yyjson_obj_iter_get_val(colorKey);
            if (! colorVal || ! yyjson_is_str(colorKey) || ! yyjson_is_str(colorVal))
            {
                Debug::Error(L"Invalid color entry in theme file {}", path.c_str());
                continue;
            }

            const char* keyStr = yyjson_get_str(colorKey);
            const char* valStr = yyjson_get_str(colorVal);
            if (! keyStr || ! valStr)
            {
                Debug::Error(L"Failed to get color entry in theme file {}", path.c_str());
                continue;
            }

            uint32_t argb = 0;
            if (! TryParseColorUtf8(valStr, argb))
            {
                Debug::Error(L"Failed to parse color value '{}' in theme file {}", Utf16FromUtf8(valStr).c_str(), path.c_str());
                continue;
            }

            const std::wstring keyWide = Utf16FromUtf8(keyStr);
            if (keyWide.empty())
            {
                Debug::Error(L"Invalid color key in theme file {}", path.c_str());
                continue;
            }

            def.colors[keyWide] = argb;
        }

        out.push_back(std::move(def));
    }

    return out.empty() ? S_FALSE : S_OK;
}

WindowPlacement NormalizeWindowPlacement(const WindowPlacement& saved, unsigned int currentDpi) noexcept
{
    WindowPlacement result = saved;

    int width  = result.bounds.width;
    int height = result.bounds.height;

    if (width < 1)
    {
        width = 1;
    }
    if (height < 1)
    {
        height = 1;
    }

    if (result.dpi && *result.dpi > 0 && currentDpi > 0 && *result.dpi != currentDpi)
    {
        const double scale = static_cast<double>(currentDpi) / static_cast<double>(*result.dpi);
        width              = std::max(1, static_cast<int>(std::lround(static_cast<double>(width) * scale)));
        height             = std::max(1, static_cast<int>(std::lround(static_cast<double>(height) * scale)));
    }

    RECT desired{};
    desired.left   = result.bounds.x;
    desired.top    = result.bounds.y;
    desired.right  = desired.left + width;
    desired.bottom = desired.top + height;

    struct WorkArea
    {
        RECT work{};
        bool primary = false;
    };

    std::vector<WorkArea> workAreas;

    const auto enumProc = [](HMONITOR hMonitor, [[maybe_unused]] HDC, [[maybe_unused]] LPRECT, LPARAM param) noexcept -> BOOL
    {
        auto& areas = *reinterpret_cast<std::vector<WorkArea>*>(param);
        MONITORINFOEXW mi{};
        mi.cbSize = sizeof(mi);
        if (GetMonitorInfoW(hMonitor, &mi))
        {
            WorkArea area{};
            area.work    = mi.rcWork;
            area.primary = (mi.dwFlags & MONITORINFOF_PRIMARY) != 0;
            areas.push_back(area);
        }
        return TRUE;
    };

    EnumDisplayMonitors(nullptr, nullptr, enumProc, reinterpret_cast<LPARAM>(&workAreas));
    if (workAreas.empty())
    {
        result.bounds.width  = width;
        result.bounds.height = height;
        return result;
    }

    auto contains = [](const RECT& outer, const RECT& inner) noexcept
    { return inner.left >= outer.left && inner.top >= outer.top && inner.right <= outer.right && inner.bottom <= outer.bottom; };

    for (const auto& area : workAreas)
    {
        if (contains(area.work, desired))
        {
            result.bounds.width  = width;
            result.bounds.height = height;
            return result;
        }
    }

    size_t bestIndex     = 0;
    uint64_t bestArea    = 0;
    bool anyIntersection = false;

    for (size_t i = 0; i < workAreas.size(); ++i)
    {
        RECT inter{};
        if (! IntersectRect(&inter, &desired, &workAreas[i].work))
        {
            continue;
        }

        anyIntersection     = true;
        const uint64_t w    = static_cast<uint64_t>(std::max(0L, inter.right - inter.left));
        const uint64_t h    = static_cast<uint64_t>(std::max(0L, inter.bottom - inter.top));
        const uint64_t area = w * h;
        if (area > bestArea)
        {
            bestArea  = area;
            bestIndex = i;
        }
    }

    if (! anyIntersection)
    {
        for (size_t i = 0; i < workAreas.size(); ++i)
        {
            if (workAreas[i].primary)
            {
                bestIndex = i;
                break;
            }
        }
    }

    const RECT work      = workAreas[bestIndex].work;
    const int workWidth  = std::max(1L, work.right - work.left);
    const int workHeight = std::max(1L, work.bottom - work.top);

    width  = std::clamp(width, 1, workWidth);
    height = std::clamp(height, 1, workHeight);

    const LONG maxX = work.right - width;
    const LONG maxY = work.bottom - height;

    const LONG x = std::clamp(desired.left, work.left, maxX);
    const LONG y = std::clamp(desired.top, work.top, maxY);

    result.bounds.x      = x;
    result.bounds.y      = y;
    result.bounds.width  = width;
    result.bounds.height = height;

    return result;
}

} // namespace Common::Settings
